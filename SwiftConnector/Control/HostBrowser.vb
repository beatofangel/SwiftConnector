' The MIT License (MIT)
' Copyright © 2013-2021 Eric Wang <beatofangel@gmx.com>
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the “Software”), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

Imports System.ComponentModel
Imports System.Configuration
Imports System.Data
Imports System.Data.Common
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Web
Imports System.Windows.Forms
Imports log4net
Imports Microsoft.Web.WebView2.Core
Imports Newtonsoft
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' 嵌入式浏览器，注意需要安装edge runtime https://developer.microsoft.com/en-us/microsoft-edge/webview2/
''' </summary>
Public Class HostBrowser

    Protected Shared logger As ILog = LogManager.GetLogger(Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private dsService As New DatasourceService

    Private textService As New TextService

    Private dapperService As New DapperService

    Private Const DEFAULT_WEBVIEW2_USERDATAFOLDER As String = "SwiftConnector\temp\"
    'Private Const HOST As String = "dbtoolsaddin.local"
    'Private Const HOST As String = "localhost"
    'Private Const PORT As Integer = 8081
    Private Property HB_HOST As String
    Private Property HB_PORT As Integer
    Private Property HB_SCHEME As String
    Private Property HB_PATH As String
    Private Property WebView2DevToolsEnabled As Boolean?
    Private _fragment As String = ""

    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    <Bindable(False)>
    <Browsable(False)>
    Public Property Fragment As String
        Get
            If IsDesignMode() Then
                Return ""
            Else
                Return String.Format("{0}/{1}", Globals.ThisAddIn.StrLangCode, _fragment)
            End If
        End Get
        Set(value As String)
            _fragment = value
            If Not IsDesignMode() Then
                Dim unused = NavigateTo(String.Format("{0}/{1}", Globals.ThisAddIn.StrLangCode, _fragment))
            End If
        End Set
    End Property

    Private Async Function Init(Optional fragment As String = "") As Threading.Tasks.Task
        If String.IsNullOrEmpty(HB_HOST) Then
            HB_HOST = ConfigurationManager.AppSettings.Get("Host")
            HB_PORT = Integer.Parse(ConfigurationManager.AppSettings.Get("Port"))
            HB_SCHEME = ConfigurationManager.AppSettings.Get("Scheme")
            HB_PATH = ConfigurationManager.AppSettings.Get("Path")
            logger.Debug(String.Format("initialize host({0}) and port({1})", HB_HOST, HB_PORT))
        End If

        Dim virtualHost As Boolean = Boolean.Parse(ConfigurationManager.AppSettings.Get("VirtualHost"))
        If virtualHost Then
            HB_HOST = "swiftconnector"
        End If

        If InnerBrowser.CoreWebView2 Is Nothing Then
            logger.Debug("initialize webview2")
            Dim env As Object
            Dim userDataFolder As String
            Dim webView2UserDataFolder As String = ConfigurationManager.AppSettings.Get("WebView2UserDataFolder")
            If String.IsNullOrEmpty(webView2UserDataFolder) Then
                userDataFolder = Path.Combine(Environ("UserProfile"), DEFAULT_WEBVIEW2_USERDATAFOLDER)
            Else
                userDataFolder = Environment.ExpandEnvironmentVariables(webView2UserDataFolder)
            End If

            Try
                env = Await CoreWebView2Environment.CreateAsync(Nothing, userDataFolder)
                Await InnerBrowser.EnsureCoreWebView2Async(env)
            Catch ex As Exception
                logger.Error(ex)
            End Try

            'Dim assemblyInfo As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
            ''Dim location As String = assemblyInfo.Location
            'Dim uriCodeBase As Uri = New Uri(assemblyInfo.CodeBase)
            'Dim basePath As String = Path.GetDirectoryName(uriCodeBase.LocalPath.ToString())
            ' 通过配置项判断是否为开发模式，开发模式：访问localhost:port，发布模式：访问虚拟主机映射
            If virtualHost Then
                InnerBrowser.CoreWebView2.SetVirtualHostNameToFolderMapping(HB_HOST, Path.Combine(GetBasePath, "Local"), CoreWebView2HostResourceAccessKind.Allow)
                logger.Debug(String.Format("enable virtual host with host({0}), folderPath({1})", HB_HOST, GetBasePath))
            End If
            AddHandler InnerBrowser.WebMessageReceived, AddressOf WebMessageReceived
        End If

        Dim uriBuilder = New UriBuilder With {
            .Scheme = HB_SCHEME,
            .Host = HB_HOST,
            .Path = HB_PATH,
            .Fragment = fragment
        }
        If Not virtualHost Then
            uriBuilder.Port = HB_PORT '仅当virtualHost = false时有效
        End If
        'InnerBrowser.CoreWebView2.Navigate("https://www.baidu.com")
        InnerBrowser.CoreWebView2.Settings.IsStatusBarEnabled = False
        InnerBrowser.Source = uriBuilder.Uri

        If WebView2DevToolsEnabled Is Nothing Then
            WebView2DevToolsEnabled = Boolean.Parse(ConfigurationManager.AppSettings.Get("WebView2DevToolsEnabled"))
            logger.Debug("enable devtools")
        End If
        If WebView2DevToolsEnabled Then
            InnerBrowser.CoreWebView2.OpenDevToolsWindow()
        End If
    End Function

    Private Sub WebMessageReceived(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
        'Dim jsonObject = JObject.Parse(e.WebMessageAsJson)
        logger.Debug(String.Format("received request({0})", e.WebMessageAsJson))
        Dim jsonString = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(e.WebMessageAsJson)
        Dim api = jsonString("api")
        Dim args = If(jsonString.ContainsKey("args"), jsonString("args"), Nothing)
        Dim cb = If(jsonString.ContainsKey("callback"), jsonString("callback"), Nothing)
        Select Case api
            Case "platformVerify"
                DoResponse(api, cb, Function() JsonConvert.SerializeObject(New Response(True, api)))
            Case "loadChangeLog"
                DoResponse(api, cb, Function() JsonConvert.SerializeObject(New Response(True, api, data:=File.ReadAllText(Path.Combine(GetBasePath, "CHANGELOG.MD")))))
            Case "loadConnections"
                DoResponse(api, cb, Function() JsonConvert.SerializeObject(New Response(True, api, data:=dsService.FindAllDataSource())))
            Case "addConnection"
                DoResponse(api, cb, Function()
                                        Dim jsonObj = JsonConvert.DeserializeObject(Of JObject)(args)
                                        Dim ds = New DataSource With {
                                        .Id = Guid.NewGuid.ToString.Replace("-", ""),
                                        .Type = jsonObj.GetValue("databaseType").ToString,
                                        .Name = jsonObj.GetValue("connectionName").ToString,
                                        .Datasource = jsonObj.GetValue("host").ToString,
                                        .Port = If(IsNull(jsonObj.GetValue("port")), Nothing, jsonObj.GetValue("port").ToString),
                                        .Database = If(IsNull(jsonObj.GetValue("databaseName")), Nothing, jsonObj.GetValue("databaseName").ToString),
                                        .Username = If(IsNull(jsonObj.GetValue("username")), Nothing, jsonObj.GetValue("username").ToString),
                                        .Password = If(IsNull(jsonObj.GetValue("password")), Nothing, jsonObj.GetValue("password").ToString),
                                        .Mode = If(IsNull(jsonObj.GetValue("protected")), 0, If("1".Equals(jsonObj.GetValue("protected").ToString), 1, 0)),
                                        .Current = False,
                                        .Lastchange = Date.Now
                                    }
                                        dsService.AddDataSource(ds)

                                        Return JsonConvert.SerializeObject(New Response(True, api))
                                    End Function)
            Case "editConnection"
                DoResponse(api, cb, Function()
                                        Dim jsonObj = JsonConvert.DeserializeObject(Of JObject)(args)
                                        Dim ds = New DataSource With {
                                        .Id = jsonObj.GetValue("id").ToString,
                                        .Type = jsonObj.GetValue("databaseType").ToString,
                                        .Name = jsonObj.GetValue("connectionName").ToString,
                                        .Datasource = jsonObj.GetValue("host").ToString,
                                        .Port = jsonObj.GetValue("port").ToString,
                                        .Database = jsonObj.GetValue("databaseName").ToString,
                                        .Username = jsonObj.GetValue("username").ToString,
                                        .Password = jsonObj.GetValue("password").ToString,
                                        .Mode = If("1".Equals(jsonObj.GetValue("protected").ToString), 1, 0),
                                        .Current = jsonObj.GetValue("current").ToString,
                                        .Lastchange = Date.Now
                                    }
                                        dsService.EditDataSource(ds)

                                        Return JsonConvert.SerializeObject(New Response(True, api))
                                    End Function)
            Case "deleteConnection"
                DoResponse(api, cb, Function()
                                        Dim jsonObj = JsonConvert.DeserializeObject(Of JObject)(args)
                                        Dim ds = New DataSource With {
                                       .Id = jsonObj.GetValue("id").ToString
                                   }
                                        dsService.DeleteDataSource(ds)

                                        Return JsonConvert.SerializeObject(New Response(True, api))
                                    End Function)
            Case "testConnection"
                DoResponse(api, cb, Function()
                                        Dim jsonObj = JsonConvert.DeserializeObject(Of JObject)(args)
                                        Dim ds = New DataSource With {
                                            .Type = jsonObj.GetValue("databaseType").ToString,
                                            .Datasource = jsonObj.GetValue("host").ToString,
                                            .Port = If(IsNull(jsonObj.GetValue("port")), Nothing, jsonObj.GetValue("port").ToString),
                                            .Database = If(IsNull(jsonObj.GetValue("databaseName")), Nothing, jsonObj.GetValue("databaseName").ToString),
                                            .Username = If(IsNull(jsonObj.GetValue("username")), Nothing, jsonObj.GetValue("username").ToString),
                                            .Password = If(IsNull(jsonObj.GetValue("password")), Nothing, jsonObj.GetValue("password").ToString)
                                        }

                                        TestConnection(ds)

                                        Return JsonConvert.SerializeObject(New Response(True, api))
                                    End Function)
            Case "switch2Current"
                DoResponse(api, cb, Function()
                                        Dim jsonObj = JsonConvert.DeserializeObject(Of JObject)(args)
                                        If dsService.SwitchDataSourceTo(jsonObj.GetValue("Id").ToString) Then
                                            'Dim title = textService.GetTextByProperty(TextType.TT_MSG_SWITCH_SUCCESS)
                                            'Dim content = textService.GetTextByProperty(TextType.TT_MSG_CONNECTION_IN_USE).Replace("{0}", jsonObj.GetValue("Name").ToString)
                                            'Dim logo = "Resources/Icon/" & DataSourceDic(jsonObj.GetValue("Type").ToObject(Of DataSourceType)) & "_large_64.png"
                                            'Toast(title, content, logo)
                                            'Toast(title, content)
                                            Return JsonConvert.SerializeObject(New Response(True, api))
                                        Else
                                            Return JsonConvert.SerializeObject(New Response(False, api, message:="Switch failed"))
                                        End If
                                    End Function)
            Case "curConnection"
                DoResponse(api, cb, Function()
                                        Return JsonConvert.SerializeObject(New Response(True, api, data:=Globals.ThisAddIn.CurDataSource.Type))
                                    End Function)
            Case "sqlQuery"
                DoResponse(api, cb, Function()
                                        Dim jsonObj = JsonConvert.DeserializeObject(Of JObject)(args)
                                        Dim sql = jsonObj.GetValue("sql").ToString()
                                        Dim data As New JObject
                                        Dim items As New JArray
                                        Dim headers As New JArray
                                        Dim forward = Boolean.Parse(jsonObj.GetValue("forward").ToString)
                                        Dim last = Integer.Parse(jsonObj.GetValue("more").ToString)
                                        Dim start = If(forward, last, last - 50)
                                        Dim more As Integer
                                        Dim maxLength = Globals.ThisAddIn.MyRibbon.RecordLimit
                                        Dim tx = dapperService.ExecuteReader(Sub(reader As DbDataReader)

                                                                                 For i = 0 To reader.FieldCount - 1
                                                                                     Dim header As New JObject
                                                                                     header.Add(New JProperty("text", reader.GetName(i)))
                                                                                     header.Add(New JProperty("value", reader.GetName(i)))
                                                                                     headers.Add(header)
                                                                                 Next

                                                                                 Dim idx = 0
                                                                                 While reader.Read
                                                                                     If idx < start Then
                                                                                         idx += 1
                                                                                     Else
                                                                                         If items.Count >= maxLength Then
                                                                                             more = start + maxLength
                                                                                             Exit While
                                                                                         End If
                                                                                         Dim item As New JObject
                                                                                         For i = 0 To reader.FieldCount - 1
                                                                                             item.Add(New JProperty(reader.GetName(i), reader.GetValue(i)))
                                                                                         Next
                                                                                         items.Add(item)
                                                                                     End If
                                                                                 End While
                                                                             End Sub, Integer.Parse(jsonObj.GetValue("tx")), sql:=sql)
                                        data.Add("tx", tx)
                                        data.Add("more", more)
                                        data.Add("forward", forward)
                                        data.Add("items", items)
                                        data.Add("headers", headers)
                                        Return JsonConvert.SerializeObject(New Response(True, api, data:=data))
                                    End Function)

            Case "closeWindow"
                DoResponse(api, cb, Function()
                                        ParentForm.WindowState = FormWindowState.Normal
                                        ParentForm.Hide()

                                        Return JsonConvert.SerializeObject(New Response(True, api))
                                    End Function)

            Case "minimizeWindow"
                DoResponse(api, cb, Function()
                                        ParentForm.WindowState = FormWindowState.Minimized

                                        Return JsonConvert.SerializeObject(New Response(True, api))
                                    End Function)
            Case "maximizeWindow"
                DoResponse(api, cb, Function()
                                        ParentForm.WindowState = FormWindowState.Maximized

                                        Return JsonConvert.SerializeObject(New Response(True, api))
                                    End Function)
            Case "restoreWindow"
                DoResponse(api, cb, Function()
                                        ParentForm.WindowState = FormWindowState.Normal

                                        Return JsonConvert.SerializeObject(New Response(True, api))
                                    End Function)
            Case Else
                Diagnostics.Debug.Print("unimplemented")
        End Select
        'InnerBrowser.CoreWebView2.ExecuteScriptAsync("response('hello');")

    End Sub

    Private Sub DoResponse(callback As String, args As String)
        logger.Debug(String.Format("return response({0})", callback))
        InnerBrowser.CoreWebView2.ExecuteScriptAsync(String.Format("{0}('{1}')", callback, HttpUtility.JavaScriptStringEncode(args)))
    End Sub

    Private Sub DoResponse(api As String, callback As String, fx As Func(Of String))
        Try
            If String.IsNullOrEmpty(callback) Then
                fx.Invoke()
            Else
                DoResponse(callback, fx.Invoke())
            End If
        Catch ex As Exception
            DoResponse(callback, JsonConvert.SerializeObject(New Response(False, api, message:=ex.ToString)))
        End Try
    End Sub

    'Private _basePath As String
    'Private ReadOnly Property GetBasePath() As String
    '    Get
    '        If String.IsNullOrEmpty(_basePath) Then
    '            Dim assemblyInfo As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
    '            'Dim location As String = assemblyInfo.Location
    '            Dim uriCodeBase As Uri = New Uri(assemblyInfo.CodeBase)
    '            _basePath = Path.GetDirectoryName(uriCodeBase.LocalPath.ToString())
    '        End If
    '        Return _basePath
    '    End Get
    'End Property

    Public Async Function SwitchLanguage() As Threading.Tasks.Task
        Dim targetFragment As String = InnerBrowser.Source.Fragment
        targetFragment = Regex.Replace(targetFragment, "^#/\w+/", String.Format("{0}/", Globals.ThisAddIn.StrLangCode), RegexOptions.IgnoreCase)
        Await NavigateTo(targetFragment)
    End Function

    Private Async Function NavigateTo(Optional fragment As String = Nothing) As Threading.Tasks.Task
        Dim path As String = If(fragment Is Nothing, Me.Fragment, fragment)
        Await Init(path)
        Await InnerBrowser.CoreWebView2.ExecuteScriptAsync(String.Format("navigateTo('/{0}')", HttpUtility.JavaScriptStringEncode(path)))
    End Function

    Public Sub HostBrowser_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
        If Visible And InnerBrowser.CoreWebView2 IsNot Nothing Then
            Dim api = "initPage"
            Dim callback = GetInitMethod()
            DoResponse(api, callback, Function()
                                          Return JsonConvert.SerializeObject(New Response(True, api, data:=GetInitParams(callback)))
                                      End Function)
        End If
    End Sub

    ''' <summary>
    ''' 首次打开webview2来不及初始化，无法接收到参数 TODO 删除
    ''' </summary>
    ''' <param name="callback"></param>
    ''' <returns></returns>
    Private Function GetInitParams(callback) As String
        Select Case callback
            Case "initConnections"
                Return Globals.ThisAddIn.CurDataSource.Id
            Case "initQueryEditor"
                Return Globals.ThisAddIn.CurDataSource.Type
            Case Else
                Return ""
        End Select
    End Function

    Private Function GetInitMethod() As String
        Dim m = "init"
        Array.ForEach(_fragment.Split("-"), Sub(e)
                                                m = m & e.Substring(0, 1).ToUpper & e.Substring(1)
                                            End Sub)
        Return m
    End Function
End Class
