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

Imports System.Data.Common
Imports System.Diagnostics
Imports System.Drawing
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports System.Web
Imports System.Windows.Forms
Imports log4net
Imports Microsoft.Toolkit.Uwp.Notifications
Imports SwiftConnector.My.Resources

<Runtime.InteropServices.ComVisible(True)>
Public Class Ribbon
    Implements Office.IRibbonExtensibility

    Protected Shared logger As ILog = LogManager.GetLogger(Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private ribbon As Office.IRibbonUI
    Private datasources As New List(Of DataSource)
    'Private curDsType As DataSourceType = DataSourceType.Unknown

    Protected Shared ReadOnly Property StyleService As StyleService = New StyleService()

    Protected Shared ReadOnly Property ConfigService As ConfigService = New ConfigService()

    Protected Shared ReadOnly Property TextService As TextService = New TextService()

    Protected Shared ReadOnly Property DatasourceService As DatasourceService = New DatasourceService()

    Private _delMode As DeleteMode?
    Public Property DelMode As DeleteMode
        Get
            If _delMode Is Nothing Then
                _delMode = ConfigService.GetDelMode()
            End If
            Return _delMode
        End Get
        Set(value As DeleteMode)
            Try
                ConfigService.Change(New Config With {
                                           .Prop = "DeleteMode",
                                           .Locale = "-", ' no locale
                                           .Val = value,
                                           .Lastchange = Date.Now
                })
                _delMode = value
            Catch ex As Exception

            End Try
        End Set
    End Property

    Private _execMode As ExecuteMode?
    Public Property ExecMode As ExecuteMode
        Get
            If _execMode Is Nothing Then
                _execMode = ConfigService.GetExecMode()
            End If
            Return _execMode
        End Get
        Set(value As ExecuteMode)
            Try
                ConfigService.Change(New Config With {
                                           .Prop = "ExecuteMode",
                                           .Locale = "-", ' no locale
                                           .Val = value,
                                           .Lastchange = Date.Now
                })
                _execMode = value
            Catch ex As Exception

            End Try
        End Set
    End Property

    Public Property ProtectedMode As Boolean
        Get
            Return If(CurDataSource Is Nothing, False, CurDataSource.Mode = OperateMode.Protected)
        End Get
        Set(value As Boolean)
            Dim mode = If(value, OperateMode.Protected, OperateMode.Normal)
            Try
                DatasourceService.EditDataSource(New DataSource With {
                    .Id = CurDataSource.Id,
                    .Datasource = CurDataSource.Datasource,
                    .Database = CurDataSource.Database,
                    .Name = CurDataSource.Name,
                    .Username = CurDataSource.Username,
                    .Password = CurDataSource.Password,
                    .Type = CurDataSource.Type,
                    .Mode = mode,
                    .Port = CurDataSource.Port,
                    .Lastchange = Date.Now
                                                 })
                CurDataSource.Mode = mode
            Catch ex As Exception

            End Try
        End Set
    End Property

    Private _showProps As Boolean?

    Public Property ShowProps As Boolean
        Get
            If _showProps Is Nothing Then
                _showProps = ConfigService.IsShowColProps()
            End If
            Return _showProps
        End Get
        Set(value As Boolean)
            Try
                ConfigService.Change(New Config With {
                                           .Prop = "ShowColumnProperties",
                                           .Locale = "-", ' no locale
                                           .Val = value,
                                           .Lastchange = Date.Now
                })
                _showProps = value
            Catch ex As Exception

            End Try
        End Set
    End Property

    Private _autoFitColumns As Boolean?

    Public Property AutoFitColumns As Boolean
        Get
            If _autoFitColumns Is Nothing Then
                _autoFitColumns = ConfigService.IsAutoFitColumns
            End If
            Return _autoFitColumns
        End Get
        Set(value As Boolean)
            Try
                ConfigService.Change(New Config With {
                                           .Prop = "AutoFitColumns",
                                           .Locale = "-", ' no locale
                                           .Val = value,
                                           .Lastchange = Date.Now
                })
                _autoFitColumns = value
            Catch ex As Exception

            End Try
        End Set
    End Property

    Private _recordLimit As Integer?

    Public Property RecordLimit As Integer
        Get
            If _recordLimit Is Nothing Then
                _recordLimit = ConfigService.GetRecordLimit()
            End If
            Return _recordLimit
        End Get
        Set(value As Integer)
            Try
                ConfigService.Change(New Config With {
                                           .Prop = "RecordLimit",
                                           .Locale = "-", ' no locale
                                           .Val = value,
                                           .Lastchange = Date.Now
                })
                _recordLimit = value
            Catch ex As Exception

            End Try
        End Set
    End Property

    ''' <summary>
    ''' 前缀（db）+ DataSource.id（current=true）
    ''' </summary>
    Private _curDataSource As String

    ''' <summary>
    ''' 当前数据库id（DataSource.id）
    ''' TODO 考虑增加ribbon功能，指定读取全局设定（已实现）/个别文档设定（不同文档关联不同的数据源）
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property CurDataSource() As DataSource
        Get
            Try
                If _curDataSource Is Nothing Then
                    datasources = DatasourceService.FindAllDataSource()
                    Dim cur As DataSource = datasources.Find(Function(ds) ds.Current = True)
                    If cur IsNot Nothing Then
                        _curDataSource = GetControlId(cur.Id)
                    End If
                End If
                Return datasources.Find(Function(ds) ds.Id = Right(_curDataSource, Len(_curDataSource) - 2))
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    ''' <summary>
    ''' 根据数据源id拼接生成控件id
    ''' </summary>
    ''' <param name="id"></param>
    ''' <returns></returns>
    Private Function GetControlId(id As String) As String
        Return String.Concat("db", id)
    End Function

    ''' <summary>
    ''' 根据控件id截取数据源id
    ''' </summary>
    ''' <param name="ctrlId"></param>
    ''' <returns></returns>
    Private Function GetDataSourceId(ctrlId As String) As String
        Return Right(ctrlId, Len(ctrlId) - 2)
    End Function

    Public Sub New()

        ' 补丁1.0.0.13：DATASOURCE表添加MODE字段
        DatasourceService.Patch_1_0_0_13()
        DatasourceService.Patch_1_0_0_15()

        datasources = DatasourceService.FindAllDataSource()
        Dim cur As DataSource = datasources.Find(Function(ds) ds.Current = True)
        If cur IsNot Nothing Then
            _curDataSource = GetControlId(cur.Id)
        End If

    End Sub

    Private Function TestLanguageSupport(lang As String) As Boolean
        ' 默认切换到本地语言
        Dim tgtLang = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName
        If TextService.GetTextByProperty(TextType.TT_RB_QUERY, lang) Is Nothing Then
            Dim unsupportedMsg1 As String
            Dim unsupportedMsg2 As String
            ' 指定语言与本地语言相同
            If tgtLang = lang Then
                ' 以英语弹出提示，并切换为英语
                tgtLang = "en"
                unsupportedMsg1 = TextService.GetTextByProperty(TextType.TT_MSG_UNSUPPORTED_LANG_1, tgtLang)
                unsupportedMsg2 = TextService.GetTextByProperty(TextType.TT_MSG_UNSUPPORTED_LANG_2, tgtLang)
            Else
                ' 尝试本地语言
                unsupportedMsg1 = TextService.GetTextByProperty(TextType.TT_MSG_UNSUPPORTED_LANG_1, tgtLang)
                unsupportedMsg2 = TextService.GetTextByProperty(TextType.TT_MSG_UNSUPPORTED_LANG_2, tgtLang)
                If unsupportedMsg1 Is Nothing Or unsupportedMsg2 Is Nothing Then
                    ' 本地语言不支持，则切换为英语
                    tgtLang = "en"
                    unsupportedMsg1 = TextService.GetTextByProperty(TextType.TT_MSG_UNSUPPORTED_LANG_1, tgtLang)
                    unsupportedMsg2 = TextService.GetTextByProperty(TextType.TT_MSG_UNSUPPORTED_LANG_2, tgtLang)
                End If
            End If

            MsgBox(String.Format(unsupportedMsg1, CultureInfo.GetCultureInfo(lang).DisplayName) & vbCrLf _
                   & String.Format(unsupportedMsg2, CultureInfo.GetCultureInfo(tgtLang).DisplayName))

            Globals.ThisAddIn.StrLangCode = tgtLang
            Return False
        End If

        Return True
    End Function

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("SwiftConnector.Ribbon.xml")
    End Function

#Region "功能区回调"
    '在此处创建回叫方法。有关添加回叫方法的详细信息，请访问 https://go.microsoft.com/fwlink/?LinkID=271226
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        TestLanguageSupport(Globals.ThisAddIn.StrLangCode)
        Me.ribbon = ribbonUI
    End Sub

    Public Sub BtnSelect_Click(ByVal control As Office.IRibbonControl)
        logger.Debug("BtnSelect_Click Start")
        If IsEditing() Then
            MsgBox("Excel is in Edit Mode.")
            Return
        End If

        Globals.ThisAddIn.Application.ScreenUpdating = False
        Dim tableObj As XlTable = Nothing
        Try
            tableObj = XlTable.Create()
            tableObj.Render()
        Catch tnfex As TableNotFoundException
            MsgBox("Table cannot be found!")
        Catch ex As Exception
            logger.Error(ex)
            If tableObj IsNot Nothing Then
                tableObj.Revoke()
            End If
            MsgBox(ex.Message)
        End Try
        Globals.ThisAddIn.Application.ScreenUpdating = True
        logger.Debug("BtnSelect_Click End")
    End Sub

    Public Sub BtnInsert_Click(ByVal control As Office.IRibbonControl)
        If IsEditing() Then
            MsgBox("Excel is in Edit Mode.")
            Return
        End If

        Dim tableObj As XlTable = Nothing
        Try
            tableObj = XlTable.Create()
            tableObj.Render(RenderMode.Memory)
            tableObj.Save()
        Catch ex As Exception
            logger.Error(ex)
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub BtnDelete_Click(ByVal control As Office.IRibbonControl)
        If IsEditing() Then
            MsgBox("Excel is in Edit Mode.")
            Return
        End If

        Dim tableObj As XlTable = Nothing
        Try
            tableObj = XlTable.Create()
            tableObj.Render(RenderMode.Memory)
            tableObj.Delete()
        Catch ex As Exception
            logger.Error(ex)
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Function BtnDelete_GetScreenTip(ByVal control As Office.IRibbonControl) As String
        Select Case DelMode
            Case DeleteMode.Truncate
                Return BUTTON_TRUNCATE_SCREEN_TIP
            Case DeleteMode.Delete
                Return BUTTON_DELETE_SCREEN_TIP
            Case Else
                ' DEAD CODE
        End Select
        Return Nothing
    End Function

    Public Function BtnDelete_GetLabel(ByVal control As Office.IRibbonControl) As String
        Dim tt As TextType? = Nothing
        Select Case DelMode
            Case DeleteMode.Truncate
                tt = TextType.TT_RB_TRUNCATE
            Case DeleteMode.Delete
                tt = TextType.TT_RB_DELETE
            Case Else
                ' DEAD CODE
        End Select
        Return TextService.GetTextByProperty(tt).Replace("\r\n", vbCrLf)
    End Function

    Public Function BtnDelete_GetImage(ByVal control As Office.IRibbonControl) As String
        Select Case DelMode
            Case DeleteMode.Truncate
                Return BUTTON_TRUNCATE_IMAGE
            Case DeleteMode.Delete
                Return BUTTON_DELETE_IMAGE
            Case Else
                ' DEAD CODE
        End Select
        Return Nothing
    End Function

    Public Function BtnSelect_GetEnabled(ByVal control As Office.IRibbonControl) As Boolean
        Globals.ThisAddIn.HotKeys("Select").Enabled = CurDataSource() IsNot Nothing
        Return CurDataSource() IsNot Nothing
    End Function

    Public Function BtnInsert_GetLabel(ByVal control As Office.IRibbonControl) As String
        Dim tt As TextType? = Nothing
        Select Case ExecMode
            Case ExecuteMode.Normal
                tt = TextType.TT_RB_SAVE
            Case ExecuteMode.Swift
                tt = TextType.TT_RB_SAVE_SWIFT
            Case Else
                ' DEAD CODE
        End Select
        Return TextService.GetTextByProperty(tt).Replace("\r\n", vbCrLf)
    End Function

    Public Function BtnInsert_GetImage(ByVal control As Office.IRibbonControl) As String
        Select Case ExecMode
            Case ExecuteMode.Normal
                Return BUTTON_INSERT_NORMAL_IMAGE
            Case ExecuteMode.Swift
                Return BUTTON_INSERT_SWIFT_IMAGE
            Case Else
                ' DEAD CODE
        End Select
        Return Nothing
    End Function

    Public Function SpltBtnInsert_GetEnabled(ByVal control As Office.IRibbonControl) As Boolean
        Dim isEnabled = CurDataSource IsNot Nothing AndAlso CurDataSource.Mode = OperateMode.Normal
        Globals.ThisAddIn.HotKeys("Insert").Enabled = isEnabled
        Return isEnabled
    End Function

    Public Sub BtnInsertNormalMode_Click(ByVal control As Office.IRibbonControl)
        ExecMode = ExecuteMode.Normal
        ribbon.InvalidateControl("BtnInsert")
        ribbon.InvalidateControl("BtnInsertNormalMode")
        ribbon.InvalidateControl("BtnInsertSwiftMode")
    End Sub

    Public Sub BtnInsertSwiftMode_Click(ByVal control As Office.IRibbonControl)
        ExecMode = ExecuteMode.Swift
        ribbon.InvalidateControl("BtnInsert")
        ribbon.InvalidateControl("BtnInsertNormalMode")
        ribbon.InvalidateControl("BtnInsertSwiftMode")
    End Sub

    Public Function BtnInsertNormalMode_GetImage(ByVal control As Office.IRibbonControl) As String
        Select Case ExecMode
            Case ExecuteMode.Normal
                Return CHECKED_IMAGE
            Case ExecuteMode.Swift
                Return UNCHECKED_IMAGE
            Case Else
                ' DEAD CODE
        End Select
        Return Nothing
    End Function

    Public Function BtnInsertSwiftMode_GetImage(ByVal control As Office.IRibbonControl) As String
        Select Case ExecMode
            Case ExecuteMode.Normal
                Return UNCHECKED_IMAGE
            Case ExecuteMode.Swift
                Return CHECKED_IMAGE
            Case Else
                ' DEAD CODE
        End Select
        Return Nothing
    End Function

    'Public Function BtnInsert_GetEnabled(ByVal control As Office.IRibbonControl) As Boolean
    '    Dim isEnabled = CurDataSource IsNot Nothing AndAlso CurDataSource.Mode = OperateMode.Normal
    '    Globals.ThisAddIn.HotKeys("Insert").Enabled = isEnabled
    '    Return isEnabled
    'End Function

    Public Function SpltBtnDeleteOrTruncate_GetEnabled(ByVal control As Office.IRibbonControl) As Boolean
        Dim isEnabled = CurDataSource IsNot Nothing AndAlso CurDataSource.Mode = OperateMode.Normal
        Globals.ThisAddIn.HotKeys("Delete").Enabled = isEnabled
        Return isEnabled
    End Function

    Public Sub BtnTruncateMode_Click(ByVal control As Office.IRibbonControl)
        DelMode = DeleteMode.Truncate
        ribbon.InvalidateControl("BtnDelete")
        ribbon.InvalidateControl("BtnTruncateMode")
        ribbon.InvalidateControl("BtnDeleteMode")
    End Sub

    Public Function BtnTruncateMode_GetImage(ByVal control As Office.IRibbonControl) As String
        Select Case DelMode
            Case DeleteMode.Truncate
                Return CHECKED_IMAGE
            Case DeleteMode.Delete
                Return UNCHECKED_IMAGE
            Case Else
                ' DEAD CODE
        End Select
        Return Nothing
    End Function

    Public Function BtnImport_GetEnabled(ByVal control As Office.IRibbonControl) As Boolean
        Return CurDataSource IsNot Nothing AndAlso CurDataSource.Mode = OperateMode.Normal
    End Function

    Public Sub BtnImport_Click(ByVal control As Office.IRibbonControl)
        If IsEditing() Then
            MsgBox("Excel is in Edit Mode.")
            Return
        End If

        If MsgBox("Import data will delete old data," & vbCrLf & "are you sure?", MsgBoxStyle.OkCancel, "Warn") <> MsgBoxResult.Ok Then
            Return
        End If

        Dim app = Globals.ThisAddIn.Application
        Dim tableLocationDic As New Dictionary(Of String, Dictionary(Of String, Excel.Range))
        For Each sheet As Excel.Worksheet In app.ActiveWindow.SelectedSheets
            Dim col = sheet.Range("1:1").SpecialCells(Excel.XlCellType.xlCellTypeVisible).Column
            Dim sheetDic As New Dictionary(Of String, Excel.Range)
            For Each row As Excel.Range In sheet.UsedRange.Rows
                If Not row.EntireRow.Hidden Then
                    Dim cell As Excel.Range = row.EntireRow.Columns(col)
                    If cell.Value IsNot Nothing AndAlso cell.Value.ToString = "TABLE" Then
                        Dim tableId = cell.Offset(0, 1)
                        If tableId.Value IsNot Nothing AndAlso Not sheetDic.ContainsKey(tableId.Value) Then
                            sheetDic.Add(tableId.Value, tableId)
                        End If
                    End If
                End If
            Next
            If sheetDic.Count > 0 Then
                tableLocationDic.Add(sheet.Name, sheetDic)
            End If
        Next

        app.ScreenUpdating = False

        For Each st In tableLocationDic
            For Each tbl In st.Value
                logger.Debug(st.Key & " " & tbl.Key & " " & tbl.Value.Row & "," & tbl.Value.Column)

                tbl.Value.Worksheet.Activate()
                tbl.Value.Select()

                Dim tableObj As XlTable = Nothing
                Try
                    tableObj = XlTable.Create()
                    tableObj.Render(RenderMode.Memory)
                    tableObj.Delete(True)
                    tableObj.Save()
                Catch ex As Exception
                    logger.Error(ex)
                    MsgBox(ex.Message)
                End Try

            Next
        Next

        app.ScreenUpdating = True

    End Sub

    Public Sub TglBtnStyleSettings_Click(ByVal control As Office.IRibbonControl, isPressed As Boolean)
        Globals.ThisAddIn.TglBtnStyleSettingsPressed(Globals.ThisAddIn.Application.ActiveWorkbook) = isPressed
        Globals.ThisAddIn.CtpSettings(Globals.ThisAddIn.Application.ActiveWorkbook).Visible = isPressed
        ribbon.InvalidateControl("TglBtnStyleSettings")
        If isPressed Then
            Dim ctp = Globals.ThisAddIn.CtpSettings(Globals.ThisAddIn.Application.ActiveWorkbook)
            Dim hb As HostBrowser = ctp.Control
            hb.Fragment = "change-log"
        End If
    End Sub

    'Public Sub TglBtnProtectedMode_Click(ByVal control As Office.IRibbonControl, isPressed As Boolean)
    '    Try
    '        ProtectedMode = isPressed
    '    Catch ex As Exception
    '        logger.Error(ex)
    '    End Try
    '    ribbon.InvalidateControl("TglBtnProtectedMode")
    '    ribbon.InvalidateControl("SpltBtnInsert")
    '    ribbon.InvalidateControl("SpltBtnDeleteOrTruncate")
    'End Sub

    'Public Function TglBtnProtectedMode_GetLabel(ByVal control As Office.IRibbonControl) As String
    '    If ProtectedMode Then
    '        Return TextService.GetTextByProperty(TextType.TT_RB_PROTECTED_MODE).Replace("\r\n", vbCrLf)
    '    Else
    '        Return TextService.GetTextByProperty(TextType.TT_RB_NORMAL_MODE).Replace("\r\n", vbCrLf)
    '    End If
    'End Function

    'Public Function TglBtnProtectedMode_GetImage(ByVal control As Office.IRibbonControl) As String
    '    If ProtectedMode Then
    '        Return PROTECTED_MODE_IMAGE
    '    Else
    '        Return NORMAL_MODE_IMAGE
    '    End If
    'End Function

    Public Sub TglBtnShowProps_Click(ByVal control As Office.IRibbonControl, isPressed As Boolean)
        Try
            ShowProps = isPressed
        Catch ex As Exception
            logger.Debug(ex)
        End Try
    End Sub

    Public Sub TglBtnAutoFitColumns_Click(ByVal control As Office.IRibbonControl, isPressed As Boolean)
        Try
            AutoFitColumns = isPressed
        Catch ex As Exception
            logger.Debug(ex)
        End Try
    End Sub

    Public Function TglBtnStyleSettings_GetPressed(ByVal control As Office.IRibbonControl) As Boolean
        Return Globals.ThisAddIn.TglBtnStyleSettingsPressed(Globals.ThisAddIn.Application.ActiveWorkbook)
    End Function

    'Public Function TglBtnProtectedMode_GetPressed(ByVal control As Office.IRibbonControl) As Boolean
    '    Return ProtectedMode
    'End Function

    Public Function TglBtnShowProps_GetPressed(ByVal control As Office.IRibbonControl) As Boolean
        Return ShowProps
    End Function

    Public Function TglBtnAutoFitColumns_GetPressed(ByVal control As Office.IRibbonControl) As Boolean
        Return AutoFitColumns
    End Function

    Public Sub CtpSettings_VisibleChanged(sender As Object, e As EventArgs)
        If Not Globals.ThisAddIn.CtpSettings(Globals.ThisAddIn.Application.ActiveWorkbook).Visible And Globals.ThisAddIn.TglBtnStyleSettingsPressed(Globals.ThisAddIn.Application.ActiveWorkbook) Then
            Globals.ThisAddIn.TglBtnStyleSettingsPressed(Globals.ThisAddIn.Application.ActiveWorkbook) = Globals.ThisAddIn.CtpSettings(Globals.ThisAddIn.Application.ActiveWorkbook).Visible
            ribbon.InvalidateControl("TglBtnStyleSettings")
        End If
    End Sub

    Public Function DMenuDatabase_GetContent(ByVal control As Office.IRibbonControl) As String
        Dim sb As StringBuilder = New StringBuilder("<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" >")
        Dim tmpDsType As DataSourceType? = Nothing
        For Each ds In datasources
            If tmpDsType Is Nothing OrElse tmpDsType <> ds.Type Then
                If tmpDsType IsNot Nothing Then
                    sb.Append("</menu>")
                End If
                Dim typeName = [Enum].GetName(GetType(DataSourceType), ds.Type)
                sb.Append("<menu id=""" & typeName & """ label=""" & typeName & """ getImage=""DMenuDatabase_DsTypeGetImage"">")
            End If
            sb.Append("<button id=""" & GetControlId(ds.Id) & """ label=""" & ds.Name & """ getImage=""DMenuDatabase_LeafGetImage"" onAction=""DMenuDatabase_LeafSelect""/>")
            tmpDsType = ds.Type
        Next
        If datasources.Count > 0 Then sb.Append("</menu>")
        sb.Append("<menuSeparator id=""menuSeparator1""/>")
        sb.Append("<button id=""BtnDBManager"" getLabel=""Common_GetLabel"" imageMso=""AddInManager"" onAction=""BtnDBManager_Click""/>")
        sb.Append("</menu>")
        Return sb.ToString
    End Function

    Public Function DMenuDatabase_GetLabel(ByVal control As Office.IRibbonControl) As String
        Return If(CurDataSource Is Nothing, "N/A", If(CurDataSource.Name.Length > 10, Left(CurDataSource.Name, 7) & "···", CurDataSource.Name))
    End Function

    Public Function DMenuDatabase_DsTypeGetImage(ByVal control As Office.IRibbonControl) As Bitmap
        Return IconResource.ResourceManager.GetObject(control.Id.ToLower & "_mini_32")
    End Function

    Public Function DMenuDatabase_GetImage(ByVal control As Office.IRibbonControl) As Bitmap
        If CurDataSource Is Nothing Then
            Return Nothing
        End If
        Return IconResource.ResourceManager.GetObject([Enum].GetName(GetType(DataSourceType), CurDataSource.Type).ToLower & "_large_64")
    End Function

    Public Sub DMenuDatabase_LeafSelect(ByVal control As Office.IRibbonControl)
        Dim switchRst = DatasourceService.SwitchDataSourceTo(GetDataSourceId(control.Id))
        If switchRst Then
            _curDataSource = control.Id

            Dim title = TextService.GetTextByProperty(TextType.TT_MSG_SWITCH_SUCCESS)
            logger.Debug(title)
            Dim content = TextService.GetTextByProperty(TextType.TT_MSG_CONNECTION_IN_USE).Replace("{0}", CurDataSource.Name)
            logger.Debug(content)
            'Dim logo = "https://raw.githubusercontent.com/beatofangel/swift-connector-ui-release/master/images/database/" & DataSourceDic(CurDataSource.Type) & "_large_64.png"
            'Toast(title, content, logo)
            Toast(title, content)
        End If
        _curDataSource = Nothing
        ribbon.InvalidateControl("DMenuDatabase")
        ribbon.InvalidateControl("BtnSelect")
        'ribbon.InvalidateControl("BtnInsert")
        ribbon.InvalidateControl("SpltBtnInsert")
        ribbon.InvalidateControl("SpltBtnDeleteOrTruncate")
        ribbon.InvalidateControl("BtnImport")
        'ribbon.InvalidateControl("TglBtnProtectedMode")
        ribbon.InvalidateControl("DMenuOperateMode")
    End Sub

    Public Function DMenuDatabase_LeafGetImage(ByVal control As Office.IRibbonControl) As String
        If _curDataSource = control.Id Then
            Return CHECKED_IMAGE
        Else
            Return UNCHECKED_IMAGE
        End If
    End Function

    Public Sub BtnDeleteMode_Click(ByVal control As Office.IRibbonControl)
        DelMode = DeleteMode.Delete
        ribbon.InvalidateControl("BtnDelete")
        ribbon.InvalidateControl("BtnTruncateMode")
        ribbon.InvalidateControl("BtnDeleteMode")
    End Sub

    Public Function BtnDeleteMode_GetImage(ByVal control As Office.IRibbonControl) As String
        Select Case DelMode
            Case DeleteMode.Truncate
                Return UNCHECKED_IMAGE
            Case DeleteMode.Delete
                Return CHECKED_IMAGE
            Case Else
                ' DEAD CODE
        End Select
        Return Nothing
    End Function

    'Public Sub BtnTableFont_Click(ByVal control As Office.IRibbonControl)
    '    fontDlg.Font = New Font("SimSun", 12.0F, FontStyle.Regular)
    '    Dim rst = fontDlg.ShowDialog()
    '    If rst = DialogResult.OK Then
    '        Debug.Print(fontDlg.Font.Name)
    '    End If
    'End Sub

    'Public Sub BtnTableBgColor_Click(ByVal control As Office.IRibbonControl)
    '    Dim rst = colorDlg.ShowDialog()
    '    If rst = DialogResult.OK Then
    '        Debug.Print(colorDlg.Color.ToString)
    '    End If
    'End Sub

    'Public Sub BtnTableFontColor_Click(ByVal control As Office.IRibbonControl)
    '    Dim rst = colorDlg.ShowDialog()
    '    If rst = DialogResult.OK Then
    '        Debug.Print(colorDlg.Color.ToString)
    '    End If
    'End Sub

    'Public Function SpltBtnAccount_GetLabel(ByVal control As Office.IRibbonControl) As String
    '    Dim winUserName = Environment.UserName
    '    Dim domainName = Environment.UserDomainName
    '    Return winUserName
    'End Function

    Public Sub BtnDBManager_Click(ByVal control As Office.IRibbonControl)
        'Using frm = New FrmDatasources
        '    frm.ShowDialog()
        'End Using

        'Using frm = New FrmConnections
        Globals.ThisAddIn.DlgConnections.ShowDialog()
        'End Using

        ' update ribbon
        _curDataSource = Nothing
        ribbon.InvalidateControl("DMenuDatabase")
        ribbon.InvalidateControl("BtnSelect")
        ribbon.InvalidateControl("SpltBtnInsert")
        ribbon.InvalidateControl("SpltBtnDeleteOrTruncate")
        ribbon.InvalidateControl("BtnImport")
        'ribbon.InvalidateControl("TglBtnProtectedMode")
        ribbon.InvalidateControl("DMenuOperateMode")
    End Sub

    Public Function DMenuRecordLimit_GetLabel(ByVal control As Office.IRibbonControl) As String
        Return RecordLimit & vbCrLf
    End Function

    Public Function DMenuRecordLimit_GetContent(ByVal control As Office.IRibbonControl) As String
        Dim sb As StringBuilder = New StringBuilder("<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" >")
        Dim recordLimitArr As Integer() = {100, 500, 1000, 5000, 10000}
        For Each rl In recordLimitArr
            sb.Append("<button id=""DMenuRecordLimit" & rl & """ label=""" & rl & """ tag=""" & rl & """ getImage=""DMenuRecordLimit_LeafGetImage"" onAction=""DMenuRecordLimit_LeafSelect""/>")
        Next
        'sb.Append("<menuSeparator id=""menuSeparator1""/>")
        'sb.Append("<button id=""BtnResetLanguage"" getLabel=""Common_GetLabel"" getImage=""BtnResetLanguage_GetImage"" onAction=""BtnResetLanguage_Click""/>")
        sb.Append("</menu>")
        Return sb.ToString
    End Function

    Public Function DMenuRecordLimit_LeafGetImage(ByVal control As Office.IRibbonControl) As String
        If ("DMenuRecordLimit" & RecordLimit) = control.Id Then
            Return CHECKED_IMAGE
        Else
            Return UNCHECKED_IMAGE
        End If
    End Function

    Public Sub DMenuRecordLimit_LeafSelect(ByVal control As Office.IRibbonControl)
        RecordLimit = control.Tag
        ribbon.InvalidateControl("DMenuRecordLimit")
    End Sub

    Public Function DMenuOperateMode_GetContent(ByVal control As Office.IRibbonControl) As String
        Dim sb As StringBuilder = New StringBuilder("<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" >")
        sb.Append("<button id=""DMenuOperateModeNormal"" getLabel=""DMenuOperateMode_LeafGetLabel"" tag=""0"" getImage=""DMenuOperateMode_LeafGetImage"" onAction=""DMenuOperateMode_LeafSelect""/>")
        sb.Append("<button id=""DMenuOperateModeProtected"" getLabel=""DMenuOperateMode_LeafGetLabel"" tag=""1"" getImage=""DMenuOperateMode_LeafGetImage"" onAction=""DMenuOperateMode_LeafSelect""/>")
        sb.Append("</menu>")
        Return sb.ToString
    End Function

    Public Function DMenuOperateMode_GetLabel(ByVal control As Office.IRibbonControl) As String
        If ProtectedMode Then
            Return TextService.GetTextByProperty(TextType.TT_RB_PROTECTED_MODE).Replace("\r\n", vbCrLf)
        Else
            Return TextService.GetTextByProperty(TextType.TT_RB_NORMAL_MODE).Replace("\r\n", vbCrLf)
        End If
    End Function

    Public Function DMenuOperateMode_GetImage(ByVal control As Office.IRibbonControl) As String
        If ProtectedMode Then
            Return PROTECTED_MODE_IMAGE
        Else
            Return NORMAL_MODE_IMAGE
        End If
    End Function

    Public Function DMenuOperateMode_LeafGetLabel(ByVal control As Office.IRibbonControl) As String
        Dim opMode As OperateMode = ParseEnum(Of OperateMode)(control.Tag)
        Select Case opMode
            Case OperateMode.Protected
                Return TextService.GetTextByProperty(TextType.TT_RB_PROTECTED_MODE).Replace("\r\n", " ")
            Case Else
                Return TextService.GetTextByProperty(TextType.TT_RB_NORMAL_MODE).Replace("\r\n", " ")
        End Select
    End Function

    Public Sub DMenuOperateMode_LeafSelect(ByVal control As Office.IRibbonControl)
        Try
            ProtectedMode = ParseEnum(Of OperateMode)(control.Tag) = OperateMode.Protected
        Catch ex As Exception
            logger.Error(ex)
        End Try
        ribbon.InvalidateControl("DMenuOperateMode")
        ribbon.InvalidateControl("SpltBtnInsert")
        ribbon.InvalidateControl("SpltBtnDeleteOrTruncate")
        ribbon.InvalidateControl("BtnImport")
    End Sub

    Public Function DMenuOperateMode_LeafGetImage(ByVal control As Office.IRibbonControl) As String
        Dim opMode As OperateMode = ParseEnum(Of OperateMode)(control.Tag)
        If ProtectedMode Then
            Return If(opMode = OperateMode.Protected, CHECKED_IMAGE, UNCHECKED_IMAGE)
        Else
            Return If(opMode = OperateMode.Normal, CHECKED_IMAGE, UNCHECKED_IMAGE)
        End If
    End Function

    Public Function DMenuOperateMode_GetScreentip(ByVal control As Office.IRibbonControl) As String
        If ProtectedMode Then
            Return TextService.GetTextByProperty(TextType.TT_RB_PROTECTED_MODE).Replace("\r\n", " ")
        Else
            Return TextService.GetTextByProperty(TextType.TT_RB_NORMAL_MODE).Replace("\r\n", " ")
        End If
    End Function

    'Public Function TglBtnProtectedMode_GetScreentip(ByVal control As Office.IRibbonControl) As String
    '    If ProtectedMode Then
    '        Return TextService.GetTextByProperty(TextType.TT_RB_PROTECTED_MODE).Replace("\r\n", " ")
    '    Else
    '        Return TextService.GetTextByProperty(TextType.TT_RB_NORMAL_MODE).Replace("\r\n", " ")
    '    End If
    'End Function

    Public Function DMenuRecordLimit_GetScreenTip(ByVal control As Office.IRibbonControl) As String
        Return "Records limit to " & RecordLimit
    End Function

    Public Sub BtnResetStyle_Click(ByVal control As Office.IRibbonControl)
        If MsgBox("All configurations will be reset," & vbCrLf & "are you sure?", MsgBoxStyle.OkCancel, "Warn") = MsgBoxResult.Ok Then
            StyleService.Reset()
            Dim ctp = Globals.ThisAddIn.CtpSettings(Globals.ThisAddIn.Application.ActiveWorkbook)
            If ctp.Visible Then ctp.Visible = False
        End If
    End Sub

    Public Function DMenuLanguage_GetContent(ByVal control As Office.IRibbonControl) As String
        Dim sb As StringBuilder = New StringBuilder("<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" >")
        For Each key In langLabelDic.Keys
            sb.Append("<button id=""" & key & """ label=""" & langLabelDic(key) & """ getImage=""DMenuLanguage_LeafGetImage"" onAction=""DMenuLanguage_LeafSelect""/>")
        Next
        sb.Append("<menuSeparator id=""menuSeparator1""/>")
        sb.Append("<button id=""BtnResetLanguage"" getLabel=""Common_GetLabel"" getImage=""BtnResetLanguage_GetImage"" onAction=""BtnResetLanguage_Click""/>")
        sb.Append("</menu>")
        Return sb.ToString
    End Function

    Public Function DMenuLanguage_GetLabel(ByVal control As Office.IRibbonControl) As String
        Return langLabelDic(Globals.ThisAddIn.StrLangCode) & vbCrLf
    End Function

    Public Function DMenuLanguage_GetImage(ByVal control As Office.IRibbonControl) As Bitmap
        Return IconResource.ResourceManager.GetObject(Globals.ThisAddIn.StrLangCode)
    End Function

    Public Async Sub DMenuLanguage_LeafSelect(ByVal control As Office.IRibbonControl)
        If TestLanguageSupport(control.Id) Then
            Globals.ThisAddIn.StrLangCode = control.Id
        End If

        Dim ctp = Globals.ThisAddIn.CtpSettings(Globals.ThisAddIn.Application.ActiveWorkbook)
        If ctp.Visible Then
            Dim hb As HostBrowser = ctp.Control
            Await hb.SwitchLanguage()
        End If
        ribbon.Invalidate()
    End Sub

    Public Function DMenuLanguage_LeafGetImage(ByVal control As Office.IRibbonControl) As String
        If Globals.ThisAddIn.StrLangCode = control.Id Then
            Return CHECKED_IMAGE
        Else
            Return UNCHECKED_IMAGE
        End If
    End Function

    Public Function BtnResetLanguage_GetImage(ByVal control As Office.IRibbonControl) As Bitmap
        Return IconResource.ResourceManager.GetObject(CultureInfo.CurrentUICulture.TwoLetterISOLanguageName)
    End Function

    Public Sub BtnResetLanguage_Click(ByVal control As Office.IRibbonControl)
        Globals.ThisAddIn.StrLangCode = Nothing
        Dim ctp = Globals.ThisAddIn.CtpSettings(Globals.ThisAddIn.Application.ActiveWorkbook)
        If ctp.Visible Then ctp.Visible = False
        ribbon.Invalidate()
    End Sub

    Public Sub BtnAbout_Click(ByVal control As Office.IRibbonControl)
        Dim ctp = Globals.ThisAddIn.CtpSettings(Globals.ThisAddIn.Application.ActiveWorkbook)
        Dim hb As HostBrowser = ctp.Control
        If Not ctp.Visible Then
            Globals.ThisAddIn.TglBtnStyleSettingsPressed(Globals.ThisAddIn.Application.ActiveWorkbook) = True
            ctp.Visible = True
            ribbon.InvalidateControl("TglBtnStyleSettings")
        End If
        hb.Fragment = "about"
        'Using about As New FrmAbout
        '    about.ShowDialog()
        'End Using
    End Sub

    Public Function Common_GetLabel(ByVal control As Office.IRibbonControl) As String
        Dim tt As TextType? = Nothing
        Select Case control.Id
            Case "SwiftConnector"
                tt = TextType.TT_RB_TAB
            Case "GrpOperations"
                tt = TextType.TT_RB_GRP_OPERATIONS
            Case "GrpDataSource"
                tt = TextType.TT_RB_GRP_DATASOURCE
            Case "GrpSettings"
                tt = TextType.TT_RB_GRP_SETTINGS
            Case "GrpLanguage"
                tt = TextType.TT_RB_GRP_LANGUAGE
            Case "GrpAbout"
                tt = TextType.TT_RB_GRP_OTHER
            Case "BtnSelect"
                tt = TextType.TT_RB_QUERY
            Case "BtnInsertNormalMode"
                tt = TextType.TT_RB_SAVE
            Case "BtnInsertSwiftMode"
                tt = TextType.TT_RB_SAVE_SWIFT
            Case "BtnTruncateMode"
                tt = TextType.TT_RB_TRUNCATE_MODE
            Case "BtnDeleteMode"
                tt = TextType.TT_RB_DELETE_MODE
            Case "BtnDBManager"
                tt = TextType.TT_RB_DS_MANAGEMENT
            Case "TglBtnStyleSettings"
                tt = TextType.TT_RB_STYLE_SETTINGS
            Case "TglBtnShowProps"
                tt = TextType.TT_RB_SHOW_PROPS
            Case "TglBtnAutoFitColumns"
                tt = TextType.TT_RB_AUTO_FIT_COLUMNS
            Case "BtnResetStyle"
                tt = TextType.TT_RB_RESET_STYLE
            Case "BtnResetLanguage"
                tt = TextType.TT_RB_RESET_LANG
            Case "BtnAbout"
                tt = TextType.TT_RB_ABOUT
            Case "BtnImport"
                tt = TextType.TT_RB_IMPORT
        End Select

        Return TextService.GetTextByProperty(tt).Replace("\r\n", vbCrLf)
    End Function

#End Region

#Region "帮助器"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
