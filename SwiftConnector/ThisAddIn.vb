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

Imports System.Deployment.Application
Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports log4net

<Assembly: log4net.Config.XmlConfigurator(ConfigFile:="Log4net.config", Watch:=True)>
Public Class ThisAddIn

    Protected Shared logger As ILog = LogManager.GetLogger(Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    'Private langCode As Integer

    Private _dlgConnections As FrmConnections

    Public ReadOnly Property DlgConnections As FrmConnections
        Get
            If _dlgConnections Is Nothing Then
                _dlgConnections = New FrmConnections()
            End If
            Return _dlgConnections
        End Get
    End Property

    Private _myRibbon As Ribbon

    Public ReadOnly Property MyRibbon As Ribbon
        Get
            If _myRibbon Is Nothing Then
                _myRibbon = New Ribbon()
            End If
            Return _myRibbon
        End Get
    End Property

    Private _hotKeys As XlHotKeys
    Public ReadOnly Property HotKeys As XlHotKeys
        Get
            Return _hotKeys
        End Get
    End Property

    Private configService As New ConfigService

    Private textService As New TextService

    Private _tglBtnStyleSettingsPressed As New Dictionary(Of String, Boolean)
    Public Property TglBtnStyleSettingsPressed(wb As Excel.Workbook) As Boolean
        Get
            If _tglBtnStyleSettingsPressed.ContainsKey(wb.FullName) Then
                Return _tglBtnStyleSettingsPressed(wb.FullName)
            Else
                _tglBtnStyleSettingsPressed.Add(wb.FullName, False)
                Return False
            End If
        End Get
        Set(value As Boolean)
            If _tglBtnStyleSettingsPressed.ContainsKey(wb.FullName) Then
                _tglBtnStyleSettingsPressed(wb.FullName) = value
            Else
                _tglBtnStyleSettingsPressed.Add(wb.FullName, value)
            End If
        End Set
    End Property
    Public ReadOnly Property CtpSettings(wb As Excel.Workbook) As Microsoft.Office.Tools.CustomTaskPane
        Get
            Dim taskPane = WbCtp.Where(Function(kv) kv.Key = wb.FullName).FirstOrDefault().Value
            If taskPane Is Nothing Then
                Dim hb As New HostBrowser
                _ctpSettings = CustomTaskPanes.Add(hb, textService.GetTextByProperty(TextType.TT_TP_SETTINGS))
                _ctpSettings.Width = 554
                _ctpSettings.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
                AddHandler _ctpSettings.VisibleChanged, AddressOf MyRibbon.CtpSettings_VisibleChanged
                AddHandler _ctpSettings.VisibleChanged, AddressOf hb.HostBrowser_VisibleChanged
                'Dim st As New SettingsTabs
                '_ctpSettings = CustomTaskPanes.Add(st, "Settings")
                '_ctpSettings.Width = 554
                '_ctpSettings.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
                'AddHandler _ctpSettings.VisibleChanged, AddressOf MyRibbon.CtpSettings_VisibleChanged
                'AddHandler _ctpSettings.VisibleChanged, AddressOf st.SettingsTabs_VisibleChanged
                'Dim st As New TableSettings
                '_ctpSettings = CustomTaskPanes.Add(st, "Settings")
                '_ctpSettings.Width = 554
                '_ctpSettings.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
                'AddHandler _ctpSettings.VisibleChanged, AddressOf myRibbon.CtpSettings_VisibleChanged
                'AddHandler _ctpSettings.VisibleChanged, AddressOf st.TableSettings_VisibleChanged
                WbCtp.Add(wb.FullName, _ctpSettings)
            Else
                _ctpSettings = taskPane
            End If
            Return _ctpSettings
        End Get
    End Property

    Private WithEvents _ctpSettings As Microsoft.Office.Tools.CustomTaskPane
    Private WbCtp As New Dictionary(Of String, Microsoft.Office.Tools.CustomTaskPane)

    Public ReadOnly Property CurDataSource As DataSource
        Get
            Return MyRibbon.CurDataSource
        End Get
    End Property

    Private _strLangCode As String
    Public Property StrLangCode As String
        Get
            If _strLangCode Is Nothing Then
                Dim cfg = configService.Read(New Config With {
                                                    .Prop = "Language",
                                                    .Locale = "-"
                                                   })
                _strLangCode = If(cfg Is Nothing OrElse cfg.Val = "-", CultureInfo.CurrentUICulture.TwoLetterISOLanguageName, cfg.Val)
            End If
            Return _strLangCode
        End Get
        Set(value As String)
            Try
                configService.Change(New Config With {
                                            .Prop = "Language",
                                            .Locale = "-",
                                            .Val = If(value Is Nothing, "-", value),
                                            .Lastchange = Date.Now
                                           })
                _strLangCode = value
            Catch ex As Exception
                MsgBox("language switch failed!")
            End Try
        End Set
    End Property

    Private Sub ThisAddIn_Startup() Handles Me.Startup

        logger.Debug("ThisAddIn_Startup")
        'Dim installer As New SwiftConnectorInstaller
        'installer.InstallApplication("http://localhost/swift-connector/SwiftConnector.vsto")
        'langCode = Application.LanguageSettings.LanguageID(Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDUI)
        System.Windows.Forms.Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException)
        _hotKeys = New XlHotKeys()
        _hotKeys.Add("Select", New XlHotKey("BtnSelect_Click", KEY_CTRL, "Q"))
        _hotKeys.Add("Insert", New XlHotKey("BtnInsert_Click", KEY_CTRL, "W"))
        _hotKeys.Add("Delete", New XlHotKey("BtnDelete_Click", KEY_CTRL, "D"))
        '_hotKeys.Add("SettingStyle", New XlHotKey("TglBtnStyleSettings_Click", KEY_CTRL, KEY_SHIFT, "Enter")) ' 与excel数组公式快捷键冲突
        _hotKeys.Bind(Sub(hotkey As String)
                          Try
                              For Each key In _hotKeys.Keys
                                  Dim xlHotKey = _hotKeys.Item(key)
                                  If xlHotKey.Enabled And xlHotKey.HotKey = hotkey Then
                                      Dim ribbonCtrl As Office.IRibbonControl = Nothing
                                      If key = "SettingStyle" Then
                                          ' ToggleButton
                                          Dim pressed = Not Globals.ThisAddIn.TglBtnStyleSettingsPressed(Globals.ThisAddIn.Application.ActiveWorkbook)
                                          CallByName(MyRibbon, xlHotKey.Method, CallType.Method, ribbonCtrl, pressed)
                                      Else
                                          ' Button
                                          CallByName(MyRibbon, xlHotKey.Method, CallType.Method, ribbonCtrl)
                                      End If
                                  End If
                              Next
                          Catch ex As COMException
                              If ex.HResult = &H800AC472 Then
                                  MsgBox("Please close all dialogs of Excel before using the hotkey to run this add-in.")
                              End If
                          End Try
                      End Sub)

        If ApplicationDeployment.IsNetworkDeployed Then
            Dim ver = ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
            logger.Debug(String.Format("local version check {0}", ver))
            'If Not ApplicationDeployment.CurrentDeployment.IsFirstRun Then
            '    If ApplicationDeployment.CurrentDeployment.Update() Then
            '        logger.Debug("update success")
            '    Else
            '        logger.Debug("update failed")
            '    End If
            'End If

            If configService.GetVersion() <> ver Then
                Dim sql = "ATTACH DATABASE '" & Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "userdata.db") & "' as 'newConfig';" &
                    "DELETE FROM TEXT_PRESET;" &
                    "INSERT INTO TEXT_PRESET (PROP,LOCALE,VAL,LASTCHANGE) SELECT T1.PROP,T1.LOCALE,T1.VAL,T1.LASTCHANGE FROM newConfig.TEXT_PRESET T1;"
                textService.ExecuteNonQuery(Sub(rstList)
                                            End Sub, sql)
                configService.Change(New Config With {
                                                .Prop = "Version",
                                                .Locale = "-",
                                                .Val = ver,
                                                .Lastchange = Date.Now
                                     })
                'Using frmChangeLog As New FrmChangeLog
                '    frmChangeLog.ShowDialog()
                'End Using
            End If
        Else
            ' 开发版需手动清理userdata.db
        End If

        AddHandler Application.WorkbookBeforeClose, AddressOf Application_WorkbookBeforeClose
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        _hotKeys.Unbind()
    End Sub

    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return MyRibbon
    End Function

    'Private Sub Application_WorkbookActivate(wb As Excel.Workbook)
    '    Dim temp = CtpSettings(wb)
    '    'Dim taskPane = WbCtp.Where(Function(kv) kv.Key = wb.FullName).FirstOrDefault().Value
    '    'If taskPane Is Nothing Then
    '    '    _ctpSettings = Globals.ThisAddIn.CustomTaskPanes.Add(New SettingsTabs, "Settings")
    '    '    _ctpSettings.Width = 400
    '    '    _ctpSettings.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
    '    '    AddHandler _ctpSettings.VisibleChanged, AddressOf myRibbon.CtpSettings_VisibleChanged
    '    '    WbCtp.Add(wb.FullName, _ctpSettings)
    '    'Else
    '    '    _ctpSettings = taskPane
    '    'End If
    'End Sub

    Private Sub Application_WorkbookBeforeClose(wb As Excel.Workbook, ByRef Cancel As Boolean)
        _tglBtnStyleSettingsPressed.Remove(wb.FullName)
        WbCtp.Remove(wb.FullName)
    End Sub
End Class
