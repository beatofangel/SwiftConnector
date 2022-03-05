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
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Drawing
Imports System.Drawing.Text
Imports System.IO
Imports System.Runtime.InteropServices
Imports log4net
Imports MySqlConnector
Imports Oracle.ManagedDataAccess.Client
Imports SwiftConnector.My.Resources
Imports Windows.UI.Notifications

Module CommonHelper

    Private logger As ILog = LogManager.GetLogger(Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private Const DEFAULT_INTERNAL_DATASOURCE As String = "SwiftConnector/userdata.db"

    Private Const DEFAULT_ROW_LIMIT As Integer = 100

    Private _internal_config_database As String

    Private _table_comment_like_query_row_limit As Integer

    ''' <summary>
    ''' 判断是否是设计模式
    ''' </summary>
    ''' <returns></returns>
    Public Function IsDesignMode() As Boolean
        Return LicenseManager.UsageMode = LicenseUsageMode.Designtime OrElse Process.GetCurrentProcess().ProcessName = "devenv"
    End Function

    ''' <summary>
    ''' 插件配置用内部数据库（SQLite）文件路径
    ''' </summary>
    ''' <returns>String</returns>
    Public ReadOnly Property InternalConfigDatabase() As String
        Get
            If String.IsNullOrEmpty(_internal_config_database) Then
                Dim dbFile As String = Nothing
                Dim icd As String = ConfigurationManager.AppSettings.Get("InternalConfigDatabase")
                If String.IsNullOrEmpty(icd) Then
                    dbFile = Path.Combine(Environ("UserProfile"), DEFAULT_INTERNAL_DATASOURCE)
                Else
                    dbFile = Environment.ExpandEnvironmentVariables(icd)
                End If

                Dim dir = Path.GetDirectoryName(dbFile)
                If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)

                If Not File.Exists(dbFile) Then
                    File.Copy(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "userdata.db"), dbFile)
                End If

                _internal_config_database = dbFile
            End If

            Return _internal_config_database
        End Get
    End Property

    ''' <summary>
    ''' 表注释模糊查询时最大条数限制
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property TableCommentLikeQueryRowLimit() As Integer
        Get
            If _table_comment_like_query_row_limit = 0 Then
                Dim rl As String = ConfigurationManager.AppSettings.Get("TableCommentLikeQueryRowLimit")
                If String.IsNullOrEmpty(rl) Then
                    _table_comment_like_query_row_limit = DEFAULT_ROW_LIMIT
                Else
                    _table_comment_like_query_row_limit = CInt(rl)
                End If
            End If
            Return _table_comment_like_query_row_limit
        End Get
    End Property

    ''' <summary>
    ''' 获取当前数据库连接
    ''' </summary>
    ''' <returns></returns>
    Public Function GetCurrentConnection() As IDbConnection
        Dim connStr As String = Nothing
        Dim rst As IDbConnection = Nothing
        Dim curDs = Globals.ThisAddIn.CurDataSource()
        rst = DbConnectionFactory.CreateConnection(curDs)
        Return rst
    End Function

    ''' <summary>
    ''' 测试数据库连接是否成功
    ''' </summary>
    ''' <param name="ds"></param>
    Public Sub TestConnection(ds As DataSource)
        Using conn = DbConnectionFactory.CreateConnection(ds)
            conn.Open()
        End Using
    End Sub

    Private loadedFontFileList As New List(Of String)
    Private pfc As New PrivateFontCollection

    ''' <summary>
    ''' 加载私有字体
    ''' </summary>
    ''' <param name="fontFileName"></param>
    ''' <param name="fontSize"></param>
    ''' <param name="fontStyle"></param>
    ''' <returns></returns>
    Public Function GetPrivateFont(fontFileName As String, fontSize As Single, fontStyle As FontStyle) As Font
        If Not loadedFontFileList.Contains(fontFileName) Then
            Dim asm = Reflection.Assembly.GetExecutingAssembly
            Dim appName = asm.ManifestModule.Name.Split(".")(0)
            Dim fontAsStream = asm.GetManifestResourceStream(String.Format("{0}.{1}", appName, fontFileName))
            Dim fontAsByte = New Byte(fontAsStream.Length - 1) {}
            fontAsStream.Read(fontAsByte, 0, CInt(fontAsStream.Length))
            fontAsStream.Close()
            Dim memPointer As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(Byte)) * fontAsByte.Length)
            Marshal.Copy(fontAsByte, 0, memPointer, fontAsByte.Length)
            pfc.AddMemoryFont(memPointer, fontAsByte.Length)
            loadedFontFileList.Add(fontFileName)
        End If
        Return New Font(pfc.Families(0), fontSize, fontStyle)
    End Function

    ''' <summary>
    ''' 将指定字体（MaterialIcons-Regular.ttf）的文字绘制成原比例图片
    ''' </summary>
    ''' <param name="text"></param>
    ''' <param name="font"></param>
    ''' <param name="textColor"></param>
    ''' <param name="backColor"></param>
    ''' <param name="quality"></param>
    ''' <returns></returns>
    Public Function DrawText(ByVal text As String, ByVal font As Font, ByVal textColor As Color, ByVal backColor As Color, Optional quality As TextRenderingHint = TextRenderingHint.AntiAliasGridFit) As Image
        Dim img As Image = New Bitmap(1, 1)
        Dim textSize As SizeF
        Using drawing As Graphics = Graphics.FromImage(img)
            textSize = drawing.MeasureString(text, font)
            img = New Bitmap(CInt(textSize.Width), CInt(textSize.Height))
        End Using
        Using drawing As Graphics = Graphics.FromImage(img), textBrush As Brush = New SolidBrush(textColor)
            drawing.TextRenderingHint = quality
            drawing.Clear(backColor)
            drawing.DrawString(text, font, textBrush, 0, 0)
            drawing.Save()
        End Using

        Return img
    End Function

    ''' <summary>
    ''' 将指定字体（MaterialIcons-Regular.ttf）的文字绘制成方形图片
    ''' </summary>
    ''' <param name="text"></param>
    ''' <param name="font"></param>
    ''' <param name="textColor"></param>
    ''' <param name="backColor"></param>
    ''' <param name="times"></param>
    ''' <param name="quality"></param>
    ''' <returns></returns>
    Public Function DrawSquareText(ByVal text As String, ByVal font As Font, ByVal textColor As Color, ByVal backColor As Color, Optional times As Single = 1.0F, Optional quality As TextRenderingHint = TextRenderingHint.AntiAliasGridFit) As Image
        Dim img As Image = New Bitmap(1, 1)
        Dim textSize As SizeF
        Dim dummyTextSize As SizeF
        Dim dummyFont As Font = New Font(font.FontFamily, font.Size * times, font.Style)
        Dim deltaHeight As Integer
        Using drawing As Graphics = Graphics.FromImage(img)
            textSize = drawing.MeasureString(text, font)
            dummyTextSize = drawing.MeasureString(text, dummyFont)
            deltaHeight = If(dummyTextSize.Width > dummyTextSize.Height, (dummyTextSize.Width - dummyTextSize.Height) \ 2, 0)
            img = New Bitmap(CInt(Math.Max(textSize.Width, textSize.Height)), CInt(Math.Max(textSize.Width, textSize.Height)))
        End Using
        Using drawing As Graphics = Graphics.FromImage(img), textBrush As Brush = New SolidBrush(textColor)
            Dim offset As Single = (img.Width - CInt(Math.Max(dummyTextSize.Width, dummyTextSize.Height))) \ 2
            drawing.Clear(backColor)
            drawing.TextRenderingHint = quality
            drawing.DrawString(text, dummyFont, textBrush, offset, offset + deltaHeight)
            drawing.Save()
        End Using

        Return img
    End Function

    ''' <summary>
    ''' <para>将指定字体（MaterialIcons-Regular.ttf）的文字（e23a(填充颜色)，e22b(边框颜色)，e23c(字体颜色)）,分段绘制成图片</para>
    ''' <para>上部为固定颜色(textColorPrimary)，下部为指定的颜色(textColorSecondary)</para>
    ''' </summary>
    ''' <param name="text"></param>
    ''' <param name="font"></param>
    ''' <param name="textColorPrimary"></param>
    ''' <param name="textColorSecondary"></param>
    ''' <param name="backColor"></param>
    ''' <param name="times"></param>
    ''' <param name="quality"></param>
    ''' <returns></returns>
    Private Function DrawTextForColorPicker(ByVal text As String, ByVal font As Font, ByVal textColorPrimary As Color, ByVal textColorSecondary As Color, ByVal backColor As Color, Optional times As Single = 1.0F, Optional quality As TextRenderingHint = TextRenderingHint.AntiAliasGridFit) As Image
        Dim img As Image = New Bitmap(1, 1)
        Dim textSize As SizeF
        Dim dummyTextSize As SizeF
        Dim dummyFont As Font = New Font(font.FontFamily, font.Size * times, font.Style)
        Dim baseHeight As Integer
        Dim coloredHeight As Integer
        Dim deltaHeight As Integer
        Dim offset As Integer
        Using drawing As Graphics = Graphics.FromImage(img)
            textSize = drawing.MeasureString(text, font)
            dummyTextSize = drawing.MeasureString(text, dummyFont)
            deltaHeight = If(dummyTextSize.Width > dummyTextSize.Height, (dummyTextSize.Width - dummyTextSize.Height) \ 2, 0)
            textSize = New SizeF(Math.Max(textSize.Width, textSize.Height), Math.Max(textSize.Width, textSize.Height))
            dummyTextSize = New SizeF(Math.Max(dummyTextSize.Width, dummyTextSize.Height), Math.Max(dummyTextSize.Width, dummyTextSize.Height))
            coloredHeight = dummyTextSize.Height * 2 \ 5
            baseHeight = CInt(dummyTextSize.Height) - coloredHeight
            offset = (textSize.Width - CInt(dummyTextSize.Width)) \ 2
            img = New Bitmap(CInt(textSize.Width), CInt(textSize.Height))
        End Using
        Dim imgBase = New Bitmap(CInt(textSize.Width), baseHeight + offset)
        Using drawing As Graphics = Graphics.FromImage(imgBase), textBrush As Brush = New SolidBrush(textColorPrimary)
            drawing.Clear(backColor)
            drawing.TextRenderingHint = quality
            drawing.DrawString(text, dummyFont, textBrush, offset, offset + deltaHeight)
            drawing.Save()
        End Using
        Dim imgColored = New Bitmap(CInt(textSize.Width), coloredHeight + offset)
        Using drawing As Graphics = Graphics.FromImage(imgColored), textBrush As Brush = New SolidBrush(textColorSecondary)
            drawing.Clear(backColor)
            drawing.TextRenderingHint = quality
            drawing.DrawString(text, dummyFont, textBrush, offset, -baseHeight + deltaHeight)
            drawing.Save()
        End Using
        Using drawing As Graphics = Graphics.FromImage(img)
            'drawing.TextRenderingHint = quality
            drawing.Clear(backColor)
            drawing.DrawImage(imgBase, New PointF(0, 0))
            drawing.DrawImage(imgColored, New PointF(0, imgBase.Height))
            drawing.Save()
        End Using

        Return img
    End Function

    ''' <summary>
    ''' 获取原比例图片
    ''' </summary>
    ''' <param name="text"></param>
    ''' <param name="textSize"></param>
    ''' <param name="textColor"></param>
    ''' <param name="backColor"></param>
    ''' <param name="quality"></param>
    ''' <returns></returns>
    Public Function GetImage(ByVal text As String, ByVal textSize As Single, ByVal textColor As Color, ByVal backColor As Color, Optional quality As TextRenderingHint = TextRenderingHint.AntiAliasGridFit) As Image
        Return DrawText(text, GetPrivateFont("MaterialIcons-Regular.ttf", textSize, FontStyle.Regular), textColor, backColor, quality)
    End Function

    ''' <summary>
    ''' 获取方形图片
    ''' </summary>
    ''' <param name="text"></param>
    ''' <param name="textSize"></param>
    ''' <param name="textColor"></param>
    ''' <param name="backColor"></param>
    ''' <param name="times"></param>
    ''' <param name="quality"></param>
    ''' <returns></returns>
    Public Function GetSquareImage(ByVal text As String, ByVal textSize As Single, ByVal textColor As Color, ByVal backColor As Color, Optional times As Single = 1.0F, Optional quality As TextRenderingHint = TextRenderingHint.AntiAliasGridFit) As Image
        Return DrawSquareText(text, GetPrivateFont("MaterialIcons-Regular.ttf", textSize, FontStyle.Regular), textColor, backColor, times, quality)
    End Function

    '''' <summary>
    '''' 为<see cref="ColorPickerButton"/>获取<see cref="Image"/>
    '''' </summary>
    '''' <param name="text"></param>
    '''' <param name="textSize"></param>
    '''' <param name="textColorPrimary"></param>
    '''' <param name="textColorSecondary"></param>
    '''' <param name="backColor"></param>
    '''' <param name="times"></param>
    '''' <param name="quality"></param>
    '''' <returns></returns>
    'Public Function GetImageForColorPicker(ByVal text As String, ByVal textSize As Single, ByVal textColorPrimary As Color, ByVal textColorSecondary As Color, ByVal backColor As Color, Optional times As Single = 1.0F, Optional quality As TextRenderingHint = TextRenderingHint.AntiAliasGridFit) As Image
    '    Return DrawTextForColorPicker(text, GetPrivateFont("MaterialIcons-Regular.ttf", textSize, FontStyle.Regular), textColorPrimary, textColorSecondary, backColor, times, quality)
    'End Function

    ''' <summary>
    ''' 计算点到线的距离
    ''' </summary>
    ''' <param name="x"></param>
    ''' <param name="y"></param>
    ''' <param name="ax"></param>
    ''' <param name="ay"></param>
    ''' <param name="bx"></param>
    ''' <param name="by"></param>
    ''' <returns></returns>
    Private Function GetDist(x As Double, y As Double, ax As Double, ay As Double, bx As Double, by As Double) As Double
        If (ax - bx) * (x - bx) + (ay - by) * (y - by) <= 0 Then Return Math.Sqrt((x - bx) * (x - bx) + (y - by) * (y - by))
        If (bx - ax) * (x - ax) + (by - ay) * (y - ay) <= 0 Then Return Math.Sqrt((x - ax) * (x - ax) + (y - ay) * (y - ay))
        Return Math.Abs((by - ay) * x - (bx - ax) * y + bx * ay - by * ax) / Math.Sqrt((ay - by) * (ay - by) + (ax - bx) * (ax - bx))
    End Function

    ''' <summary>
    ''' 计算点到线的距离
    ''' </summary>
    ''' <param name="point"></param>
    ''' <param name="lineStart"></param>
    ''' <param name="lineEnd"></param>
    ''' <returns></returns>
    Public Function GetDist(point As Point, lineStart As Point, lineEnd As Point) As Double
        Return GetDist(point.X, point.Y, lineStart.X, lineStart.Y, lineEnd.X, lineEnd.Y)
    End Function

    '''' <summary>
    '''' 递归查询当前控件所属<see cref="RegionType"/>
    '''' </summary>
    '''' <param name="ctrl"></param>
    '''' <returns></returns>
    'Public Function GetRegionType(ctrl As Windows.Forms.Control) As RegionType
    '    Dim p = ctrl.Parent
    '    While p IsNot Nothing
    '        If TypeOf p Is RegionPanel Then
    '            Exit While
    '        End If
    '        p = p.Parent
    '    End While
    '    Dim region As RegionPanel = p
    '    Return If(region Is Nothing, RegionType.RT_TABLE_GLOBAL, region.RegionType)
    'End Function

    Public Function GetEnumerableType(type As Type) As Type
        If type.IsInterface AndAlso type.GetGenericTypeDefinition() = GetType(IEnumerable(Of)) Then Return type.GetGenericArguments()(0)

        For Each intType As Type In type.GetInterfaces()

            If intType.IsGenericType AndAlso intType.GetGenericTypeDefinition() = GetType(IEnumerable(Of)) Then
                Return intType.GetGenericArguments()(0)
            End If
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' 获取边框线设定
    ''' </summary>
    ''' <param name="style"></param>
    ''' <returns></returns>
    Public Function ConvertFromPredefinedBorderStyle(style As PredefinedBorderStyle) As List(Of KeyValuePair(Of StyleType, Boolean))
        Dim borderStyles As New List(Of KeyValuePair(Of StyleType, Boolean))
        Select Case style
            Case PredefinedBorderStyle.BORDER_ALL
                borderStyles.AddRange({
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_EDGE_LEFT, True),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_EDGE_TOP, True),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_EDGE_RIGHT, True),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_EDGE_BOTTOM, True),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_INSIDE_HORIZONTAL, True),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_INSIDE_VERTICAL, True),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_DIAGONAL_UP, False),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_DIAGONAL_DOWN, False)
                    })
            Case PredefinedBorderStyle.BORDER_INNER
                borderStyles.AddRange({
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_EDGE_LEFT, False),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_EDGE_TOP, False),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_EDGE_RIGHT, False),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_EDGE_BOTTOM, False),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_INSIDE_HORIZONTAL, True),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_INSIDE_VERTICAL, True),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_DIAGONAL_UP, False),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_DIAGONAL_DOWN, False)
                    })
            Case PredefinedBorderStyle.BORDER_OUTER
                borderStyles.AddRange({
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_EDGE_LEFT, True),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_EDGE_TOP, True),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_EDGE_RIGHT, True),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_EDGE_BOTTOM, True),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_INSIDE_HORIZONTAL, False),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_INSIDE_VERTICAL, False),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_DIAGONAL_UP, False),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_DIAGONAL_DOWN, False)
                    })
            Case PredefinedBorderStyle.BORDER_CLEAR
                borderStyles.AddRange({
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_EDGE_LEFT, False),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_EDGE_TOP, False),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_EDGE_RIGHT, False),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_EDGE_BOTTOM, False),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_INSIDE_HORIZONTAL, False),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_INSIDE_VERTICAL, False),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_DIAGONAL_UP, False),
                        New KeyValuePair(Of StyleType, Boolean)(StyleType.ST_BORDER_DIAGONAL_DOWN, False)
                    })
            Case PredefinedBorderStyle.BORDER_CUSTOM
        End Select
        Return borderStyles
    End Function

    ''' <summary>
    ''' <see cref="Excel.XlBordersIndex"/>转换为<see cref="StyleType"/>
    ''' </summary>
    ''' <param name="border"></param>
    ''' <returns></returns>
    Public Function ConvertFromXlBorderIndex(border As Excel.XlBordersIndex) As StyleType
        Dim rst As StyleType
        Select Case border
            Case Excel.XlBordersIndex.xlDiagonalDown
                rst = StyleType.ST_BORDER_DIAGONAL_DOWN
            Case Excel.XlBordersIndex.xlDiagonalUp
                rst = StyleType.ST_BORDER_DIAGONAL_UP
            Case Excel.XlBordersIndex.xlEdgeLeft
                rst = StyleType.ST_BORDER_EDGE_LEFT
            Case Excel.XlBordersIndex.xlEdgeTop
                rst = StyleType.ST_BORDER_EDGE_TOP
            Case Excel.XlBordersIndex.xlEdgeRight
                rst = StyleType.ST_BORDER_EDGE_RIGHT
            Case Excel.XlBordersIndex.xlEdgeBottom
                rst = StyleType.ST_BORDER_EDGE_BOTTOM
            Case Excel.XlBordersIndex.xlInsideHorizontal
                rst = StyleType.ST_BORDER_INSIDE_HORIZONTAL
            Case Excel.XlBordersIndex.xlInsideVertical
                rst = StyleType.ST_BORDER_INSIDE_VERTICAL
        End Select

        Return rst
    End Function

    ''' <summary>
    ''' <see cref="StyleType"/>转换为<see cref="Excel.XlBordersIndex"/>
    ''' </summary>
    ''' <param name="border"></param>
    ''' <returns></returns>
    Public Function ConvertToXlBorderIndex(border As StyleType) As Excel.XlBordersIndex
        Dim rst As Excel.XlBordersIndex
        Select Case border
            Case StyleType.ST_BORDER_DIAGONAL_DOWN
                rst = Excel.XlBordersIndex.xlDiagonalDown
            Case StyleType.ST_BORDER_DIAGONAL_UP
                rst = Excel.XlBordersIndex.xlDiagonalUp
            Case StyleType.ST_BORDER_EDGE_LEFT
                rst = Excel.XlBordersIndex.xlEdgeLeft
            Case StyleType.ST_BORDER_EDGE_TOP
                rst = Excel.XlBordersIndex.xlEdgeTop
            Case StyleType.ST_BORDER_EDGE_RIGHT
                rst = Excel.XlBordersIndex.xlEdgeRight
            Case StyleType.ST_BORDER_EDGE_BOTTOM
                rst = Excel.XlBordersIndex.xlEdgeBottom
            Case StyleType.ST_BORDER_INSIDE_HORIZONTAL
                rst = Excel.XlBordersIndex.xlInsideHorizontal
            Case StyleType.ST_BORDER_INSIDE_VERTICAL
                rst = Excel.XlBordersIndex.xlInsideVertical
        End Select

        Return rst
    End Function

    ''' <summary>
    ''' <see cref="Excel.XlLineStyle"/>转换为<see cref="Pen"/>.DashStyle
    ''' </summary>
    ''' <param name="line"></param>
    ''' <returns></returns>
    Public Function ConvertFromXlLineStyle(line As Excel.XlLineStyle) As Drawing2D.DashStyle
        Dim rst As Drawing2D.DashStyle
        Select Case line
            Case Excel.XlLineStyle.xlContinuous
                rst = Drawing2D.DashStyle.Solid
            Case Excel.XlLineStyle.xlDash
                rst = Drawing2D.DashStyle.Dash
            Case Excel.XlLineStyle.xlDot
                rst = Drawing2D.DashStyle.Dot
            Case Excel.XlLineStyle.xlDashDot
                rst = Drawing2D.DashStyle.DashDot
            Case Excel.XlLineStyle.xlDashDotDot
                rst = Drawing2D.DashStyle.DashDotDot
        End Select

        Return rst
    End Function

    ''' <summary>
    ''' <see cref="Pen"/>.DashStyle转换为<see cref="Excel.XlLineStyle"/>
    ''' </summary>
    ''' <param name="line"></param>
    ''' <returns></returns>
    Public Function ConvertToXlLineStyle(line As Drawing2D.DashStyle) As Excel.XlLineStyle
        Dim rst As Excel.XlLineStyle
        Select Case line
            Case Drawing2D.DashStyle.Solid
                rst = Excel.XlLineStyle.xlContinuous
            Case Drawing2D.DashStyle.Dash
                rst = Excel.XlLineStyle.xlDash
            Case Drawing2D.DashStyle.Dot
                rst = Excel.XlLineStyle.xlDot
            Case Drawing2D.DashStyle.DashDot
                rst = Excel.XlLineStyle.xlDashDot
            Case Drawing2D.DashStyle.DashDotDot
                rst = Excel.XlLineStyle.xlDashDotDot
        End Select

        Return rst
    End Function

    ''' <summary>
    ''' <see cref="Excel.XlBorderWeight"/>转换为<see cref="Pen"/>.Width
    ''' </summary>
    ''' <param name="weight"></param>
    ''' <returns></returns>
    Public Function ConvertFromXlBorderWeight(weight As Excel.XlBorderWeight) As Single
        Dim rst As Single
        Select Case weight
            Case Excel.XlBorderWeight.xlThin
                rst = BORDER_WEIGHT_THIN
            Case Excel.XlBorderWeight.xlMedium
                rst = BORDER_WEIGHT_MEDIUM
            Case Excel.XlBorderWeight.xlThick
                rst = BORDER_WEIGHT_THICK
            Case Else
                rst = BORDER_WEIGHT_THIN
        End Select

        Return rst
    End Function

    ''' <summary>
    ''' 非空判断（用于从数据库读取的字段）
    ''' </summary>
    ''' <param name="val"></param>
    ''' <returns></returns>
    Public Function IsNull(val As Object) As Boolean
        Return TypeOf val Is DBNull OrElse val Is Nothing
    End Function

    Public Function ParseEnum(Of T As Structure)(val As Object) As T?
        If TypeOf val Is DBNull OrElse val Is Nothing Then
            Return Nothing
        Else
            Dim rst As New T
            If [Enum].TryParse(Of T)(val, rst) Then
                Return rst
            Else
                Return Nothing
            End If
        End If
    End Function

    ''' <summary>
    ''' 系统颜色转换为16进制颜色字符串
    ''' </summary>
    ''' <param name="c"></param>
    ''' <returns></returns>
    Public Function GetHexColor(c As Color) As String
        Return "#" & c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2")
    End Function

    ''' <summary>
    ''' 16进制颜色字符串转换为系统颜色
    ''' 例："#FF0000" -> Color.Red
    ''' </summary>
    ''' <param name="c"></param>
    ''' <returns></returns>
    Public Function GetSystemColor(c As String) As Color
        Return ColorTranslator.FromHtml(c)
    End Function

    ''' <summary>
    ''' 根据所在区域获取区域继承关系（用于样式查询）
    ''' </summary>
    ''' <param name="region"></param>
    ''' <returns></returns>
    Public Function GetRegionInheritance(region As RegionType) As RegionType()
        Dim rst As New List(Of RegionType) From {region}
        Select Case region
            Case RegionType.RT_TABLE_HEADER
            Case RegionType.RT_COLUMN_HEADER
            Case RegionType.RT_COLUMN_HEADER_COLNAME
                rst.Add(RegionType.RT_COLUMN_HEADER)
            Case RegionType.RT_COLUMN_HEADER_COMMENT
                rst.Add(RegionType.RT_COLUMN_HEADER)
            Case RegionType.RT_COLUMN_HEADER_PROP
                rst.Add(RegionType.RT_COLUMN_HEADER)
            Case RegionType.RT_COLUMN_HEADER_PK
                rst.Add(RegionType.RT_COLUMN_HEADER)
            Case RegionType.RT_ROW_HEADER
            Case RegionType.RT_ROW_DATA
            Case RegionType.RT_ROW_DATA_NOT_FOUND
        End Select
        rst.Add(RegionType.RT_TABLE_GLOBAL)
        Return rst.ToArray
    End Function

    '''' <summary>
    '''' 判断Excel是否处于编辑状态
    '''' </summary>
    '''' <returns></returns>
    'Public Function IsEditing() As Boolean
    '    If Globals.ThisAddIn.Application.Interactive = False Then Return False
    '    Try
    '        Globals.ThisAddIn.Application.Interactive = False
    '        Globals.ThisAddIn.Application.Interactive = True
    '    Catch
    '        Return True
    '    End Try
    '    Return False
    'End Function

    ''' <summary>
    ''' 判断Excel是否处于编辑状态
    ''' </summary>
    ''' <returns></returns>
    Public Function IsEditing() As Boolean
        Const menuItemType As Integer = 1
        Const newMenuId As Integer = 18
        Dim newMenu As Microsoft.Office.Core.CommandBarControl = Globals.ThisAddIn.Application.CommandBars("Worksheet Menu Bar").FindControl(menuItemType, newMenuId, Type.Missing, Type.Missing, True)
        Return newMenu IsNot Nothing AndAlso Not newMenu.Enabled
    End Function

    Public Function MeasureFontSize(g As Graphics, ff As FontFamily, height As Single) As Font
        Dim fontSize As Single = 12
        Dim rst As Font = Nothing
        While True
            rst = New Font(ff, fontSize)
            Dim h = g.MeasureString(rst.Name, rst).Height
            If Math.Abs(h - height) < 0.1 Then Exit While
            If h > height Then
                fontSize -= 0.05
            Else
                fontSize += 0.05
            End If
        End While

        Return rst
    End Function

    Private _basePath As String
    Public ReadOnly Property GetBasePath() As String
        Get
            If String.IsNullOrEmpty(_basePath) Then
                Dim assemblyInfo As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
                'Dim location As String = assemblyInfo.Location
                Dim uriCodeBase As Uri = New Uri(assemblyInfo.CodeBase)
                _basePath = Path.GetDirectoryName(uriCodeBase.LocalPath.ToString())
            End If
            Return _basePath
        End Get
    End Property

    Public Sub Toast(ParamArray args As String())

        With Environment.OSVersion.Version
            ' Windows 10 (introduced in 10.0.10240.0)
            If .Major >= 10 And .Minor >= 0 And .Build >= 10240 Then
                'Dim builder = New ToastContentBuilder

                'If args.Length > 0 Then
                '    builder.AddText(args(0))
                'End If
                'If args.Length > 1 Then
                '    builder.AddText(args(1))
                'End If
                ''.AddInlineImage(New Uri(Path.Combine(GetBasePath, "logo here"))) ' TODO "Resources\Icon\oracle_mini_32.png"
                'If args.Length > 2 Then
                '    Dim logoUri = New Uri(args(2), UriKind.Relative)
                '    'Dim logoUri = New Uri(Path.Combine(GetBasePath, args(2)))
                '    logger.Debug(logoUri.ToString)
                '    builder.AddAppLogoOverride(logoUri)
                'End If
                ''builder.Show()
                'Try
                '    ToastNotificationManager.CreateToastNotifier("swift-connector").Show(New ToastNotification(builder.GetXml) With {.Tag = "swift-connector"})
                '    'builder.Show()
                'Catch ex As Exception
                '    logger.Error(ex)
                'End Try

                'AddHandler ToastNotificationManagerCompat.OnActivated, Sub(a)
                '                                                           Debug.Print(a.Argument)
                '                                                       End Sub

                Dim toastXml = ToastNotificationManager.GetTemplateContent(ToastTemplateType.ToastImageAndText02)

                Dim textParts = toastXml.GetElementsByTagName("text")
                If args.Length > 0 Then
                    textParts(0).AppendChild(toastXml.CreateTextNode(args(0)))
                End If
                If args.Length > 1 Then
                    textParts(1).AppendChild(toastXml.CreateTextNode(args(1)))
                End If
                If args.Length > 2 Then
                    Dim imageParts = toastXml.GetElementsByTagName("image")
                    imageParts(0).Attributes.GetNamedItem("src").NodeValue = "file:///" & Path.Combine(GetBasePath, args(2))
                End If

                'imageParts(0).Attributes.GetNamedItem("src").NodeValue = "data:image/png;base64," & Convert.ToBase64String(IconResource.zh.ToByteArray(Imaging.ImageFormat.Bmp))

                'Dim audioParts = toastXml.CreateElement("audio")
                'audioParts.SetAttribute("src", "ms-winsoundevent:Notification.Reminder")
                'toastXml.DocumentElement.AppendChild(audioParts)

                'Dim commandParts = toastXml.CreateElement("commands")
                'toastXml.DocumentElement.AppendChild(commandParts)
                'Dim command = toastXml.CreateElement("command")
                'command.SetAttribute("id", "dismiss")
                'command.SetAttribute("arguments", "testdismiss")
                'commandParts.AppendChild(command)

                'AddHandler toast.Activated, Sub(a, obj)
                '                                Debug.Print("Activated")
                '                            End Sub

                ToastNotificationManager.CreateToastNotifier("Swift Connector").Show(New ToastNotification(toastXml) With {
                    .Tag = "swift-connector"
                })
            Else
                MsgBox(args(1), Title:=args(0))
            End If
        End With

    End Sub

    Public DataSourceDic As Dictionary(Of DataSourceType, String) = New Dictionary(Of DataSourceType, String) From {
            {DataSourceType.Oracle, "oracle"},
            {DataSourceType.MySQL, "mysql"},
            {DataSourceType.SqlServer, "sqlserver"},
            {DataSourceType.PostgreSQL, "postgresql"},
            {DataSourceType.SQLite, "sqlite"}
        }

End Module
