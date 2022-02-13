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

Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.Common
Imports System.Drawing
Imports System.Drawing.Design

Public Module Definitions

    Public Enum EditMode
        Add = 0
        Edit = 1
        Delete = 2
    End Enum

    Public Enum DeleteMode
        Truncate = 0
        Delete = 1
    End Enum

    Public Enum ExecuteMode
        Normal = 0
        Swift = 1
    End Enum

    Public Enum DataSourceType
        Unknown = 0
        Oracle = 1
        MySQL = 2
        SqlServer = 3
        PostgreSQL = 4
        SQLite = 5
    End Enum

    Public Enum OperateMode
        Normal = 0
        [Protected] = 1
    End Enum

    Public Enum RenderMode
        Memory = 0
        Excel = 1
    End Enum

    Public Enum RegionType
        RT_TABLE_GLOBAL = 0
        RT_TABLE_HEADER = 1
        RT_COLUMN_HEADER = 2
        RT_COLUMN_HEADER_COLNAME = 3
        RT_COLUMN_HEADER_COMMENT = 4
        RT_COLUMN_HEADER_PROP = 5
        RT_COLUMN_HEADER_PK = 6
        RT_ROW_HEADER = 7
        RT_ROW_DATA = 8
        RT_ROW_DATA_NOT_FOUND = 9
    End Enum

    Public Enum StyleType
        ST_FORMAT = 0
        ST_FONT = 1
        ST_INTERIOR = 2
        ST_BORDER_DIAGONAL_DOWN = 3
        ST_BORDER_DIAGONAL_UP = 4
        ST_BORDER_EDGE_TOP = 5
        ST_BORDER_EDGE_BOTTOM = 6
        ST_BORDER_EDGE_LEFT = 7
        ST_BORDER_EDGE_RIGHT = 8
        ST_BORDER_INSIDE_HORIZONTAL = 9
        ST_BORDER_INSIDE_VERTICAL = 10
        ST_COL_PROPS_DISPLAY = 11
        ST_RECORD_LIMIT = 12
        ST_DELETE_MODE = 13
    End Enum

    Public Enum TextType
        TT_NO_DATA_FOUND = 0
        TT_ROW_HEADER_TABLE = 1
        TT_ROW_HEADER_COLUMN = 2
        TT_ROW_HEADER_COMMENT = 3
        TT_ROW_HEADER_PROP = 4
        TT_RB_QUERY = 5
        TT_RB_SAVE = 6
        TT_RB_DELETE = 7
        TT_RB_TRUNCATE = 8
        TT_RB_GRP_OPERATIONS = 9
        TT_RB_DS_MANAGEMENT = 10
        TT_RB_GRP_DATASOURCE = 11
        TT_RB_STYLE_SETTINGS = 12
        TT_RB_SHOW_PROPS = 13
        TT_RB_RESET_STYLE = 14
        TT_RB_GRP_SETTINGS = 15
        TT_RB_GRP_LANGUAGE = 16
        TT_RB_ABOUT = 17
        TT_RB_GRP_OTHER = 18
        TT_RB_DELETE_MODE = 19
        TT_RB_TRUNCATE_MODE = 20
        TT_RB_RESET_LANG = 21
        TT_RB_TAB = 22
        TT_RB_AUTO_FIT_COLUMNS = 23
        TT_RB_SAVE_SWIFT = 24
        TT_RB_PROTECTED_MODE = 25
        TT_RB_NORMAL_MODE = 26
        TT_RB_IMPORT = 27
        TT_TP_SETTINGS = 28
        TT_MSG_UNSUPPORTED_LANG_1 = 1001
        TT_MSG_UNSUPPORTED_LANG_2 = 1002
    End Enum

    Public Enum Locale
        en_US = 1033
        zh_CN = 2052
        ja_JP = 1041
    End Enum

    Public Enum PredefinedBorderStyle
        BORDER_CLEAR = 0
        BORDER_ALL = 1
        BORDER_INNER = 2
        BORDER_OUTER = 3
        BORDER_CUSTOM = 4
    End Enum

    Public Class Response
        Public Sub New(success As Boolean, api As String, Optional data As Object = Nothing, Optional message As String = Nothing)
            Me.Success = success
            Me.Api = api
            Me.Data = data
            Me.Message = message
        End Sub

        Public Property Api As String
        Public Property Success As Boolean
        Public Property Data As Object
        Public Property Message As String
    End Class

    Public Class PredefinedBorder
        Public Weight As Excel.XlBorderWeight
        Public Color As Color
        Public Style As Excel.XlLineStyle
        Public PresetStyle As PredefinedBorderStyle

        Public Sub New()
            Me.PresetStyle = PredefinedBorderStyle.BORDER_CLEAR
            Me.Style = Excel.XlLineStyle.xlLineStyleNone
        End Sub

        Public Sub New(presetStyle As PredefinedBorderStyle)
            Me.Style = Excel.XlLineStyle.xlContinuous
            Me.Weight = Excel.XlBorderWeight.xlThin
            Me.Color = Color.Black
            Me.PresetStyle = presetStyle
        End Sub

        Public Sub New(presetStyle As PredefinedBorderStyle, color As Color)
            Me.Style = Excel.XlLineStyle.xlContinuous
            Me.Weight = Excel.XlBorderWeight.xlThin
            Me.Color = color
            Me.PresetStyle = presetStyle
        End Sub
        Public Sub New(presetStyle As PredefinedBorderStyle, color As Color, weight As Excel.XlBorderWeight)
            Me.Style = Excel.XlLineStyle.xlContinuous
            Me.Weight = weight
            Me.Color = color
            Me.PresetStyle = presetStyle
        End Sub
    End Class

    '''' <summary>
    '''' 支持的Oracle数据类型
    '''' </summary>
    'Public Enum OracleDataType
    '    [CHAR]
    '    NCHAR
    '    NVARCHAR2
    '    VARCHAR
    '    VARCHAR2
    '    NUMBER
    '    [DATE]
    '    TIMESTAMP
    '    RAW
    '    CLOB
    '    NCLOB
    '    BLOB
    'End Enum

    Public Structure Range
        Public Sub New(row As Integer, column As Integer)
            Me.row = row
            Me.column = column
        End Sub

        Public row As Integer
        Public column As Integer
    End Structure

    Public Class DbArrayParameter
        Private _paramName As String
        Private _type As DbType
        Private _arrayValue As New Object()

        Public Sub New(paramName As String, params As Object(), Optional type As DbType? = Nothing)
            _paramName = paramName
            _type = type
            _arrayValue = params
        End Sub

        Public Property ParamName As String
            Get
                Return _paramName
            End Get
            Set(value As String)
                _paramName = value
            End Set
        End Property

        Public Property Type As DbType
            Get
                Return _type
            End Get
            Set(value As DbType)
                _type = value
            End Set
        End Property

        Public Property ArrayValue As Object()
            Get
                Return _arrayValue
            End Get
            Set(value As Object())
                _arrayValue = value
            End Set
        End Property

    End Class

#Region "EventArgs"

    Public Class CheckedEventArgs
        Inherits CancelEventArgs

        Public Property Checked As Boolean

        Public Sub New(checked As Boolean)
            Me.Checked = checked
        End Sub
    End Class

    Public Class ColorChangedEventArgs
        Inherits CancelEventArgs

        Public Property Color As Color

        Public Sub New(color As Color)
            Me.Color = color
        End Sub
    End Class

    Public Class BorderPenEventArgs
        Inherits EventArgs

        Public Property Pen As Drawing.Pen

        Public Sub New(pen As Drawing.Pen)
            Me.Pen = pen.Clone
        End Sub
    End Class

    Public Class BorderStyleChangedEventArgs
        Inherits EventArgs

        Public Property Style As PredefinedBorderStyle

        Public Sub New(style As PredefinedBorderStyle)
            Me.Style = style
        End Sub
    End Class

    Public Class ClipboardEventArgs
        Inherits EventArgs

        Public Property ClipboardText As String

        Public Sub New(clipboardText As String)
            Me.ClipboardText = clipboardText
        End Sub
    End Class

#End Region

    Public Class BorderGrid

        Private offset As Integer = 0
        Private _left As Drawing.Point
        Private _top As Drawing.Point
        Private _topLeft As Drawing.Point
        Private _topRight As Drawing.Point
        Private _bottom As Drawing.Point
        Private _bottomLeft As Drawing.Point
        Private _bottomRight As Drawing.Point
        Private _right As Drawing.Point
        Private location As Drawing.Point

        Public ReadOnly Property Left As Drawing.Point
            Get
                Return New Drawing.Point(_left.X + offset, _left.Y)
            End Get
        End Property
        Public ReadOnly Property Top As Drawing.Point
            Get
                Return New Point(_top.X, _top.Y + offset)
            End Get
        End Property
        Public ReadOnly Property TopLeft As Drawing.Point
            Get
                Return New Point(_topLeft.X + offset, _topLeft.Y + offset)
            End Get
        End Property
        Public ReadOnly Property TopRight As Drawing.Point
            Get
                Return New Point(_topRight.X - offset, _topRight.Y + offset)
            End Get
        End Property
        Public ReadOnly Property Bottom As Drawing.Point
            Get
                Return New Point(_bottom.X, _bottom.Y - offset)
            End Get
        End Property
        Public ReadOnly Property BottomLeft As Drawing.Point
            Get
                Return New Point(_bottomLeft.X + offset, _bottomLeft.Y - offset)
            End Get
        End Property
        Public ReadOnly Property BottomRight As Drawing.Point
            Get
                Return New Point(_bottomRight.X - offset, _bottomRight.Y - offset)
            End Get
        End Property
        Public ReadOnly Property Right As Drawing.Point
            Get
                Return New Point(_right.X - offset, _right.Y)
            End Get
        End Property

        Public ReadOnly Property Left2 As Drawing.Point
            Get
                Return New Drawing.Point(location.X + Left.X, location.Y + Left.Y)
            End Get
        End Property
        Public ReadOnly Property Top2 As Drawing.Point
            Get
                Return New Point(location.X + Top.X, location.Y + Top.Y)
            End Get
        End Property
        Public ReadOnly Property TopLeft2 As Drawing.Point
            Get
                Return New Point(location.X + TopLeft.X, location.Y + TopLeft.Y)
            End Get
        End Property
        Public ReadOnly Property TopRight2 As Drawing.Point
            Get
                Return New Point(location.X + TopRight.X, location.Y + TopRight.Y)
            End Get
        End Property
        Public ReadOnly Property Bottom2 As Drawing.Point
            Get
                Return New Point(location.X + Bottom.X, location.Y + Bottom.Y)
            End Get
        End Property
        Public ReadOnly Property BottomLeft2 As Drawing.Point
            Get
                Return New Point(location.X + BottomLeft.X, location.Y + BottomLeft.Y)
            End Get
        End Property
        Public ReadOnly Property BottomRight2 As Drawing.Point
            Get
                Return New Point(location.X + BottomRight.X, location.Y + BottomRight.Y)
            End Get
        End Property
        Public ReadOnly Property Right2 As Drawing.Point
            Get
                Return New Point(location.X + Right.X, location.Y + Right.Y)
            End Get
        End Property

        ''' <summary>
        ''' 根据线宽设置偏移量，用于绘制边框线
        ''' </summary>
        ''' <param name="rect"></param>
        ''' <param name="offset"></param>
        Public Sub New(rect As Drawing.Rectangle, offset As Integer)
            Me.location = rect.Location
            Me.offset = offset
            _left = New Point(0, rect.Height / 2 - 1)
            _top = New Point(rect.Width / 2 - 1, 0)
            _topLeft = New Point(0, 0)
            _topRight = New Point(rect.Width - 2, 0)
            _bottom = New Point(rect.Width / 2 - 1, rect.Height - 2)
            _bottomLeft = New Point(0, rect.Height - 2)
            _bottomRight = New Point(rect.Width - 2, rect.Height - 2)
            _right = New Point(rect.Width - 2, rect.Height / 2 - 1)
        End Sub

        ''' <summary>
        ''' 不根据线宽设置偏移量，用于点击边框线的判定
        ''' </summary>
        ''' <param name="rect"></param>
        Public Sub New(rect As Drawing.Rectangle)
            Me.location = rect.Location
            _left = New Point(0, rect.Height / 2)
            _top = New Point(rect.Width / 2, 0)
            _topLeft = New Point(0, 0)
            _topRight = New Point(rect.Width, 0)
            _bottom = New Point(rect.Width / 2, rect.Height)
            _bottomLeft = New Point(0, rect.Height)
            _bottomRight = New Point(rect.Width, rect.Height)
            _right = New Point(rect.Width, rect.Height / 2)
        End Sub

    End Class

End Module
