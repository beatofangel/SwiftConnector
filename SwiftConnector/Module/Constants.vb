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

Module Constants

#Region "Ribbon"
    Public Const BUTTON_TRUNCATE_SCREEN_TIP = "DDL <TRUNCATE>"
    Public Const BUTTON_DELETE_SCREEN_TIP = "DML <DELETE>"

    Public Const BUTTON_TRUNCATE_LABEL = "Truncate"
    Public Const BUTTON_DELETE_LABEL = "Delete"

    Public Const BUTTON_TRUNCATE_IMAGE = "TableDelete"
    Public Const BUTTON_DELETE_IMAGE = "TableDeleteRows"

    Public Const BUTTON_INSERT_NORMAL_IMAGE = "TableInsert"
    Public Const BUTTON_INSERT_SWIFT_IMAGE = "TableInsert"

    Public Const CHECKED_IMAGE As String = "FormControlCheckBox"
    Public Const UNCHECKED_IMAGE As String = "ColorWhite"
    Public Const SYNCHRONIZE_DATA As String = "BusinessDataSyncWithLocalCache"

    Public Const PROTECTED_MODE_IMAGE As String = "ColumnActionsReadOnly"
    Public Const NORMAL_MODE_IMAGE As String = "BorderDrawMenu"

    Public Const KEY_CTRL As String = "^"
    Public Const KEY_SHIFT As String = "+"
    Public Const KEY_ALT As String = "%"
    Public Const KEY_ENTER As String = "~"
#End Region

#Region "Custom Task Pane"
    'Public Const TILE_FONT_SIZE = 14.0F
    'Public Const TILE_FONT_ZOOM = 8.0F / 7.0F
    Public Const TILE_FONT_SIZE = 18.0F
    Public Const TILE_FONT_ZOOM = 1.0F
    Public Const BORDER_WEIGHT_THIN = 1
    Public Const BORDER_WEIGHT_MEDIUM = 2
    Public Const BORDER_WEIGHT_THICK = 4
    Public Const STYLE_LINE_STYLE = "LineStyle"
    Public Const STYLE_COLOR = "Color"
    Public Const STYLE_WEIGHT = "Weight"
#End Region
End Module
