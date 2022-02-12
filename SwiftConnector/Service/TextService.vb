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
Imports System.Data.SQLite

Public Class TextService
    Inherits SQLiteBaseService
    Implements ITextService

    Enum TextCategory
        PRESET = 0
        CUSTOM = 1
    End Enum

    Private Const NO_LOCALE = "-"

    Public Function GetCustomTextByProperty(prop As TextType, Optional lang As String = Nothing, Optional localeIndependent As Boolean = False) As String Implements ITextService.GetCustomTextByProperty
        Return GetText(TextCategory.CUSTOM, prop, lang, localeIndependent)
    End Function

    Public Function GetPresetTextByProperty(prop As TextType, Optional lang As String = Nothing, Optional localeIndependent As Boolean = False) As String Implements ITextService.GetPresetTextByProperty
        Return GetText(TextCategory.PRESET, prop, lang, localeIndependent)
    End Function

    Public Function GetTextByProperty(prop As TextType, Optional lang As String = Nothing, Optional localeIndependent As Boolean = False) As String Implements ITextService.GetTextByProperty
        Dim rst As Object = Nothing

        '查询自定义文本
        rst = GetCustomTextByProperty(prop, lang, localeIndependent)
        If rst Is Nothing Then
            '查询预设文本
            rst = GetPresetTextByProperty(prop, lang, localeIndependent)
        End If
        Return rst
    End Function

    Private Function GetText(cat As TextCategory, prop As TextType, Optional lang As String = Nothing, Optional localeIndependent As Boolean = False) As String
        Dim rst As String = Nothing
        Dim sql As String = "SELECT VAL FROM TEXT_" & [Enum].GetName(cat.GetType, cat) & " WHERE PROP=@PROP AND LOCALE=@LOCALE"
        Dim parameters As New List(Of DbParameter) From {
            New SQLiteParameter("@PROP", Data.DbType.Int32) With {.Value = prop},
            New SQLiteParameter("@LOCALE", Data.DbType.String) With {
                .Value = If(localeIndependent, NO_LOCALE, If(lang Is Nothing, Globals.ThisAddIn.StrLangCode, lang))
            }
        }
        ExecuteReader(Sub(reader As DbDataReader)
                          If reader.Read Then
                              rst = CStr(reader.Item(0))
                          End If
                      End Sub, sql, parameters.ToArray)
        Return rst
    End Function

End Class
