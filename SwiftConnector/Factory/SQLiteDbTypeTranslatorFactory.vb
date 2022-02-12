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

Imports System.Data

Public Class SQLiteDbTypeTranslatorFactory
    Implements IDbTypeTranslatorFactory

    Public Function Format(dbType As DbType, val As String) As Object Implements IDbTypeTranslatorFactory.Format
        Dim rst As Object = Nothing
        Select Case dbType
            Case Else
                ' 无需转换
                rst = val
        End Select
        Return rst
    End Function

    Public Function Translate(independantDbType As String) As DbType Implements IDbTypeTranslatorFactory.Translate
        Dim rst As DbType

        Dim strDbType = independantDbType
        Select Case strDbType
            Case "TEXT"
                rst = DbType.String
            Case "INTEGER"
                rst = DbType.Int64
            Case "BLOB"
                rst = DbType.Object
            Case "REAL"
                rst = DbType.Double
                'Case "NUMERIC"
                '    rst = DbType.VarNumeric
            Case "DATETIME"
                rst = DbType.DateTime
            Case Else
                rst = DbType.String
        End Select

        Return rst
    End Function

End Class
