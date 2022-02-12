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

''' <summary>
''' 属性字段工厂类
''' </summary>
Public Class PropertyColumnFactory

    ''' <summary>
    ''' <para>生成属性</para> 
    ''' <para>注意：扩展数据库类型支持时，必须实现指定<see cref="DataSourceType"/>分支，否则将抛出<see cref="NotImplementedException"/>异常</para>
    ''' </summary>
    ''' <param name="DataType"></param>
    ''' <param name="DataLength"></param>
    ''' <param name="DataPrecision"></param>
    ''' <param name="DataScale"></param>
    ''' <param name="Nullable"></param>
    ''' <returns></returns>
    Public Shared Function CreateProperty(DataType As String, DataLength As Integer, DataPrecision As Integer, DataScale As Integer, Nullable As String) As String
        Dim rst As String
        Select Case Globals.ThisAddIn.CurDataSource.Type
            Case DataSourceType.Oracle
                ' TODO 完善对全部类型的精确显示
                Select Case DataType
                    Case "DATE"
                        rst = String.Format("{0}[{1}]", DataType, Nullable)
                    Case "VARCHAR2", "NVARCHAR2"
                        rst = String.Format("{0}({1})[{2}]", DataType, DataLength, Nullable)
                    Case "NUMBER"
                        rst = String.Format("{0}({1},{2})[{3}]", DataType, DataPrecision, DataScale, Nullable)
                    Case "RAW"
                        rst = String.Format("{0}({1})[{2}]", DataType, DataLength, Nullable)
                    Case Else   ' 包含 "TIMESTAMP(X)","TIMESTAMP(X) WITH TIME ZONE","TIMESTAMP(X) WITH LOCAL TIME ZONE"
                        rst = String.Format("{0}[{1}]", DataType, Nullable)
                End Select
            Case DataSourceType.MySQL
                Select Case DataType.ToUpper
                    Case "VARCHAR"
                        rst = String.Format("{0}({1})[{2}]", DataType, DataLength, Nullable)
                    Case "DECIMAL"
                        rst = String.Format("{0}({1},{2})[{3}]", DataType, DataPrecision, DataScale, Nullable)
                    Case "DATETIME", "TIMESTAMP"
                        If DataLength > 0 Then
                            rst = String.Format("{0}({1})[{2}]", DataType, DataLength, Nullable)
                        Else
                            rst = String.Format("{0}[{1}]", DataType, Nullable)
                        End If
                    Case Else
                        rst = String.Format("{0}[{1}]", DataType, Nullable)
                End Select
            Case DataSourceType.PostgreSQL
                Throw New NotImplementedException("PostgreSQL is not currently supported!")
            Case DataSourceType.SqlServer
                Throw New NotImplementedException("SqlServer is not currently supported!")
            Case DataSourceType.SQLite
                Select Case DataType.ToUpper
                    Case Else
                        rst = String.Format("{0}[{1}]", DataType, Nullable)
                End Select
            Case Else
                Throw New NotImplementedException(Globals.ThisAddIn.CurDataSource.Name & " is not currently supported!")
        End Select

        Return rst
    End Function
End Class
