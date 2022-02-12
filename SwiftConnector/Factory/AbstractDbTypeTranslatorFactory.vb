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
''' 数据类型转换抽象工厂
''' </summary>
Public Class AbstractDbTypeTranslatorFactory

    ''' <summary>
    ''' <para>生成数据类型转换工厂</para> 
    ''' <para>注意：扩展数据库类型支持时，必须实现指定<see cref="DataSourceType"/>分支，否则将抛出<see cref="NotImplementedException"/>异常</para>
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function CreateFactory() As IDbTypeTranslatorFactory
        Dim rst As IDbTypeTranslatorFactory = Nothing
        Dim curDs = Globals.ThisAddIn.CurDataSource()
        Select Case curDs.Type
            Case DataSourceType.Oracle
                rst = New OracleDbTypeTranslatorFactory()
            Case DataSourceType.MySQL
                rst = New MySQLDbTypeTranslatorFactory()
            Case DataSourceType.PostgreSQL
                Throw New NotImplementedException("PostgreSQL is not currently supported!")
            Case DataSourceType.SqlServer
                Throw New NotImplementedException("SqlServer is not currently supported!")
            Case DataSourceType.SQLite
                rst = New SQLiteDbTypeTranslatorFactory()
            Case Else
                Throw New NotImplementedException(curDs.Name & " is not currently supported!")
        End Select

        Return rst
    End Function
End Class
