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
Imports System.Data.Common
Imports System.Data.SQLite
Imports MySqlConnector
Imports Oracle.ManagedDataAccess.Client

''' <summary>
''' 数据库连接工厂
''' </summary>
Public Class DbConnectionFactory

    ''' <summary>
    ''' <para>生成数据库连接</para> 
    ''' <para>注意：扩展数据库类型支持时，必须实现指定<see cref="DataSourceType"/>分支，否则将抛出<see cref="NotImplementedException"/>异常</para>
    ''' </summary>
    ''' <param name="ds"></param>
    ''' <returns></returns>
    Public Shared Function CreateConnection(ds As DataSource) As IDbConnection
        Dim rst As IDbConnection = Nothing
        Dim connStr As String = Nothing
        Select Case ds.Type
            Case DataSourceType.Oracle
                connStr = "User id=" & ds.Username & ";Password=" & ds.Password & ";Data Source=" &
                    "//" & ds.Datasource & If(String.IsNullOrEmpty(ds.Port), "", ":" & ds.Port) & If(String.IsNullOrEmpty(ds.Database), "", "/" & ds.Database)
                'connStr = "User id=" & ds.Username & ";Password=" & ds.Password & ";Data Source=" &
                '    "(DESCRIPTION=(ADDRESS=(PROTOCOL=tcp)" &
                '    "(HOST=" & ds.Datasource & ")(PORT=" & If(String.IsNullOrEmpty(ds.Port), "1521", ds.Sid) & "))(CONNECT_DATA=" &
                '    "(SERVICE_NAME=" & If(String.IsNullOrEmpty(ds.Sid), "XE", ds.Sid) & ")))"
                'connStr = "Password=" & ds.Password & ";User ID=" & ds.Username & ";Data Source=" & ds.Datasource
                rst = New OracleConnection(connStr)
            Case DataSourceType.MySQL
                Dim connBuilder = New MySqlConnectionStringBuilder() With {
                    .Server = ds.Datasource,
                    .Port = If(String.IsNullOrEmpty(ds.Port), 3306, Integer.Parse(ds.Port)),
                    .Database = ds.Database,
                    .UserID = ds.Username,
                    .Password = ds.Password,
                    .SslMode = MySqlSslMode.Preferred
                }
                rst = New MySqlConnection(connBuilder.ConnectionString)
                ' fixed [The transaction associated with this command is not the connection’s active transaction] problem by using "IgnoreCommandTransaction=true" option
                'connStr = "SERVER=" & curDs.Datasource & ";DATABASE=" & curDs.Database & ";USER=" & curDs.Username & ";PASSWORD=" & curDs.Password & ";IgnoreCommandTransaction=true"
                'connStr = "Server=" & ds.Datasource & ";Port=" & If(String.IsNullOrEmpty(ds.Port), "3306", ds.Sid) & ";Database=" & ds.Database & ";Uid=" & ds.Username & ";Pwd=" & ds.Password & ";SslMode=none"
                'rst = New MySqlConnection(connStr)
            Case DataSourceType.PostgreSQL
                Throw New NotImplementedException("PostgreSQL is not currently supported!")
            Case DataSourceType.SqlServer
                Throw New NotImplementedException("SqlServer is not currently supported!")
            Case DataSourceType.SQLite
                Dim connBuilder = New SQLiteConnectionStringBuilder() With {
                    .DataSource = ds.Datasource,
                    .Pooling = True,
                    .Version = 3,
                    .FailIfMissing = True
                }
                If Not String.IsNullOrEmpty(ds.Password) Then
                    connBuilder.Password = ds.Password
                End If
                rst = New SQLiteConnection(connBuilder.ConnectionString)
                'connStr = "Data Source=" & ds.Datasource & ";Version=3;Pooling=True;Mode=ReadWrite;" & If(String.IsNullOrEmpty(ds.Password), "", "Password=" & ds.Password)
                'rst = New SQLiteConnection(connStr)
            Case Else
                Throw New NotImplementedException(ds.Name & " is not currently supported!")
        End Select

        Return rst
    End Function

End Class
