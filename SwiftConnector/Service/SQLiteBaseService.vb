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
Imports System.IO

Public Class SQLiteBaseService
    Inherits BaseService

    ''' <summary>
    ''' 连接数据库
    ''' </summary>
    ''' <param name="callback">回调方法（用于通过<see cref="DbCommand"/>和<see cref="DbConnection"/>进行具体的数据库操作）</param>
    Protected Sub doConnect(callback As Action(Of DbCommand, DbConnection))

        'logger.Debug("doConnect start")

        Dim sqlConnectionSb = New SQLiteConnectionStringBuilder With {.DataSource = InternalConfigDatabase}
        'logger.Debug(String.Format("connect to {0}", sqlConnectionSb.ToString()))
        Try
            Using conn As DbConnection = New SQLiteConnection(sqlConnectionSb.ToString())
                'logger.Debug("connected")
                conn.Open()
                Using trans As DbTransaction = conn.BeginTransaction()
                    Try
                        Using comm As DbCommand = conn.CreateCommand()
                            callback.Invoke(comm, conn)
                        End Using
                        trans.Commit()
                    Catch ex As Exception
                        trans.Rollback()
                        Throw
                    End Try
                End Using
            End Using
        Catch ex As Exception
            logger.Error(ex.Message)
            logger.Error(ex.StackTrace)
        End Try

        'logger.Debug("doConnect end")
    End Sub

    ''' <summary>
    ''' 执行读操作
    ''' </summary>
    ''' <param name="callback">回调方法（用于从<see cref="DbDataReader"/>读取数据等操作 ）</param>
    ''' <param name="sql">sql语句</param>
    ''' <param name="parameters">sql参数</param>
    Public Sub ExecuteReader(callback As Action(Of DbDataReader), sql As String, Optional parameters As Object() = Nothing)
        doConnect(Sub(cmd As DbCommand, conn As DbConnection)
                      cmd.CommandText = sql
                      If parameters IsNot Nothing Then
                          For Each parameter In parameters
                              If TypeOf parameter Is DbParameter Then
                                  cmd.Parameters.Add(parameter)
                              ElseIf TypeOf parameter Is DbArrayParameter Then
                                  Dim arrayParameter As DbArrayParameter = parameter
                                  If arrayParameter.ArrayValue.Length > 0 Then
                                      cmd.AddArrayParameters(Of SQLiteParameter)(arrayParameter.ParamName, arrayParameter.ArrayValue)
                                  End If
                              Else
                                  Throw New ArgumentException("Unsupported parameter")
                              End If
                          Next
                      End If
                      'If parameters IsNot Nothing Then
                      '    cmd.Parameters.AddRange(parameters)
                      'End If
                      cmd.Prepare()
                      Using reader = cmd.ExecuteReader()
                          callback.Invoke(reader)
                      End Using
                  End Sub)
    End Sub

    ''' <summary>
    ''' 执行写操作
    ''' </summary>
    ''' <param name="callback">回调方法（用于根据写操作返回值<see cref="Integer"/>进行其他的操作 ）</param>
    ''' <param name="sql">sql语句</param>
    ''' <param name="parameters">sql参数</param>
    Public Sub ExecuteNonQuery(callback As Action(Of Integer), sql As String, Optional parameters As DbParameter() = Nothing)

        ExecuteNonQuery(Sub(rst As List(Of Integer))
                            For Each r In rst
                                callback.Invoke(r)
                            Next
                        End Sub, {sql}, If(parameters Is Nothing, Nothing, {parameters}))
    End Sub

    ''' <summary>
    ''' 执行写操作(多条sql语句)
    ''' </summary>
    ''' <param name="callback">回调方法（用于根据写操作返回值<see cref="List(Of Integer)"/>进行其他的操作 ）</param>
    ''' <param name="sqlArray">sql语句数组</param>
    ''' <param name="parametersArray">sql参数数组</param>
    Public Sub ExecuteNonQuery(callback As Action(Of List(Of Integer)), sqlArray As String(), Optional parametersArray As DbParameter()() = Nothing)
        doConnect(Sub(cmd As DbCommand, conn As DbConnection)
                      Dim rst As New List(Of Integer)
                      For idx = 0 To sqlArray.Length - 1
                          cmd.CommandText = sqlArray(idx)
                          If parametersArray IsNot Nothing AndAlso parametersArray(idx) IsNot Nothing Then
                              cmd.Parameters.AddRange(parametersArray(idx))
                          End If
                          cmd.Prepare()
                          rst.Add(cmd.ExecuteNonQuery())
                      Next
                      callback.Invoke(rst)
                  End Sub)
    End Sub

End Class
