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
Imports Dapper

Public Class DapperService
    Inherits BaseService
    Implements IDapperService

    'Public Function ExecuteReader(sql As String, Optional param As Object = Nothing) As IDataReader Implements IDapperService.ExecuteReader
    '    Dim rst As IDataReader
    '    Using db As IDbConnection = GetOpenConnection(), tx As IDbTransaction = db.BeginTransaction
    '        rst = db.ExecuteReader(sql, param)
    '    End Using

    '    Return rst
    'End Function

    'Public Function Execute(sql As String, Optional param As Object = Nothing) As Integer Implements IDapperService.Execute
    '    Dim rst As Integer
    '    Using db As IDbConnection = GetOpenConnection(), tx As IDbTransaction = db.BeginTransaction
    '        rst = db.Execute(sql, param)
    '    End Using

    '    Return rst
    'End Function

    Public Sub ExecuteReader(callback As Action(Of IDataReader), sql As String, Optional param As Object = Nothing) Implements IDapperService.ExecuteReader
        logger.Debug("ExecuteReader Start")
        Using db As IDbConnection = GetOpenConnection()
            Using tx As IDbTransaction = db.BeginTransaction
                callback.Invoke(db.ExecuteReader(sql, param, transaction:=tx))
            End Using
        End Using
        logger.Debug("ExecuteReader End")
    End Sub

    Public Sub Execute(callback As Action(Of Integer), sql As String, Optional param As Object = Nothing) Implements IDapperService.Execute
        logger.Debug("Execute Start")
        Using db As IDbConnection = GetOpenConnection()
            Using tx As IDbTransaction = db.BeginTransaction
                callback.Invoke(db.Execute(sql, param:=param, transaction:=tx))
                tx.Commit()
            End Using
        End Using
        logger.Debug("Execute End")
    End Sub

    Public Function ExecuteReader(callback As Action(Of IDataReader), txHash As Integer?, sql As String, Optional param As Object = Nothing) As Integer Implements IDapperService.ExecuteReader
        logger.Debug("ExecuteReader Start")
        Return beginTx(Sub(db, tx)
                           callback.Invoke(db.ExecuteReader(sql, param, transaction:=tx))
                           logger.Debug("ExecuteReader End")
                       End Sub, txHash)
    End Function

    Public Function Execute(callback As Action(Of Integer), txHash As Integer?, sql As String, Optional param As Object = Nothing) As Integer Implements IDapperService.Execute
        logger.Debug("Execute Start")
        Return beginTx(Sub(db, tx)
                           callback.Invoke(db.Execute(sql, param:=param, transaction:=tx))
                           logger.Debug("Execute End")
                       End Sub, txHash)
    End Function

    'Public Function GetParameterPrefix() As String Implements IDapperService.GetParameterPrefix
    '    Using db As DbConnection = GetOpenConnection()
    '        Dim tbl = db.GetSchema(DbMetaDataCollectionNames.DataSourceInformation)
    '        Dim markerFormat As String = tbl.Rows.Item(0).Item("ParameterMarkerFormat")
    '        Dim temp = markerFormat.Split({"{0}"}, StringSplitOptions.RemoveEmptyEntries)
    '        If temp IsNot Nothing Then
    '            Return temp(0)
    '        End If
    '    End Using
    '    Return String.Empty
    'End Function
End Class
