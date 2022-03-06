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

Imports System.Collections.Concurrent
Imports System.Data
Imports System.Transactions
Imports log4net

Public MustInherit Class BaseService
    Protected Shared logger As ILog = LogManager.GetLogger(Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Protected Shared txStore As New ConcurrentDictionary(Of Integer, IDbTransaction)

    Protected Function beginTx(callback As Action(Of IDbConnection, IDbTransaction), Optional txHash As Integer? = Nothing) As Integer
        Dim txKey = txHash
        Using db As IDbConnection = GetOpenConnection()
            If txKey Is Nothing Or Not txStore.ContainsKey(txKey) Then
                Dim tx As IDbTransaction = db.BeginTransaction
                txKey = tx.GetHashCode
                txStore.TryAdd(txKey, tx)
            End If
            callback.Invoke(db, txStore(txKey))
        End Using
        Return txKey
    End Function

    Protected Function GetOpenConnection() As IDbConnection
        Dim db As IDbConnection = GetCurrentConnection()
        logger.Debug(String.Format("connecting to {0}", db.ConnectionString))
        db.Open()
        logger.Debug(String.Format("connected to {0}", db.ConnectionString))
        Return db
    End Function

End Class
