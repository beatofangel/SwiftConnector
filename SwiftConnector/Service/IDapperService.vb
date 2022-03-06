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

Public Interface IDapperService
    'Function ExecuteReader(sql As String, Optional param As Object = Nothing) As IDataReader

    'Function Execute(sql As String, Optional param As Object = Nothing) As Integer

    Sub ExecuteReader(callback As Action(Of IDataReader), sql As String, Optional param As Object = Nothing)

    Sub Execute(callback As Action(Of Integer), sql As String, Optional param As Object = Nothing)

    Function ExecuteReader(callback As Action(Of IDataReader), txHash As Integer?, sql As String, Optional param As Object = Nothing) As Integer

    Function Execute(callback As Action(Of Integer), txHash As Integer?, sql As String, Optional param As Object = Nothing) As Integer

    'Function GetParameterPrefix() As String
End Interface
