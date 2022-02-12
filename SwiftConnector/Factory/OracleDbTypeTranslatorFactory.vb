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
Imports System.Text.RegularExpressions
Imports Oracle.ManagedDataAccess.Client

Public Class OracleDbTypeTranslatorFactory
    Implements IDbTypeTranslatorFactory

    Public Function Format(dbType As DbType, val As String) As Object Implements IDbTypeTranslatorFactory.Format
        Dim rst As Object = Nothing
        Select Case dbType
            Case DbType.Date, DbType.DateTime
                Dim dt As Date
                If Date.TryParse(val, dt) Then
                    rst = dt
                End If
            Case DbType.Binary ' for RAW type
                If val.Length Mod 2 = 0 Then
                    rst = Enumerable.Range(0, val.Length / 2).Select(Function(x) Convert.ToByte(val.Substring(x * 2, 2), 16)).ToArray
                Else
                    Throw New ArgumentException("The length of the hexadecimal string to be converted to RAW type must be a multiple of 2!")
                End If
            Case Else
                ' TODO 完善对全部类型的精确转换
                rst = val
        End Select
        Return rst
    End Function

    Public Function Translate(independantDbType As String) As DbType Implements IDbTypeTranslatorFactory.Translate
        Dim oDbType As OracleDbType
        Dim rst As DbType

        Dim strDbType = independantDbType
        ' 预处理timestamp类型
        If strDbType.StartsWith("TIMESTAMP") Then
            Dim regex As New Regex("^(TIMESTAMP(\(\d\))( WITH( LOCAL)? TIME ZONE)?)$", RegexOptions.IgnoreCase)
            strDbType = regex.Replace(strDbType, New MatchEvaluator(Function(match)
                                                                        Dim rep As String
                                                                        If match.Groups(4).Success Then
                                                                            rep = [Enum].GetName(GetType(OracleDbType), OracleDbType.TimeStampLTZ)
                                                                        ElseIf match.Groups(3).Success Then
                                                                            rep = [Enum].GetName(GetType(OracleDbType), OracleDbType.TimeStampTZ)
                                                                        ElseIf match.Groups(1).Success Then
                                                                            rep = [Enum].GetName(GetType(OracleDbType), OracleDbType.TimeStamp)
                                                                        Else
                                                                            Throw New ArgumentException(strDbType & " is not currently supported!")
                                                                        End If
                                                                        Return rep
                                                                    End Function))
        End If

        If [Enum].TryParse(strDbType, True, oDbType) Then
            Select Case oDbType
                Case OracleDbType.BFile
                    rst = DbType.Object
                Case OracleDbType.Blob
                    rst = DbType.Object
                Case OracleDbType.Byte
                    rst = DbType.Byte
                Case OracleDbType.Char
                    rst = DbType.StringFixedLength
                Case OracleDbType.Clob
                    rst = DbType.Object
                Case OracleDbType.Date
                    rst = DbType.Date
                Case OracleDbType.Decimal
                    rst = DbType.Decimal
                Case OracleDbType.Double
                    rst = DbType.Double
                Case OracleDbType.Int16
                    rst = DbType.Int16
                Case OracleDbType.Int32
                    rst = DbType.Int32
                Case OracleDbType.Int64
                    rst = DbType.Int64
                Case OracleDbType.IntervalDS
                    'rst = DbType.TimeSpan
                    Throw New NotImplementedException(strDbType & " is not currently supported!")
                Case OracleDbType.IntervalYM
                    rst = DbType.Int64
                Case OracleDbType.Long
                    rst = DbType.String
                Case OracleDbType.LongRaw
                    rst = DbType.Binary
                Case OracleDbType.NChar
                    rst = DbType.StringFixedLength
                Case OracleDbType.NClob
                    rst = DbType.Object
                Case OracleDbType.NVarchar2
                    rst = DbType.String
                Case OracleDbType.Raw
                    rst = DbType.Binary
                Case OracleDbType.RefCursor
                    rst = DbType.Object
                Case OracleDbType.Single
                    rst = DbType.Single
                Case OracleDbType.TimeStamp
                    rst = DbType.DateTime
                Case OracleDbType.TimeStampLTZ
                    rst = DbType.DateTime
                Case OracleDbType.TimeStampTZ
                    rst = DbType.DateTime
                Case OracleDbType.Varchar2
                    rst = DbType.String
                Case OracleDbType.XmlType
                    rst = DbType.String
            End Select
        End If

        Return rst
    End Function
End Class
