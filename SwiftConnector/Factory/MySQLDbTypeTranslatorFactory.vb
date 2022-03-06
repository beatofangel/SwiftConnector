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
Imports MySqlConnector

Public Class MySQLDbTypeTranslatorFactory
    Inherits BaseService
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
        Dim myDbType As MySqlDbType
        Dim rst As DbType

        Dim strDbType = independantDbType
        If strDbType = "int" Then
            strDbType = "int32"
        End If
        If [Enum].TryParse(strDbType, True, myDbType) Then
            rst = Convert2DbType(myDbType)
        Else
            logger.Debug("type translate failed from type [" & strDbType & "]")
        End If

        Return rst
    End Function

    Public Function Parse(dbType As String) As DbType Implements IDbTypeTranslatorFactory.Parse
        Return Convert2DbType([Enum].Parse(GetType(MySqlDbType), dbType, True))
    End Function

    Private Function Convert2DbType(myDbType As MySqlDbType) As DbType
        Dim rst As DbType
        Select Case myDbType
            Case MySqlDbType.Bool
                Throw New NotImplementedException(myDbType.ToString & " is not currently supported!")
            Case MySqlDbType.[Decimal]
                rst = DbType.Decimal
            Case MySqlDbType.[Byte]
                rst = DbType.Byte
            Case MySqlDbType.Int16
                rst = DbType.Int16
            Case MySqlDbType.Int32
                rst = DbType.Int32
            Case MySqlDbType.Float
                rst = DbType.Single
            Case MySqlDbType.[Double]
                rst = DbType.Double
            Case MySqlDbType.Null
                Throw New NotImplementedException(myDbType.ToString & " is not currently supported!")
            Case MySqlDbType.Timestamp
                rst = DbType.Date
            Case MySqlDbType.Int64
                rst = DbType.Int64
            Case MySqlDbType.Int24
                Throw New NotImplementedException(myDbType.ToString & " is not currently supported!")
            Case MySqlDbType.[Date]
                rst = DbType.Date
            Case MySqlDbType.Time
                rst = DbType.Date
            Case MySqlDbType.DateTime
                rst = DbType.Date
            Case MySqlDbType.Year
                rst = DbType.UInt16
            Case MySqlDbType.Newdate
                Throw New NotImplementedException(myDbType.ToString & " is not currently supported!")
            Case MySqlDbType.VarString
                rst = DbType.String
            Case MySqlDbType.Bit
                rst = DbType.Binary
            Case MySqlDbType.JSON
                rst = DbType.Object
            Case MySqlDbType.NewDecimal
                Throw New NotImplementedException(myDbType.ToString & " is not currently supported!")
            Case MySqlDbType.[Enum]
                rst = DbType.String
            Case MySqlDbType.[Set]
                rst = DbType.String
            Case MySqlDbType.TinyBlob
                rst = DbType.Object
            Case MySqlDbType.MediumBlob
                rst = DbType.Object
            Case MySqlDbType.LongBlob
                rst = DbType.Object
            Case MySqlDbType.Blob
                rst = DbType.Object
            Case MySqlDbType.VarChar
                rst = DbType.String
            Case MySqlDbType.[String]
                rst = DbType.String
            Case MySqlDbType.Geometry
                Throw New NotImplementedException(myDbType.ToString & " is not currently supported!")
            Case MySqlDbType.UByte
                rst = DbType.UInt16
            Case MySqlDbType.UInt16
                rst = DbType.UInt16
            Case MySqlDbType.UInt32
                rst = DbType.UInt32
            Case MySqlDbType.UInt64
                rst = DbType.UInt64
            Case MySqlDbType.UInt24
                Throw New NotImplementedException(myDbType.ToString & " is not currently supported!")
            Case MySqlDbType.Binary
                rst = DbType.Binary
            Case MySqlDbType.VarBinary
                rst = DbType.Binary
            Case MySqlDbType.TinyText
                rst = DbType.String
            Case MySqlDbType.MediumText
                rst = DbType.Object
            Case MySqlDbType.LongText
                rst = DbType.Object
            Case MySqlDbType.Text
                rst = DbType.String
            Case MySqlDbType.Guid
                rst = DbType.Binary
        End Select
        Return rst
    End Function
End Class
