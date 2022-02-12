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

Imports System.Collections
Imports System.Data

Public Class XlColumnDefinition
    Inherits DataRecordAdapter

    Public Property Name As String

    Public Property Comment As String

    Public Property DataType As String

    Public Property DataLength As Integer

    Public Property DataPrecision As Integer

    Public Property DataScale As Integer

    Public Property Nullable As String

    Public ReadOnly Property [Property] As String
        Get
            Return PropertyColumnFactory.CreateProperty(DataType, DataLength, DataPrecision, DataScale, Nullable)
            'Dim rst As String
            ''Dim curDs As DataSourceType = [Enum].Parse(GetType(DataSourceType), Globals.ThisAddIn.CurDataSource.Type)
            'Select Case Globals.ThisAddIn.CurDataSource.Type
            ''Select Case curDs
            '    Case DataSourceType.Oracle
            '        ' TODO 完善对全部类型的精确显示
            '        Select Case DataType
            '            Case "DATE"
            '                rst = String.Format("{0}[{1}]", DataType, Nullable)
            '            Case "VARCHAR2", "NVARCHAR2"
            '                rst = String.Format("{0}({1})[{2}]", DataType, DataLength, Nullable)
            '            Case "NUMBER"
            '                rst = String.Format("{0}({1},{2})[{3}]", DataType, DataPrecision, DataScale, Nullable)
            '            Case "RAW"
            '                rst = String.Format("{0}({1})[{2}]", DataType, DataLength, Nullable)
            '            Case Else   ' 包含 "TIMESTAMP(X)","TIMESTAMP(X) WITH TIME ZONE","TIMESTAMP(X) WITH LOCAL TIME ZONE"
            '                rst = String.Format("{0}[{1}]", DataType, Nullable)
            '        End Select
            '    Case DataSourceType.MySQL
            '        Select Case DataType.ToUpper
            '            Case "VARCHAR"
            '                rst = String.Format("{0}({1})[{2}]", DataType, DataLength, Nullable)
            '            Case "DECIMAL"
            '                rst = String.Format("{0}({1},{2})[{3}]", DataType, DataPrecision, DataScale, Nullable)
            '            Case "DATETIME", "TIMESTAMP"
            '                If DataLength > 0 Then
            '                    rst = String.Format("{0}({1})[{2}]", DataType, DataLength, Nullable)
            '                Else
            '                    rst = String.Format("{0}[{1}]", DataType, Nullable)
            '                End If
            '            Case Else
            '                rst = String.Format("{0}[{1}]", DataType, Nullable)
            '        End Select
            '    Case DataSourceType.PostgreSQL
            '        Throw New NotImplementedException("PostgreSQL is not currently supported!")
            '    Case DataSourceType.SqlServer
            '        Throw New NotImplementedException("SqlServer is not currently supported!")
            '    Case DataSourceType.SQLite
            '        Throw New NotImplementedException("SQLite is not currently supported!")
            '    Case Else
            '        Throw New NotImplementedException(Globals.ThisAddIn.CurDataSource.Name & " is not currently supported!")
            'End Select

            'Return rst
        End Get
    End Property

    Public Overloads Shared Function Create(record As IDataRecord) As XlColumnDefinition

        Return New XlColumnDefinition() With {
            .Name = record(0),
            .Comment = If(TypeOf record(1) Is DBNull, record(0), record(1)),
            .DataType = record(2),
            .DataLength = If(IsNull(record(3)), 0, record(3)),
            .DataPrecision = If(IsNull(record(4)), 0, record(4)),
            .DataScale = If(IsNull(record(5)), 0, record(5)),
            .Nullable = record(6)
        }
    End Function
End Class
