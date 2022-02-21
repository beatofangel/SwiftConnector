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
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Reflection
Imports System.Runtime.CompilerServices

Module Extensions
    <Extension()>
    Function AddArrayParameters(Of S As DbParameter)(ByVal cmd As DbCommand, ByVal paramNameRoot As String, ByVal values As IEnumerable(Of Object), ByVal Optional dbType As DbType? = Nothing, ByVal Optional size As Integer? = Nothing) As S()
        Dim parameters = New List(Of DbParameter)()
        Dim parameterNames = New List(Of String)()
        Dim paramNbr = 1

        Dim d = GetType(S)

        For Each value In values
            Dim paramName = String.Format("@{0}{1}", paramNameRoot, Math.Min(System.Threading.Interlocked.Increment(paramNbr), paramNbr - 1))
            parameterNames.Add(paramName)
            Dim p As DbParameter = Activator.CreateInstance(d, paramName, value)
            If dbType.HasValue Then p.DbType = dbType.Value
            If size.HasValue Then p.Size = size.Value
            cmd.Parameters.Add(p)
            parameters.Add(p)
        Next

        cmd.CommandText = cmd.CommandText.Replace("{" & paramNameRoot & "}", String.Join(",", parameterNames))
        Return parameters.ToArray()
    End Function

    <Extension()>
    Public Iterator Function GetData(Of T)(reader As IDataReader, BuildObject As Func(Of IDataRecord, T)) As IEnumerable(Of T)
        Try
            While reader.Read
                Yield BuildObject(reader)
            End While
        Finally
            reader.Dispose()
        End Try
    End Function

    <Extension()>
    Public Function ToArray(data As DataTable) As Object(,)
        Dim ret = TryCast(Array.CreateInstance(GetType(Object), data.Rows.Count, data.Columns.Count), Object(,))

        For i = 0 To data.Rows.Count - 1
            For j = 0 To data.Columns.Count - 1
                ret(i, j) = data.Rows(i)(j)
            Next
        Next

        Return ret
    End Function

    <Extension()>
    Public Function GetParameterPrefix(service As IDapperService) As String
        Using db As DbConnection = GetCurrentConnection()
            db.Open()
            Dim tbl = db.GetSchema(DbMetaDataCollectionNames.DataSourceInformation)
            Dim markerFormat As String = tbl.Rows.Item(0).Item("ParameterMarkerFormat")
            Dim temp = markerFormat.Split({"{0}"}, StringSplitOptions.RemoveEmptyEntries)
            If temp IsNot Nothing AndAlso temp.Length > 0 Then
                Return temp(0)
            End If
        End Using
        Return String.Empty
    End Function

    ''' <summary>
    ''' 实现.NET Guid类型与Oracle Guid类型的互相转换
    ''' </summary>
    ''' <param name="guid"></param>
    ''' <returns></returns>
    <Extension()>
    Public Function FlipEndian(guid As Guid) As Guid
        Dim newBytes = New Byte(15) {}
        Dim oldBytes = guid.ToByteArray()

        For i As Integer = 8 To 15
            newBytes(i) = oldBytes(i)
        Next

        newBytes(3) = oldBytes(0)
        newBytes(2) = oldBytes(1)
        newBytes(1) = oldBytes(2)
        newBytes(0) = oldBytes(3)
        newBytes(5) = oldBytes(4)
        newBytes(4) = oldBytes(5)
        newBytes(6) = oldBytes(7)
        newBytes(7) = oldBytes(6)

        Return New Guid(newBytes)
    End Function

    <Extension()>
    Function ToByteArray(image As Image, format As ImageFormat) As Byte()
        Using ms As MemoryStream = New MemoryStream()
            image.Save(ms, format)
            Return ms.ToArray()
        End Using
    End Function
End Module
