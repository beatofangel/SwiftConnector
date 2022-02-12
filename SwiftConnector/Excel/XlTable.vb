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
Imports log4net

''' <summary>
''' 表对象(Excel)
''' </summary>
Public Class XlTable
    Inherits XlTableUnit

    Protected Shared logger As ILog = LogManager.GetLogger(Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Property TableHeader As XlTableHeader

    Public Property ColumnHeader As XlColumnHeader

    Public Property RowHeader As XlRowHeader

    Public Property RowRecord As XlRowRecord

    Public Property TableId As String
        Get
            Return If(TableHeader Is Nothing, String.Empty, TableHeader.TableId)
        End Get
        Set(value As String)
            If TableHeader Is Nothing Then
                Throw New Exception("TableHeader has not been initialized.")
            Else
                TableHeader.TableId = value
            End If
        End Set
    End Property

    Public Property TableType As String
        Get
            Return If(TableHeader Is Nothing, String.Empty, TableHeader.TableType)
        End Get
        Set(value As String)
            If TableHeader Is Nothing Then
                Throw New Exception("TableHeader has not been initialized.")
            Else
                TableHeader.TableType = value
            End If
        End Set
    End Property

    Public Property TableName As String
        Get
            Return If(TableHeader Is Nothing, String.Empty, TableHeader.TableName)
        End Get
        Set(value As String)
            If TableHeader Is Nothing Then
                Throw New Exception("TableHeader has not been initialized.")
            Else
                TableHeader.TableName = value
            End If
        End Set
    End Property

    Public ReadOnly Property PrimaryKeys As List(Of String)
        Get
            Return If(ColumnHeader Is Nothing, Nothing, ColumnHeader.PrimaryKeys)
        End Get
    End Property

    Public ReadOnly Property Columns(Optional mode As ExecuteMode = ExecuteMode.Normal) As List(Of XlColumnDefinition)
        Get
            Return If(ColumnHeader Is Nothing, Nothing, ColumnHeader.Columns(mode))
        End Get
    End Property

    Public Sub New()
        RegionType = RegionType.RT_TABLE_GLOBAL
    End Sub

    Public Shared Function Create() As XlTable
        Dim table As New XlTable
        table.TableHeader = New XlTableHeader(table)
        'If String.IsNullOrWhiteSpace(table.TableId) Then
        '    Throw New Exception("Table cannot be found!")
        'End If
        table.ColumnHeader = New XlColumnHeader(table)
        table.RowHeader = New XlRowHeader(table)
        table.RowRecord = New XlRowRecord(table)
        Return table
    End Function

    ''' <summary>
    ''' 保存
    ''' </summary>
    Public Sub Save()
        If RowRecord.Range Is Nothing Then
            MsgBox("No record found.")
            Return
        End If
        Dim mode = Globals.ThisAddIn.MyRibbon.ExecMode
        Dim prefix = DapperService.GetParameterPrefix()
        RowRecord.Range(1, 1).Activate
        For row = 1 To RowRecord.Range.Rows.Count
            Dim colVals As New DynamicParameters
            For col = 1 To RowRecord.Range.Columns.Count
                Dim colVal As String = RowRecord.Range(row, col).Value
                Dim colType As DbType = dbTypeTranslator.Translate(Columns(mode)(col - 1).DataType)
                colVals.Add(Columns(mode)(col - 1).Name, dbTypeTranslator.Format(colType, colVal), colType)
            Next
            Dim rowToBeActivated = row + 1
            Try
                DapperService.Execute(Sub(rst)
                                          RowRecord.Range(rowToBeActivated, 1).Activate
                                      End Sub, sqlCmd.SqlInsertRecord(TableId, Columns(mode).Select(Function(e) e.Name).ToList, prefix), colVals)
            Catch ex As Exception
                ' 发生异常后继续后续处理
                logger.Error(ex)
                If MsgBox(ex.Message, MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                    RowRecord.Range(rowToBeActivated, 1).Activate
                Else
                    Exit For
                End If
            End Try
        Next
    End Sub

    ''' <summary>
    ''' 删除
    ''' </summary>
    ''' <param name="silent">静默模式：默认false</param>
    Public Sub Delete(Optional silent As Boolean = False)
        If Not TableType.ToLower.Contains("table") Then
            MsgBox(String.Format("Non-table target {0}({1}) of type '{2}' cannot be deleted or truncated.", TableId, TableName, TableType))
            Return
        End If

        If Globals.ThisAddIn.MyRibbon.DelMode = DeleteMode.Delete Then
            If PrimaryKeys.Count = 0 Then
                MsgBox("Primary keys were not found, please use 'TRUNCATE' mode instead. ")
                Return
            End If
            If RowRecord.Range Is Nothing Then
                MsgBox("No record found.")
                Return
            End If
            Dim prefix = DapperService.GetParameterPrefix()
            RowRecord.Range(1, 1).Activate
            For row = 1 To RowRecord.Range.Rows.Count
                Dim pkVals As New DynamicParameters
                For col = 1 To RowRecord.Range.Columns.Count
                    If PrimaryKeys.Contains(ColumnHeader.Range(1, col).Value) Then
                        Dim colVal As String = RowRecord.Range(row, col).Value
                        Dim colType As DbType = dbTypeTranslator.Translate(ColumnHeader.Columns()(col - 1).DataType)
                        pkVals.Add(Columns()(col - 1).Name, dbTypeTranslator.Format(colType, colVal), colType)
                        If PrimaryKeys.Count = pkVals.ParameterNames.Count Then
                            Exit For
                        End If
                    End If
                Next
                Dim rowToBeActivated = row + 1
                DapperService.Execute(Sub(rst)
                                          RowRecord.Range(rowToBeActivated, 1).Activate
                                      End Sub, sqlCmd.SqlDeleteTable(TableId, PrimaryKeys, prefix), pkVals)
            Next

        Else
            DapperService.Execute(Sub(rst)
                                      If Not silent Then
                                          MsgBox(String.Format("Table {0}({1}) has been cleared.", TableId, TableName))
                                      End If
                                  End Sub, sqlCmd.SqlTruncateTable(TableId))
        End If
    End Sub

    ''' <summary>
    ''' 读取
    ''' </summary>
    ''' <param name="mode"></param>
    Public Overrides Sub Render(Optional mode As RenderMode = RenderMode.Excel)
        logger.Debug("Render Start")
        TableHeader.Render(mode)
        If Not String.IsNullOrEmpty(TableHeader.TableId) Then
            If Globals.ThisAddIn.MyRibbon.ExecMode = ExecuteMode.Swift And mode = RenderMode.Memory Then
                ColumnHeader.SwiftRender()
            Else
                ColumnHeader.Render(mode)
            End If

            RowHeader.Render(mode)
            If mode = RenderMode.Excel Then
                RowRecord.Render(mode)
                RowHeader.RenderRowIndex(mode)
                If ConfigService.IsAutoFitColumns Then
                    ColumnHeader.Range.EntireColumn.AutoFit()
                End If
            Else
                RowHeader.RenderRowIndex(mode)
                RowRecord.Render(mode)
            End If
        End If
        logger.Debug("Render End")
    End Sub

    Public Overrides Sub Revoke()

        Dim app = Globals.ThisAddIn.Application
        Dim startRow As Integer
        Dim endRow As Integer
        If Not TableHeader Is Nothing Then
            TableHeader.Revoke()
            startRow = TableHeader.Range.Row
        End If

        If Not ColumnHeader Is Nothing Then
            ColumnHeader.Revoke()
        End If

        If Not RowHeader Is Nothing Then
            RowHeader.Revoke()
        End If

        If Not RowRecord Is Nothing Then
            RowRecord.Revoke()

            If Not RowRecord.Range Is Nothing Then
                endRow = RowRecord.Range.Row + RowRecord.Range.Rows.Count - 1
            End If
            If Not RowRecord.NoDataFoundRange Is Nothing Then
                endRow = RowRecord.NoDataFoundRange.Row
            End If
        End If

        If endRow < startRow Then
            Return
        End If
        app.Range(app.Cells(startRow + 1, 1), app.Cells(endRow + 1, 1)).EntireRow.Delete()
    End Sub

End Class
