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
Imports log4net

''' <summary>
''' 行记录对象(Excel)
''' </summary>
Public Class XlRowRecord
    Inherits XlTableUnit

    Protected Shared logger As ILog = LogManager.GetLogger(Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Property OwnerTable As XlTable

    Public Property Range As Excel.Range

    Public Property NoDataFoundRange As Excel.Range

    Private ReadOnly Property RegionInheritanceNoData As RegionType() = GetRegionInheritance(RegionType.RT_ROW_DATA_NOT_FOUND)

    Public Sub New(table As XlTable)
        RegionType = RegionType.RT_ROW_DATA
        OwnerTable = table
    End Sub

    Public Overrides Sub Render(Optional mode As RenderMode = RenderMode.Excel)
        logger.Debug("Render Start")
        If OwnerTable Is Nothing Then Return
        If OwnerTable.ColumnHeader Is Nothing Then
            Throw New Exception("Column header has not been initialized.")
        End If
        If OwnerTable.RowHeader Is Nothing Then
            Throw New Exception("Row header has not been initialized.")
        End If

        Dim columnHeaderRowCount As Integer = 2
        Dim IsShowColProps As Boolean = styleService.IsShowColProps
        If IsShowColProps Then columnHeaderRowCount += 1

        Dim app = Globals.ThisAddIn.Application
        Dim startRow = OwnerTable.RowHeader.Range.Row + OwnerTable.RowHeader.Range.Rows.Count
        Dim startCol = OwnerTable.ColumnHeader.Range.Column
        Dim endRow = startRow
        Dim endCol = OwnerTable.ColumnHeader.Range.Column + OwnerTable.ColumnHeader.Range.Columns.Count - 1
        If mode = RenderMode.Memory Then
            If OwnerTable.RowHeader.RowIndexRange IsNot Nothing Then
                endRow = startRow + OwnerTable.RowHeader.RowIndexRange.Rows.Count - 1
                Range = app.Range(app.Cells(startRow, startCol), app.Cells(endRow, endCol))
            End If
        Else
            Dim data As Object(,) = Nothing
            Dim sqlQueryRecordByTableId As String = String.Empty
            dapperService.ExecuteReader(Sub(reader)
                                            If reader.Read Then
                                                sqlQueryRecordByTableId = CStr(reader.Item(0))
                                            End If
                                        End Sub, sqlCmd.SqlQueryRecordSqlByTableId(OwnerTable.TableId))
            dapperService.ExecuteReader(Sub(reader)
                                            Dim dt As New DataTable
                                            dt.Load(reader)
                                            If dt.Rows.Count > 0 Then
                                                data = dt.ToArray
                                            End If
                                        End Sub, sqlQueryRecordByTableId)
            If data Is Nothing Then
                app.Range(app.Cells(startRow, startCol), app.Cells(startRow, startCol)).EntireRow.Insert()
                NoDataFoundRange = app.Range(app.Cells(startRow, startCol), app.Cells(startRow, startCol))

                setFormatDataNotFound()
                NoDataFoundRange.Value = textService.GetTextByProperty(TextType.TT_NO_DATA_FOUND)
            Else
                endRow = endRow + UBound(data, 1)
                app.Range(app.Cells(startRow, startCol), app.Cells(endRow, endCol)).EntireRow.Insert()
                Range = app.Range(app.Cells(startRow, startCol), app.Cells(endRow, endCol))

                setFormat()
                Range.Value = data
            End If
        End If
        logger.Debug("Render End")
    End Sub

    Public Overrides Sub Revoke()
        If Range IsNot Nothing Then
            Range.Clear()
        End If
        If NoDataFoundRange IsNot Nothing Then
            NoDataFoundRange.Clear()
        End If
    End Sub

    Private Sub setFormat()
        With Range
            .EntireRow.Clear()
            .NumberFormatLocal = "@"
            styleService.XlSetCellBorderDiagonalUp(RegionInheritance, .Borders(Excel.XlBordersIndex.xlDiagonalUp))
            styleService.XlSetCellBorderDiagonalDown(RegionInheritance, .Borders(Excel.XlBordersIndex.xlDiagonalDown))
            styleService.XlSetCellBorderEdgeTop(RegionInheritance, .Borders(Excel.XlBordersIndex.xlEdgeTop))
            styleService.XlSetCellBorderEdgeBottom(RegionInheritance, .Borders(Excel.XlBordersIndex.xlEdgeBottom))
            styleService.XlSetCellBorderEdgeLeft(RegionInheritance, .Borders(Excel.XlBordersIndex.xlEdgeLeft))
            styleService.XlSetCellBorderEdgeRight(RegionInheritance, .Borders(Excel.XlBordersIndex.xlEdgeRight))
            styleService.XlSetCellBorderInsideHorizontal(RegionInheritance, .Borders(Excel.XlBordersIndex.xlInsideHorizontal))
            styleService.XlSetCellBorderInsideVertical(RegionInheritance, .Borders(Excel.XlBordersIndex.xlInsideVertical))
            styleService.XlSetCellFont(RegionInheritance, .Font)
            styleService.XlSetCellInterior(RegionInheritance, .Interior)
        End With
    End Sub

    Private Sub setFormatDataNotFound()

        With NoDataFoundRange
            .EntireRow.Clear()
            .NumberFormatLocal = "@"
            styleService.XlSetCellBorderDiagonalUp(RegionInheritanceNoData, .Borders(Excel.XlBordersIndex.xlDiagonalUp))
            styleService.XlSetCellBorderDiagonalDown(RegionInheritanceNoData, .Borders(Excel.XlBordersIndex.xlDiagonalDown))
            styleService.XlSetCellBorderEdgeTop(RegionInheritanceNoData, .Borders(Excel.XlBordersIndex.xlEdgeTop))
            styleService.XlSetCellBorderEdgeBottom(RegionInheritanceNoData, .Borders(Excel.XlBordersIndex.xlEdgeBottom))
            styleService.XlSetCellBorderEdgeLeft(RegionInheritanceNoData, .Borders(Excel.XlBordersIndex.xlEdgeLeft))
            styleService.XlSetCellBorderEdgeRight(RegionInheritanceNoData, .Borders(Excel.XlBordersIndex.xlEdgeRight))
            styleService.XlSetCellBorderInsideHorizontal(RegionInheritanceNoData, .Borders(Excel.XlBordersIndex.xlInsideHorizontal))
            styleService.XlSetCellBorderInsideVertical(RegionInheritanceNoData, .Borders(Excel.XlBordersIndex.xlInsideVertical))
            styleService.XlSetCellFont(RegionInheritanceNoData, .Font)
            styleService.XlSetCellInterior(RegionInheritanceNoData, .Interior)
        End With
    End Sub
End Class
