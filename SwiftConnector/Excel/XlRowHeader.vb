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

Imports log4net

''' <summary>
''' 行头对象(Excel)
''' </summary>
Public Class XlRowHeader
    Inherits XlTableUnit

    Protected Shared logger As ILog = LogManager.GetLogger(Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Property OwnerTable As XlTable
    Public Property Range As Excel.Range
    Public Property RowIndexRange As Excel.Range

    Public Sub New(table As XlTable)
        RegionType = RegionType.RT_ROW_HEADER
        OwnerTable = table
    End Sub

    Public Overrides Sub Render(Optional mode As RenderMode = RenderMode.Excel)
        logger.Debug("Render Start")
        If OwnerTable Is Nothing Then Return
        If OwnerTable.TableHeader Is Nothing Then
            Throw New Exception("Table header has not been initialized.")
        End If

        Dim rowHeaderRowCount As Integer = 2
        Dim IsShowColProps As Boolean = configService.IsShowColProps
        If IsShowColProps Then rowHeaderRowCount += 1

        Dim app = Globals.ThisAddIn.Application
        Range = app.Range(OwnerTable.TableHeader.Range(1, 0), OwnerTable.TableHeader.Range(1 + rowHeaderRowCount, 0))
        If mode = RenderMode.Excel Then
            setFormat(Range)
            Range(1, 1).Value = textService.GetTextByProperty(TextType.TT_ROW_HEADER_TABLE) ' TODO 根据tblId类型显示 Table/View等
            Range(2, 1).Value = textService.GetTextByProperty(TextType.TT_ROW_HEADER_COLUMN)
            Range(3, 1).Value = textService.GetTextByProperty(TextType.TT_ROW_HEADER_COMMENT)
            If IsShowColProps Then Range(4, 1).Value = textService.GetTextByProperty(TextType.TT_ROW_HEADER_PROP)
        End If
        logger.Debug("Render End")
    End Sub

    Public Overrides Sub Revoke()
        If Range IsNot Nothing Then
            Range.Clear()
        End If
        If RowIndexRange IsNot Nothing Then
            RowIndexRange.Clear()
        End If
    End Sub

    Public Sub RenderRowIndex(Optional mode As RenderMode = RenderMode.Excel)
        logger.Debug("RenderRowIndex Start")
        Dim app = Globals.ThisAddIn.Application
        If mode = RenderMode.Memory Then

            Dim row As Integer = 0
            Dim rowHeaderIndexCell As Excel.Range = Range(Range.Rows.Count + 1, 1)
            'While rowHeaderIndexCell.Offset(row, 0).Value <> vbNullString
            While Not String.IsNullOrEmpty(rowHeaderIndexCell.Offset(row, 0).Value)

                ' strict mode
                'If rowHeaderIndexCell.Offset(row, 0).Value = row Then
                '    row += 1
                'End If

                ' loose mode
                row += 1
            End While

            If row > 0 Then
                RowIndexRange = app.Range(rowHeaderIndexCell, rowHeaderIndexCell.Offset(row - 1, 0))
            End If
        Else
            If OwnerTable.RowRecord Is Nothing Then Return

            If OwnerTable.RowRecord.Range IsNot Nothing Then
                RowIndexRange = app.Range(OwnerTable.RowRecord.Range.Item(1, 0), OwnerTable.RowRecord.Range.Item(OwnerTable.RowRecord.Range.Rows.Count, 0))
                setFormat(RowIndexRange)
                For i = 1 To RowIndexRange.Rows.Count
                    RowIndexRange(i, 1).Value = i
                Next
            End If
        End If
        logger.Debug("RenderRowIndex End")
    End Sub

    Private Sub setFormat(rng As Excel.Range)
        With rng
            .Clear()
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
End Class
