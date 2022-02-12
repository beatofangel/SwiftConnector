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
Imports log4net

''' <summary>
''' 表头对象(Excel)
''' </summary>
Public Class XlTableHeader
    Inherits XlTableUnit

    Protected Shared logger As ILog = LogManager.GetLogger(Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private _queryStringCell As Excel.Range

    Public Property OwnerTable As XlTable

    Public Property Range As Excel.Range
    Public Property TableId As String
    Public Property TableName As String
    Public Property TableType As String
    Public ReadOnly Property QueryString As String
        Get
            Return _queryStringCell.Value
        End Get
    End Property

    Public Sub New(table As XlTable)
        RegionType = RegionType.RT_TABLE_HEADER
        OwnerTable = table
        Dim rowOffset As Integer
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        If String.IsNullOrWhiteSpace(app.ActiveCell.Value) Then
            If rowBoundaryCheck() Then
                rowOffset = -1
                If String.IsNullOrWhiteSpace(app.ActiveCell.Offset(rowOffset, 0).Value) Then
                    'MsgBox("Plz input table id or name.")
                    Throw New Exception("Please input table id or name.")
                    'Environment.Exit(0)
                End If
            Else
                Throw New Exception("Please input table id or name.")
            End If
        End If
        Dim colOffset As Integer = calcColumnOffset()
        Range = app.Range(app.ActiveCell.Offset(rowOffset, colOffset), app.ActiveCell.Offset(rowOffset, colOffset + 1))
        _queryStringCell = app.ActiveCell.Offset(rowOffset, 0)

        Dim reg As Regex = New Regex("^\w+[\d|\w]+$", RegexOptions.IgnoreCase)
        If reg.IsMatch(QueryString) Then
            dapperService.ExecuteReader(Sub(reader As IDataReader)
                                            If reader.Read Then
                                                TableId = reader.Item(0)
                                                TableType = reader.Item(1)
                                                TableName = If(TypeOf reader.Item(2) Is DBNull, reader.Item(0), reader.Item(2))
                                            End If
                                        End Sub, sqlCmd.SqlQueryTblDefByTableId(QueryString))
        End If

    End Sub

    Public Overrides Sub Render(Optional mode As RenderMode = RenderMode.Excel)
        logger.Debug("Render Start")
        Dim app = Globals.ThisAddIn.Application
        If mode = RenderMode.Memory Then
            With Range
                .Item(1, 1) = app.ActiveCell.Value
                .Item(1, 2) = app.ActiveCell.Offset(0, 1).Value
            End With
        Else
            If String.IsNullOrEmpty(TableId) Then
                ' 弹出模态框，显示按照表注释进行模糊查询得到的表清单
                'Using frm = New FrmTableList(QueryString)
                '    ' 选择后设定TableId
                '    Select Case frm.ShowDialog()
                '        Case Windows.Forms.DialogResult.OK
                '            TableId = frm.TableInfo.Item(1)
                '            TableType = frm.TableInfo.Item(2)
                '            TableName = frm.TableInfo.Item(3)
                '        Case Windows.Forms.DialogResult.Abort
                '            Throw New TableNotFoundException()
                '    End Select
                'End Using
            End If

            ' 判断是否设定了TableId
            If Not String.IsNullOrEmpty(TableId) Then
                setFormat()
                Range.Item(1, 1) = TableId
                Range.Item(1, 2) = TableName

                ' 插入新行
                Range.Offset(1).EntireRow.Insert(Excel.XlDirection.xlDown)
                Range.Offset(1).EntireRow.ClearFormats()
                Dim targetCell As Excel.Range = Range.Item(1, 1)
                targetCell.Activate()
            End If
        End If
        logger.Debug("Render End")
    End Sub

    Public Overrides Sub Revoke()
        If Range IsNot Nothing Then
            Range.EntireRow.Clear()
            Dim targetCell As Excel.Range = Range.Item(1, 1)
            targetCell.Value = TableId
        End If
    End Sub

    ''' <summary>
    ''' 计算列偏移量
    ''' </summary>
    ''' <returns>Integer</returns>
    Private Function calcColumnOffset() As Integer
        Return If(colBoundaryCheck(), 0, 1)
    End Function

    ''' <summary>
    ''' 列边界校验：当为第1列时，返回False；否则，返回True
    ''' </summary>
    ''' <returns>Boolean</returns>
    Private Function colBoundaryCheck() As Boolean
        Return Globals.ThisAddIn.Application.ActiveCell.Column > 1
    End Function

    ''' <summary>
    ''' 列边界校验：当为第1行时，返回False；否则，返回True
    ''' </summary>
    ''' <returns></returns>
    Private Function rowBoundaryCheck() As Boolean
        Return Globals.ThisAddIn.Application.ActiveCell.Row > 1
    End Function

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
End Class
