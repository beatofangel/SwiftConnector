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
''' 列头对象(Excel)
''' </summary>
Public Class XlColumnHeader
    Inherits XlTableUnit

    Protected Shared logger As ILog = LogManager.GetLogger(Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private columnDefinitions As List(Of XlColumnDefinition)

    Private columnDefinitionsSwift As List(Of XlColumnDefinition)

    Private pkDefinitions As List(Of String)

    Public Property Range As Excel.Range

    Public Property OwnerTable As XlTable

    Public ReadOnly Property PrimaryKeys As List(Of String)
        Get
            Return pkDefinitions
        End Get
    End Property

    Public ReadOnly Property Columns(Optional mode As ExecuteMode = ExecuteMode.Normal) As List(Of XlColumnDefinition)
        Get
            Return If(mode = ExecuteMode.Normal, columnDefinitions, columnDefinitionsSwift)
        End Get
    End Property

#Region "RegionType"
    Private ReadOnly Property RegionInheritanceColName As RegionType() = GetRegionInheritance(RegionType.RT_COLUMN_HEADER_COLNAME)

    Private ReadOnly Property RegionInheritanceComment As RegionType() = GetRegionInheritance(RegionType.RT_COLUMN_HEADER_COMMENT)

    Private ReadOnly Property RegionInheritanceProp As RegionType() = GetRegionInheritance(RegionType.RT_COLUMN_HEADER_PROP)

    Private ReadOnly Property RegionInheritancePK As RegionType() = GetRegionInheritance(RegionType.RT_COLUMN_HEADER_PK)

#End Region

    Public Sub New(table As XlTable)
        RegionType = RegionType.RT_COLUMN_HEADER
        OwnerTable = table
    End Sub

    Public Overrides Sub Render(Optional mode As RenderMode = RenderMode.Excel)
        logger.Debug("Render Start")
        If OwnerTable Is Nothing Then Return

        Dim columnHeaderRowCount As Integer = 2
        Dim IsShowColProps As Boolean = configService.IsShowColProps
        If IsShowColProps Then columnHeaderRowCount += 1
        DapperService.ExecuteReader(Sub(reader)
                                        columnDefinitions = reader.GetData(AddressOf XlColumnDefinition.Create).ToList
                                    End Sub, sqlCmd.SqlQueryColDefByTableId(OwnerTable.TableId))

        DapperService.ExecuteReader(Sub(reader)
                                        ' 需要测试无主键的case
                                        pkDefinitions = reader.GetData(Function(record As IDataRecord) CStr(record(0))).ToList
                                    End Sub, sqlCmd.SqlQueryPkDefByTableId(OwnerTable.TableId))

        Dim app = Globals.ThisAddIn.Application
        If mode = RenderMode.Memory Then
            Range = app.Range(OwnerTable.TableHeader.Range.Item(OwnerTable.TableHeader.Range.Rows.Count + 1, 1), OwnerTable.TableHeader.Range.Item(OwnerTable.TableHeader.Range.Rows.Count + columnHeaderRowCount, columnDefinitions.Count))
        Else
            app.Range(OwnerTable.TableHeader.Range.Item(OwnerTable.TableHeader.Range.Rows.Count + 1, 1), OwnerTable.TableHeader.Range.Item(OwnerTable.TableHeader.Range.Rows.Count + columnHeaderRowCount, columnDefinitions.Count)).EntireRow.Insert()
            Range = app.Range(OwnerTable.TableHeader.Range.Item(OwnerTable.TableHeader.Range.Rows.Count + 1, 1), OwnerTable.TableHeader.Range.Item(OwnerTable.TableHeader.Range.Rows.Count + columnHeaderRowCount, columnDefinitions.Count))

            SetFormat(IsShowColProps)

            SetColDefToRange(columnDefinitions, Range, IsShowColProps)

        End If
        logger.Debug("Render End")
    End Sub

    ''' <summary>
    ''' 迅捷渲染（用于“迅捷”模式insert数据，仅限RenderMode.Memory）
    ''' </summary>
    Public Sub SwiftRender()
        logger.Debug("Swfit Render Start")
        If OwnerTable Is Nothing Then Return
        columnDefinitionsSwift = New List(Of XlColumnDefinition)
        Dim columnHeaderRowCount As Integer = 2
        Dim IsShowColProps As Boolean = ConfigService.IsShowColProps
        If IsShowColProps Then columnHeaderRowCount += 1
        DapperService.ExecuteReader(Sub(reader)
                                        columnDefinitions = reader.GetData(AddressOf XlColumnDefinition.Create).ToList
                                    End Sub, sqlCmd.SqlQueryColDefByTableId(OwnerTable.TableId))

        DapperService.ExecuteReader(Sub(reader)
                                        ' 需要测试无主键的case
                                        pkDefinitions = reader.GetData(Function(record As IDataRecord) CStr(record(0))).ToList
                                    End Sub, sqlCmd.SqlQueryPkDefByTableId(OwnerTable.TableId))

        Dim app = Globals.ThisAddIn.Application
        Dim colCell As Excel.Range = OwnerTable.TableHeader.Range.Item(2, 1)
        While Not String.IsNullOrEmpty(colCell.Value)
            Dim colName = colCell.Value
            Dim colDef = columnDefinitions.Find(Function(cd) cd.Name = colName)
            If colDef Is Nothing Then
                Throw New Exception(String.Format("""{0}"" is not defined in table [{1}]", colName, OwnerTable.TableId))
            End If
            columnDefinitionsSwift.Add(colDef)
            colCell = colCell.Offset(0, 1)
        End While
        Range = app.Range(OwnerTable.TableHeader.Range.Item(OwnerTable.TableHeader.Range.Rows.Count + 1, 1), OwnerTable.TableHeader.Range.Item(OwnerTable.TableHeader.Range.Rows.Count + columnHeaderRowCount, columnDefinitionsSwift.Count))
        logger.Debug("Swfit Render End")
    End Sub

    Public Overrides Sub Revoke()
        If Range IsNot Nothing Then
            Range.Clear()
        End If
    End Sub

    Private Sub SetColDefToRange(colDefs As List(Of XlColumnDefinition), rng As Excel.Range, IsShowColProps As Boolean)
        For col = 1 To rng.Columns.Count
            rng(1, col).Value = colDefs(col - 1).Name
            rng(2, col).Value = colDefs(col - 1).Comment
            If IsShowColProps Then
                rng(3, col).Value = colDefs(col - 1).Property
            End If
        Next
    End Sub

    Private Sub SetFormat(IsShowColProps As Boolean)
        SetGlobalFormat()
        SetColNameFormat()
        SetCommentFormat()
        SetPropFormat(IsShowColProps)
        For colIdx = 0 To columnDefinitions.Count - 1
            For pkIdx = 0 To pkDefinitions.Count - 1
                If pkDefinitions(pkIdx) = columnDefinitions(colIdx).Name Then
                    ' TODO 考虑拆分PK region 分别设置样式
                    SetPkFormat(Range(1, colIdx + 1), RegionInheritancePK)
                    SetPkFormat(Range(2, colIdx + 1), RegionInheritancePK)
                    If IsShowColProps Then SetPkFormat(Range(3, colIdx + 1), RegionInheritancePK)
                End If
            Next
        Next
    End Sub

    Private Sub SetGlobalFormat()
        With Range
            .EntireRow.Clear()
            .NumberFormatLocal = "@"
        End With
    End Sub

    Private Sub SetColNameFormat()
        With Range.Rows(1)
            .NumberFormatLocal = "@"
            styleService.XlSetCellBorderDiagonalUp(RegionInheritanceColName, .Borders(Excel.XlBordersIndex.xlDiagonalUp))
            styleService.XlSetCellBorderDiagonalDown(RegionInheritanceColName, .Borders(Excel.XlBordersIndex.xlDiagonalDown))
            styleService.XlSetCellBorderEdgeTop(RegionInheritanceColName, .Borders(Excel.XlBordersIndex.xlEdgeTop))
            styleService.XlSetCellBorderEdgeBottom(RegionInheritanceColName, .Borders(Excel.XlBordersIndex.xlEdgeBottom))
            styleService.XlSetCellBorderEdgeLeft(RegionInheritanceColName, .Borders(Excel.XlBordersIndex.xlEdgeLeft))
            styleService.XlSetCellBorderEdgeRight(RegionInheritanceColName, .Borders(Excel.XlBordersIndex.xlEdgeRight))
            styleService.XlSetCellBorderInsideHorizontal(RegionInheritanceColName, .Borders(Excel.XlBordersIndex.xlInsideHorizontal))
            styleService.XlSetCellBorderInsideVertical(RegionInheritanceColName, .Borders(Excel.XlBordersIndex.xlInsideVertical))
            styleService.XlSetCellFont(RegionInheritanceColName, .Font)
            styleService.XlSetCellInterior(RegionInheritanceColName, .Interior)
        End With
    End Sub

    Private Sub SetCommentFormat()
        With Range.Rows(2)
            .NumberFormatLocal = "@"
            styleService.XlSetCellBorderDiagonalUp(RegionInheritanceComment, .Borders(Excel.XlBordersIndex.xlDiagonalUp))
            styleService.XlSetCellBorderDiagonalDown(RegionInheritanceComment, .Borders(Excel.XlBordersIndex.xlDiagonalDown))
            styleService.XlSetCellBorderEdgeTop(RegionInheritanceComment, .Borders(Excel.XlBordersIndex.xlEdgeTop))
            styleService.XlSetCellBorderEdgeBottom(RegionInheritanceComment, .Borders(Excel.XlBordersIndex.xlEdgeBottom))
            styleService.XlSetCellBorderEdgeLeft(RegionInheritanceComment, .Borders(Excel.XlBordersIndex.xlEdgeLeft))
            styleService.XlSetCellBorderEdgeRight(RegionInheritanceComment, .Borders(Excel.XlBordersIndex.xlEdgeRight))
            styleService.XlSetCellBorderInsideHorizontal(RegionInheritanceComment, .Borders(Excel.XlBordersIndex.xlInsideHorizontal))
            styleService.XlSetCellBorderInsideVertical(RegionInheritanceComment, .Borders(Excel.XlBordersIndex.xlInsideVertical))
            styleService.XlSetCellFont(RegionInheritanceComment, .Font)
            styleService.XlSetCellInterior(RegionInheritanceComment, .Interior)
        End With
    End Sub

    Private Sub SetPropFormat(IsShowColProps As Boolean)
        If Not IsShowColProps Then Return
        With Range.Rows(3)
            .NumberFormatLocal = "@"
            styleService.XlSetCellBorderDiagonalUp(RegionInheritanceProp, .Borders(Excel.XlBordersIndex.xlDiagonalUp))
            styleService.XlSetCellBorderDiagonalDown(RegionInheritanceProp, .Borders(Excel.XlBordersIndex.xlDiagonalDown))
            styleService.XlSetCellBorderEdgeTop(RegionInheritanceProp, .Borders(Excel.XlBordersIndex.xlEdgeTop))
            styleService.XlSetCellBorderEdgeBottom(RegionInheritanceProp, .Borders(Excel.XlBordersIndex.xlEdgeBottom))
            styleService.XlSetCellBorderEdgeLeft(RegionInheritanceProp, .Borders(Excel.XlBordersIndex.xlEdgeLeft))
            styleService.XlSetCellBorderEdgeRight(RegionInheritanceProp, .Borders(Excel.XlBordersIndex.xlEdgeRight))
            styleService.XlSetCellBorderInsideHorizontal(RegionInheritanceProp, .Borders(Excel.XlBordersIndex.xlInsideHorizontal))
            styleService.XlSetCellBorderInsideVertical(RegionInheritanceProp, .Borders(Excel.XlBordersIndex.xlInsideVertical))
            styleService.XlSetCellFont(RegionInheritanceProp, .Font)
            styleService.XlSetCellInterior(RegionInheritanceProp, .Interior)
        End With
    End Sub

    Private Sub SetPkFormat(rng As Excel.Range, regions As RegionType())
        With rng
            .NumberFormatLocal = "@"
            styleService.XlSetCellBorderDiagonalUp(regions, .Borders(Excel.XlBordersIndex.xlDiagonalUp))
            styleService.XlSetCellBorderDiagonalDown(regions, .Borders(Excel.XlBordersIndex.xlDiagonalDown))
            styleService.XlSetCellBorderEdgeTop(regions, .Borders(Excel.XlBordersIndex.xlEdgeTop))
            styleService.XlSetCellBorderEdgeBottom(regions, .Borders(Excel.XlBordersIndex.xlEdgeBottom))
            styleService.XlSetCellBorderEdgeLeft(regions, .Borders(Excel.XlBordersIndex.xlEdgeLeft))
            styleService.XlSetCellBorderEdgeRight(regions, .Borders(Excel.XlBordersIndex.xlEdgeRight))
            styleService.XlSetCellBorderInsideHorizontal(regions, .Borders(Excel.XlBordersIndex.xlInsideHorizontal))
            styleService.XlSetCellBorderInsideVertical(regions, .Borders(Excel.XlBordersIndex.xlInsideVertical))
            styleService.XlSetCellFont(regions, .Font)
            styleService.XlSetCellInterior(regions, .Interior)
        End With
    End Sub

End Class
