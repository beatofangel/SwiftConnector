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

Public Interface IStyleService

    Function GetStyleByRegionAndType(regionsSeniority As RegionType(), style As StyleType, prop As String, Optional localeIndependent As Boolean = False) As Object
    Function GetPresetStyleByRegionAndType(region As RegionType, style As StyleType, prop As String, Optional localeIndependent As Boolean = False) As Object
    Function GetCustomStyleByRegionAndType(region As RegionType, style As StyleType, prop As String, Optional localeIndependent As Boolean = False) As Object
    Function IsShowColProps() As Boolean
    Function GetRecordLimit() As Integer
    Function GetDelMode() As DeleteMode
    Sub XlSetCellFont(regions As RegionType(), target As Excel.Font)
    Sub XlSetCellInterior(regions As RegionType(), target As Excel.Interior)
    Sub XlSetCellBorderDiagonalUp(regions As RegionType(), target As Excel.Border)
    Sub XlSetCellBorderDiagonalDown(regions As RegionType(), target As Excel.Border)
    Sub XlSetCellBorderEdgeTop(regions As RegionType(), target As Excel.Border)
    Sub XlSetCellBorderEdgeBottom(regions As RegionType(), target As Excel.Border)
    Sub XlSetCellBorderEdgeLeft(regions As RegionType(), target As Excel.Border)
    Sub XlSetCellBorderEdgeRight(regions As RegionType(), target As Excel.Border)
    Sub XlSetCellBorderInsideHorizontal(regions As RegionType(), target As Excel.Border)
    Sub XlSetCellBorderInsideVertical(regions As RegionType(), target As Excel.Border)

    Sub Change(style As Style)

    Sub BatchChange(styles As Style())

    Function Read(style As Style) As Style

    Function ReadPredefinedBorderStyle(region As RegionType) As PredefinedBorder

    Sub Reset(Optional region As RegionType? = Nothing, Optional locale As String = Nothing)
End Interface
