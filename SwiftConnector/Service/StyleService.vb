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
Imports System.Data.SQLite
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel

Public Class StyleService
    Inherits SQLiteBaseService
    Implements IStyleService

    Enum StyleCategory
        PRESET = 0
        CUSTOM = 1
    End Enum

    Private Const NO_LOCALE = "-"

    Private Function GetStyle(cat As StyleCategory, region As RegionType, style As StyleType, prop As String, Optional localeIndependent As Boolean = False)
        Dim rst As Object = Nothing
        Dim sql As String = "SELECT VAL FROM STYLE_" & [Enum].GetName(cat.GetType, cat) & " WHERE REGION=@REGION AND TYPE=@TYPE AND LOCALE=@LOCALE AND PROP=@PROP"
        Dim parameters As New List(Of DbParameter) From {
            New SQLiteParameter("@REGION", Data.DbType.Int32) With {.Value = region},
            New SQLiteParameter("@TYPE", Data.DbType.Int32) With {.Value = style},
            New SQLiteParameter("@LOCALE", Data.DbType.String) With {.Value = If(localeIndependent, NO_LOCALE, Globals.ThisAddIn.StrLangCode)},
            New SQLiteParameter("@PROP", Data.DbType.String) With {.Value = prop}
        }
        ExecuteReader(Sub(reader As DbDataReader)
                          If reader.Read Then
                              rst = reader.Item(0)
                          End If
                      End Sub, sql, parameters.ToArray)
        Return rst
    End Function

    ''' <summary>
    ''' 获取STYLE自定义设定
    ''' </summary>
    ''' <param name="region">区域类型</param>
    ''' <param name="style">风格类型</param>
    ''' <param name="prop">属性名</param>
    ''' <param name="localeIndependent">语言环境无关FLAG（True: 无关，False：有关）</param>
    ''' <returns></returns>
    Public Function GetCustomStyleByRegionAndType(region As RegionType, style As StyleType, prop As String, Optional localeIndependent As Boolean = False) As Object Implements IStyleService.GetCustomStyleByRegionAndType
        Return GetStyle(StyleCategory.CUSTOM, region, style, prop, localeIndependent)
    End Function

    ''' <summary>
    ''' 获取STYLE默认设定
    ''' </summary>
    ''' <param name="region">区域类型</param>
    ''' <param name="style">风格类型</param>
    ''' <param name="prop">属性名</param>
    ''' <param name="localeIndependent">语言环境无关FLAG（True: 无关，False：有关）</param>
    ''' <returns></returns>
    Public Function GetPresetStyleByRegionAndType(region As RegionType, style As StyleType, prop As String, Optional localeIndependent As Boolean = False) As Object Implements IStyleService.GetPresetStyleByRegionAndType
        Return GetStyle(StyleCategory.PRESET, region, style, prop, localeIndependent)
    End Function

    ''' <summary>
    ''' 获取STYLE设定
    ''' </summary>
    ''' <param name="regionsSeniority">区域类型(顺序：子区域->父区域)</param>
    ''' <param name="style">风格类型</param>
    ''' <param name="prop">属性名</param>
    ''' <param name="localeIndependent">语言环境无关FLAG（True: 无关，False：有关）</param>
    ''' <returns></returns>
    Public Function GetStyleByRegionAndType(regionsSeniority As RegionType(), style As StyleType, prop As String, Optional localeIndependent As Boolean = False) As Object Implements IStyleService.GetStyleByRegionAndType
        Dim rst As Object = Nothing
        For Each region In regionsSeniority
            '查询自定义样式
            rst = GetCustomStyleByRegionAndType(region, style, prop, localeIndependent)
            If rst IsNot Nothing Then
                Exit For
            End If
        Next
        If rst Is Nothing Then
            For Each region In regionsSeniority
                '查询预设样式
                rst = GetPresetStyleByRegionAndType(region, style, prop, localeIndependent)
                If rst IsNot Nothing Then
                    Exit For
                End If
            Next
        End If
        Return rst
    End Function

    Public Function IsShowColProps() As Boolean Implements IStyleService.IsShowColProps
        Dim rst As Boolean = False
        Dim showColProps = GetStyleByRegionAndType({RegionType.RT_TABLE_GLOBAL}, StyleType.ST_COL_PROPS_DISPLAY, "ShowColumnProperties", True)
        If showColProps IsNot Nothing Then
            rst = CBool(showColProps)
        End If

        Return rst
    End Function

    Public Function GetRecordLimit() As Integer Implements IStyleService.GetRecordLimit
        Dim rst As Integer = 0 ' no limit
        Dim recordLimit = GetStyleByRegionAndType({RegionType.RT_TABLE_GLOBAL}, StyleType.ST_RECORD_LIMIT, "RecordLimit", True)
        If recordLimit IsNot Nothing Then
            rst = CInt(recordLimit)
        End If

        Return rst
    End Function

    Public Function GetDelMode() As DeleteMode Implements IStyleService.GetDelMode
        Dim rst As DeleteMode = DeleteMode.Truncate
        Dim delMode = GetStyleByRegionAndType({RegionType.RT_TABLE_GLOBAL}, StyleType.ST_DELETE_MODE, "DeleteMode", True)
        If delMode IsNot Nothing Then
            rst = [Enum].Parse(GetType(DeleteMode), delMode)
        End If

        Return rst
    End Function
    ''' <summary>
    ''' 设置字体
    ''' </summary>
    ''' <param name="regions"></param>
    ''' <param name="target"></param>
    Public Sub XlSetCellFont(regions As RegionType(), target As Excel.Font) Implements IStyleService.XlSetCellFont
        Dim rst As Object
        rst = GetStyleByRegionAndType(regions, StyleType.ST_FONT, "Name")
        If rst IsNot Nothing Then
            target.Name = rst
        End If
        rst = GetStyleByRegionAndType(regions, StyleType.ST_FONT, "Size")
        If rst IsNot Nothing Then
            target.Size = rst
        End If
        rst = GetStyleByRegionAndType(regions, StyleType.ST_FONT, "Color")
        If rst IsNot Nothing Then
            'target.Color = StrDecimalToRGB(CStr(rst))
            target.Color = GetSystemColor(rst)
        End If
        rst = GetStyleByRegionAndType(regions, StyleType.ST_FONT, "Bold")
        If rst IsNot Nothing Then
            target.Bold = rst
        End If
        rst = GetStyleByRegionAndType(regions, StyleType.ST_FONT, "Italic")
        If rst IsNot Nothing Then
            target.Italic = rst
        End If
    End Sub

    ''' <summary>
    ''' 设置背景色
    ''' </summary>
    ''' <param name="regions"></param>
    ''' <param name="target"></param>
    Public Sub XlSetCellInterior(regions As RegionType(), target As Excel.Interior) Implements IStyleService.XlSetCellInterior
        Dim rst As Object
        rst = GetStyleByRegionAndType(regions, StyleType.ST_INTERIOR, "Color")
        If rst IsNot Nothing Then
            'target.Color = StrDecimalToRGB(CStr(rst))
            If rst = "#FFFFFF" Then ' 底色为白色时,跳过涂色
                'target.ColorIndex = XlColorIndex.xlColorIndexNone
            Else
                target.Color = GetSystemColor(rst)
            End If
        End If
    End Sub

    ''' <summary>
    ''' 设置左下右上边框
    ''' </summary>
    ''' <param name="regions">区域</param>
    ''' <param name="target">目标属性</param>
    Public Sub XlSetCellBorderDiagonalUp(regions As RegionType(), target As Excel.Border) Implements IStyleService.XlSetCellBorderDiagonalUp
        Dim rst As Object
        rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_DIAGONAL_UP, "LineStyle")
        If rst IsNot Nothing Then
            Dim lineStyle As XlLineStyle = [Enum].Parse(GetType(XlLineStyle), rst)
            If lineStyle <> XlLineStyle.xlLineStyleNone Then
                target.LineStyle = lineStyle
                rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_DIAGONAL_UP, "Weight")
                If rst IsNot Nothing Then
                    target.Weight = [Enum].Parse(GetType(XlBorderWeight), rst)
                End If
                rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_DIAGONAL_UP, "Color")
                If rst IsNot Nothing Then
                    'target.Color = StrDecimalToRGB(CStr(rst))
                    target.Color = GetSystemColor(rst)
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' 设置左上右下边框
    ''' </summary>
    ''' <param name="regions">区域</param>
    ''' <param name="target">目标属性</param>
    Public Sub XlSetCellBorderDiagonalDown(regions As RegionType(), target As Excel.Border) Implements IStyleService.XlSetCellBorderDiagonalDown
        Dim rst As Object
        rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_DIAGONAL_DOWN, "LineStyle")
        If rst IsNot Nothing Then
            Dim lineStyle As XlLineStyle = [Enum].Parse(GetType(XlLineStyle), rst)
            If lineStyle <> XlLineStyle.xlLineStyleNone Then
                target.LineStyle = lineStyle
                rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_DIAGONAL_DOWN, "Weight")
                If rst IsNot Nothing Then
                    target.Weight = [Enum].Parse(GetType(XlBorderWeight), rst)
                End If
                rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_DIAGONAL_DOWN, "Color")
                If rst IsNot Nothing Then
                    'target.Color = StrDecimalToRGB(CStr(rst))
                    target.Color = GetSystemColor(rst)
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' 设置上边框
    ''' </summary>
    ''' <param name="regions">区域</param>
    ''' <param name="target">目标属性</param>
    Public Sub XlSetCellBorderEdgeTop(regions As RegionType(), target As Excel.Border) Implements IStyleService.XlSetCellBorderEdgeTop
        Dim rst As Object
        rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_EDGE_TOP, "LineStyle")
        If rst IsNot Nothing Then
            Dim lineStyle As XlLineStyle = [Enum].Parse(GetType(XlLineStyle), rst)
            If lineStyle <> XlLineStyle.xlLineStyleNone Then
                target.LineStyle = lineStyle
                rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_EDGE_TOP, "Weight")
                If rst IsNot Nothing Then
                    target.Weight = [Enum].Parse(GetType(XlBorderWeight), rst)
                End If
                rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_EDGE_TOP, "Color")
                If rst IsNot Nothing Then
                    'target.Color = StrDecimalToRGB(CStr(rst))
                    target.Color = GetSystemColor(rst)
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' 设置下边框
    ''' </summary>
    ''' <param name="regions">区域</param>
    ''' <param name="target">目标属性</param>
    Public Sub XlSetCellBorderEdgeBottom(regions As RegionType(), target As Excel.Border) Implements IStyleService.XlSetCellBorderEdgeBottom
        Dim rst As Object
        rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_EDGE_BOTTOM, "LineStyle")
        If rst IsNot Nothing Then
            Dim lineStyle As XlLineStyle = [Enum].Parse(GetType(XlLineStyle), rst)
            If lineStyle <> XlLineStyle.xlLineStyleNone Then
                target.LineStyle = lineStyle
                rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_EDGE_BOTTOM, "Weight")
                If rst IsNot Nothing Then
                    target.Weight = [Enum].Parse(GetType(XlBorderWeight), rst)
                End If
                rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_EDGE_BOTTOM, "Color")
                If rst IsNot Nothing Then
                    'target.Color = StrDecimalToRGB(CStr(rst))
                    target.Color = GetSystemColor(rst)
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' 设置左边框
    ''' </summary>
    ''' <param name="regions">区域</param>
    ''' <param name="target">目标属性</param>
    Public Sub XlSetCellBorderEdgeLeft(regions As RegionType(), target As Excel.Border) Implements IStyleService.XlSetCellBorderEdgeLeft
        Dim rst As Object
        rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_EDGE_LEFT, "LineStyle")
        If rst IsNot Nothing Then
            Dim lineStyle As XlLineStyle = [Enum].Parse(GetType(XlLineStyle), rst)
            If lineStyle <> XlLineStyle.xlLineStyleNone Then
                target.LineStyle = lineStyle
                rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_EDGE_LEFT, "Weight")
                If rst IsNot Nothing Then
                    target.Weight = [Enum].Parse(GetType(XlBorderWeight), rst)
                End If
                rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_EDGE_LEFT, "Color")
                If rst IsNot Nothing Then
                    'target.Color = StrDecimalToRGB(CStr(rst))
                    target.Color = GetSystemColor(rst)
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' 设置右边框
    ''' </summary>
    ''' <param name="regions">区域</param>
    ''' <param name="target">目标属性</param>
    Public Sub XlSetCellBorderEdgeRight(regions As RegionType(), target As Excel.Border) Implements IStyleService.XlSetCellBorderEdgeRight
        Dim rst As Object
        rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_EDGE_RIGHT, "LineStyle")
        If rst IsNot Nothing Then
            Dim lineStyle As XlLineStyle = [Enum].Parse(GetType(XlLineStyle), rst)
            If lineStyle <> XlLineStyle.xlLineStyleNone Then
                target.LineStyle = lineStyle
                rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_EDGE_RIGHT, "Weight")
                If rst IsNot Nothing Then
                    target.Weight = [Enum].Parse(GetType(XlBorderWeight), rst)
                End If
                rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_EDGE_RIGHT, "Color")
                If rst IsNot Nothing Then
                    'target.Color = StrDecimalToRGB(CStr(rst))
                    target.Color = GetSystemColor(rst)
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' 设置内边框（水平）
    ''' </summary>
    ''' <param name="regions">区域</param>
    ''' <param name="target">目标属性</param>
    Public Sub XlSetCellBorderInsideHorizontal(regions As RegionType(), target As Excel.Border) Implements IStyleService.XlSetCellBorderInsideHorizontal
        Dim rst As Object
        rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_INSIDE_HORIZONTAL, "LineStyle")
        If rst IsNot Nothing Then
            Dim lineStyle As XlLineStyle = [Enum].Parse(GetType(XlLineStyle), rst)
            If lineStyle <> XlLineStyle.xlLineStyleNone Then
                target.LineStyle = lineStyle
                rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_INSIDE_HORIZONTAL, "Weight")
                If rst IsNot Nothing Then
                    target.Weight = [Enum].Parse(GetType(XlBorderWeight), rst)
                End If
                rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_INSIDE_HORIZONTAL, "Color")
                If rst IsNot Nothing Then
                    'target.Color = StrDecimalToRGB(CStr(rst))
                    target.Color = GetSystemColor(rst)
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' 设置内边框（垂直）
    ''' </summary>
    ''' <param name="regions">区域</param>
    ''' <param name="target">目标属性</param>
    Public Sub XlSetCellBorderInsideVertical(regions As RegionType(), target As Excel.Border) Implements IStyleService.XlSetCellBorderInsideVertical
        Dim rst As Object
        rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_INSIDE_VERTICAL, "LineStyle")
        If rst IsNot Nothing Then
            Dim lineStyle As XlLineStyle = [Enum].Parse(GetType(XlLineStyle), rst)
            If lineStyle <> XlLineStyle.xlLineStyleNone Then
                target.LineStyle = lineStyle
                rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_INSIDE_VERTICAL, "Weight")
                If rst IsNot Nothing Then
                    target.Weight = [Enum].Parse(GetType(XlBorderWeight), rst)
                End If
                rst = GetStyleByRegionAndType(regions, StyleType.ST_BORDER_INSIDE_VERTICAL, "Color")
                If rst IsNot Nothing Then
                    'target.Color = StrDecimalToRGB(CStr(rst))
                    target.Color = GetSystemColor(rst)
                End If
            End If
        End If
    End Sub

    Public Sub Change(style As Style) Implements IStyleService.Change
        Dim sql As String = "INSERT OR REPLACE INTO STYLE_CUSTOM (REGION,TYPE,LOCALE,PROP,VAL,LASTCHANGE) VALUES (@REGION,@TYPE,@LOCALE,@PROP,@VAL,@LASTCHANGE)"
        Dim parameters As DbParameter() = {
            New SQLiteParameter("@REGION", DbType.Int32) With {.Value = style.Region},
            New SQLiteParameter("@TYPE", DbType.Int32) With {.Value = style.Type},
            New SQLiteParameter("@LOCALE", DbType.String) With {.Value = style.Locale},
            New SQLiteParameter("@PROP", DbType.String) With {.Value = style.Prop},
            New SQLiteParameter("@VAL", DbType.String) With {.Value = style.Val},
            New SQLiteParameter("@LASTCHANGE", DbType.DateTime) With {.Value = style.Lastchange}
                                   }
        ExecuteNonQuery(Sub(rst)
                            logger.Debug(String.Format("{0}-{1}-{2}-{3}=>{4}@{5}", style.Region, style.Type, style.Locale, style.Prop, style.Val, style.Lastchange))
                        End Sub, sql, parameters)
    End Sub

    Public Sub BatchChange(styles As Style()) Implements IStyleService.BatchChange
        Dim sql As String = "INSERT OR REPLACE INTO STYLE_CUSTOM (REGION,TYPE,LOCALE,PROP,VAL,LASTCHANGE) VALUES (@REGION,@TYPE,@LOCALE,@PROP,@VAL,@LASTCHANGE)"
        Dim sqls As New List(Of String)
        Dim parametersList As New List(Of DbParameter())
        For Each style In styles
            sqls.Add(sql)
            Dim parameters As DbParameter() = {
                New SQLiteParameter("@REGION", DbType.Int32) With {.Value = style.Region},
                New SQLiteParameter("@TYPE", DbType.Int32) With {.Value = style.Type},
                New SQLiteParameter("@LOCALE", DbType.String) With {.Value = style.Locale},
                New SQLiteParameter("@PROP", DbType.String) With {.Value = style.Prop},
                New SQLiteParameter("@VAL", DbType.String) With {.Value = style.Val},
                New SQLiteParameter("@LASTCHANGE", DbType.DateTime) With {.Value = style.Lastchange}
            }
            parametersList.Add(parameters)
        Next
        ExecuteNonQuery(Sub(rst)

                        End Sub, sqls.ToArray, parametersList.ToArray)
    End Sub

    Public Function Read(style As Style) As Style Implements IStyleService.Read
        Dim sql As String = "SELECT REGION,TYPE,LOCALE,PROP,VAL,LASTCHANGE FROM STYLE_CUSTOM WHERE REGION=@REGION AND TYPE=@TYPE AND LOCALE=@LOCALE AND PROP=@PROP"
        Dim parameters As DbParameter() = {
            New SQLiteParameter("@REGION", DbType.Int32) With {.Value = style.Region},
            New SQLiteParameter("@TYPE", DbType.Int32) With {.Value = style.Type},
            New SQLiteParameter("@LOCALE", DbType.String) With {.Value = style.Locale},
            New SQLiteParameter("@PROP", DbType.String) With {.Value = style.Prop}
                                   }
        Dim rst As Style = Nothing
        ExecuteReader(Sub(reader As DbDataReader)
                          If reader.Read Then
                              rst = New StyleCustom()
                              rst.Region = reader.Item(0)
                              rst.Type = reader.Item(1)
                              rst.Locale = reader.Item(2)
                              rst.Prop = reader.Item(3)
                              rst.Val = reader.Item(4)
                              rst.Lastchange = reader.Item(5)
                          End If
                      End Sub, sql, parameters)

        Return rst
    End Function

    Public Function ReadPredefinedBorderStyle(region As RegionType) As PredefinedBorder Implements IStyleService.ReadPredefinedBorderStyle
        Dim rst As PredefinedBorder = Nothing
        For Each r In GetRegionInheritance(region)

            Dim sql As String = "WITH RECURSIVE GENERATE_SERIES(VALUE) AS (SELECT 3 UNION ALL SELECT VALUE + 1 FROM GENERATE_SERIES WHERE VALUE + 1 <= 10) " &
                                "SELECT T0.VALUE, IFNULL(T4.LINESTYLE,T8.LINESTYLE) AS LINESTYLE, IFNULL(T4.COLOR,T8.COLOR) AS COLOR, IFNULL(T4.WEIGHT,T8.WEIGHT) AS WEIGHT " &
                                "FROM GENERATE_SERIES T0 " &
                                "LEFT JOIN ( " &
                                " SELECT T1.TYPE, T1.VAL AS LINESTYLE, T2.VAL AS COLOR, T3.VAL AS WEIGHT FROM STYLE_CUSTOM T1 " &
                                " LEFT JOIN STYLE_CUSTOM T2 ON T1.REGION = T2.REGION AND T1.TYPE = T2.TYPE AND T1.LOCALE = T2.LOCALE AND T2.PROP = 'Color' " &
                                " LEFT JOIN STYLE_CUSTOM T3 ON T3.REGION = T2.REGION AND T3.TYPE = T2.TYPE AND T3.LOCALE = T2.LOCALE AND T3.PROP = 'Weight' " &
                                " WHERE T1.REGION=@REGION AND T1.LOCALE=@LOCALE AND T1.PROP='LineStyle' " &
                                ") T4 " &
                                "ON T0.VALUE = T4.TYPE " &
                                "LEFT JOIN ( " &
                                " SELECT T5.TYPE, T5.VAL AS LINESTYLE, T6.VAL AS COLOR, T7.VAL AS WEIGHT FROM STYLE_PRESET T5 " &
                                " LEFT JOIN STYLE_PRESET T6 ON T5.REGION = T6.REGION AND T5.TYPE = T6.TYPE AND T5.LOCALE = T6.LOCALE AND T6.PROP = 'Color' " &
                                " LEFT JOIN STYLE_PRESET T7 ON T5.REGION = T7.REGION AND T5.TYPE = T7.TYPE AND T5.LOCALE = T7.LOCALE AND T7.PROP = 'Weight' " &
                                " WHERE T5.REGION=@REGION AND T5.LOCALE=@LOCALE AND T5.PROP='LineStyle' " &
                                ") T8 " &
                                "ON T0.VALUE = T8.TYPE "
            Dim parameters As New List(Of Object) From {
                New SQLiteParameter("@REGION", DbType.Int32) With {.Value = r},
                New SQLiteParameter("@LOCALE", DbType.String) With {.Value = Globals.ThisAddIn.StrLangCode}
            }
            Dim borderAll = {StyleType.ST_BORDER_EDGE_LEFT, StyleType.ST_BORDER_EDGE_TOP, StyleType.ST_BORDER_EDGE_RIGHT, StyleType.ST_BORDER_EDGE_BOTTOM, StyleType.ST_BORDER_INSIDE_HORIZONTAL, StyleType.ST_BORDER_INSIDE_VERTICAL}
            Dim borderOutside = {StyleType.ST_BORDER_EDGE_LEFT, StyleType.ST_BORDER_EDGE_TOP, StyleType.ST_BORDER_EDGE_RIGHT, StyleType.ST_BORDER_EDGE_BOTTOM}
            Dim borderInside = {StyleType.ST_BORDER_INSIDE_HORIZONTAL, StyleType.ST_BORDER_INSIDE_VERTICAL}
            ExecuteReader(Sub(reader)
                              Dim styleList As New List(Of Object())
                              While reader.Read
                                  styleList.Add({
                                                reader.Item(0),
                                                reader.Item(1),
                                                reader.Item(2),
                                                reader.Item(3)
                                  })
                              End While
                              Dim edgeLeft As Object() = styleList.Find(Function(row) [Enum].Parse(GetType(StyleType), row(0)) = StyleType.ST_BORDER_EDGE_LEFT)
                              Dim edgeTop As Object() = styleList.Find(Function(row) [Enum].Parse(GetType(StyleType), row(0)) = StyleType.ST_BORDER_EDGE_TOP)
                              Dim edgeRight As Object() = styleList.Find(Function(row) [Enum].Parse(GetType(StyleType), row(0)) = StyleType.ST_BORDER_EDGE_RIGHT)
                              Dim edgeBottom As Object() = styleList.Find(Function(row) [Enum].Parse(GetType(StyleType), row(0)) = StyleType.ST_BORDER_EDGE_BOTTOM)
                              Dim insideHorizontal As Object() = styleList.Find(Function(row) [Enum].Parse(GetType(StyleType), row(0)) = StyleType.ST_BORDER_INSIDE_HORIZONTAL)
                              Dim insideVertical As Object() = styleList.Find(Function(row) [Enum].Parse(GetType(StyleType), row(0)) = StyleType.ST_BORDER_INSIDE_VERTICAL)
                              Dim diagonalUp As Object() = styleList.Find(Function(row) [Enum].Parse(GetType(StyleType), row(0)) = StyleType.ST_BORDER_DIAGONAL_UP)
                              Dim diagonalDown As Object() = styleList.Find(Function(row) [Enum].Parse(GetType(StyleType), row(0)) = StyleType.ST_BORDER_DIAGONAL_DOWN)

                              Dim maxLineStyle As Integer, minLineStyle As Integer
                              Dim maxLineColor As String, minLineColor As String
                              Dim maxLineWeight As Integer, minLineWeight As Integer

                              ' border style not set
                              If styleList.FindAll(Function(s) IsNull(s(1))).Count = 8 Then
                                  Return
                              End If

                              ' border clear
                              maxLineStyle = styleList.Max(Of Integer)(Function(s) If(IsNull(s(1)), 9999, s(1)))
                              minLineStyle = styleList.Min(Of Integer)(Function(s) If(IsNull(s(1)), -9999, s(1)))
                              If maxLineStyle = minLineStyle And maxLineStyle = Excel.XlLineStyle.xlLineStyleNone Then
                                  rst = New PredefinedBorder()
                              End If


                              ' border all
                              If (IsNull(diagonalUp(1)) OrElse [Enum].Parse(GetType(Excel.XlLineStyle), diagonalUp(1)) = Excel.XlLineStyle.xlLineStyleNone) AndAlso
                                (IsNull(diagonalDown(1)) OrElse [Enum].Parse(GetType(Excel.XlLineStyle), diagonalDown(1)) = Excel.XlLineStyle.xlLineStyleNone) Then
                                  Dim styleAllList = styleList.FindAll(Function(row) borderAll.Contains([Enum].Parse(GetType(StyleType), row(0))))
                                  maxLineStyle = styleAllList.Max(Of Integer)(Function(s) If(IsNull(s(1)), 9999, s(1)))
                                  minLineStyle = styleAllList.Min(Of Integer)(Function(s) If(IsNull(s(1)), -9999, s(1)))
                                  maxLineColor = styleAllList.Max(Of String)(Function(s) If(IsNull(s(2)), " ", s(2)))
                                  minLineColor = styleAllList.Min(Of String)(Function(s) If(IsNull(s(2)), "~", s(2)))
                                  maxLineWeight = styleAllList.Max(Of Integer)(Function(s) If(IsNull(s(3)), 9999, s(3)))
                                  minLineWeight = styleAllList.Min(Of Integer)(Function(s) If(IsNull(s(3)), -9999, s(3)))
                                  If maxLineStyle = minLineStyle And maxLineStyle = Excel.XlLineStyle.xlContinuous And
                                    maxLineColor = minLineColor And maxLineWeight = minLineWeight Then
                                      rst = New PredefinedBorder(PredefinedBorderStyle.BORDER_ALL, GetSystemColor(maxLineColor))
                                  End If
                              End If

                              ' border outside
                              If (IsNull(diagonalUp(1)) OrElse [Enum].Parse(GetType(Excel.XlLineStyle), diagonalUp(1)) = Excel.XlLineStyle.xlLineStyleNone) AndAlso
                                (IsNull(diagonalDown(1)) OrElse [Enum].Parse(GetType(Excel.XlLineStyle), diagonalDown(1)) = Excel.XlLineStyle.xlLineStyleNone) AndAlso
                                (IsNull(insideHorizontal(1)) OrElse [Enum].Parse(GetType(Excel.XlLineStyle), insideHorizontal(1)) = Excel.XlLineStyle.xlLineStyleNone) AndAlso
                                (IsNull(insideVertical(1)) OrElse [Enum].Parse(GetType(Excel.XlLineStyle), insideVertical(1)) = Excel.XlLineStyle.xlLineStyleNone) Then
                                  Dim styleOutsideList = styleList.FindAll(Function(row) borderOutside.Contains([Enum].Parse(GetType(StyleType), row(0))))
                                  maxLineStyle = styleOutsideList.Max(Of Integer)(Function(s) If(IsNull(s(1)), 9999, s(1)))
                                  minLineStyle = styleOutsideList.Min(Of Integer)(Function(s) If(IsNull(s(1)), -9999, s(1)))
                                  maxLineColor = styleOutsideList.Max(Of String)(Function(s) If(IsNull(s(2)), " ", s(2)))
                                  minLineColor = styleOutsideList.Min(Of String)(Function(s) If(IsNull(s(2)), "~", s(2)))
                                  maxLineWeight = styleOutsideList.Max(Of Integer)(Function(s) If(IsNull(s(3)), 9999, s(3)))
                                  minLineWeight = styleOutsideList.Min(Of Integer)(Function(s) If(IsNull(s(3)), -9999, s(3)))
                                  If maxLineStyle = minLineStyle And maxLineStyle = Excel.XlLineStyle.xlContinuous And
                                    maxLineColor = minLineColor And maxLineWeight = minLineWeight Then
                                      rst = New PredefinedBorder(PredefinedBorderStyle.BORDER_OUTER, GetSystemColor(maxLineColor))
                                  End If
                              End If

                              ' border inside
                              If (IsNull(diagonalUp(1)) OrElse [Enum].Parse(GetType(Excel.XlLineStyle), diagonalUp(1)) = Excel.XlLineStyle.xlLineStyleNone) AndAlso
                                (IsNull(diagonalDown(1)) OrElse [Enum].Parse(GetType(Excel.XlLineStyle), diagonalDown(1)) = Excel.XlLineStyle.xlLineStyleNone) AndAlso
                                (IsNull(edgeLeft(1)) OrElse [Enum].Parse(GetType(Excel.XlLineStyle), edgeLeft(1)) = Excel.XlLineStyle.xlLineStyleNone) AndAlso
                                (IsNull(edgeTop(1)) OrElse [Enum].Parse(GetType(Excel.XlLineStyle), edgeTop(1)) = Excel.XlLineStyle.xlLineStyleNone) AndAlso
                                (IsNull(edgeRight(1)) OrElse [Enum].Parse(GetType(Excel.XlLineStyle), edgeRight(1)) = Excel.XlLineStyle.xlLineStyleNone) AndAlso
                                (IsNull(edgeBottom(1)) OrElse [Enum].Parse(GetType(Excel.XlLineStyle), edgeBottom(1)) = Excel.XlLineStyle.xlLineStyleNone) Then
                                  Dim styleInsideList = styleList.FindAll(Function(row) borderInside.Contains([Enum].Parse(GetType(StyleType), row(0))))
                                  maxLineStyle = styleInsideList.Max(Of Integer)(Function(s) If(IsNull(s(1)), 9999, s(1)))
                                  minLineStyle = styleInsideList.Min(Of Integer)(Function(s) If(IsNull(s(1)), -9999, s(1)))
                                  maxLineColor = styleInsideList.Max(Of String)(Function(s) If(IsNull(s(2)), " ", s(2)))
                                  minLineColor = styleInsideList.Min(Of String)(Function(s) If(IsNull(s(2)), "~", s(2)))
                                  maxLineWeight = styleInsideList.Max(Of Integer)(Function(s) If(IsNull(s(3)), 9999, s(3)))
                                  minLineWeight = styleInsideList.Min(Of Integer)(Function(s) If(IsNull(s(3)), -9999, s(3)))
                                  If maxLineStyle = minLineStyle And maxLineStyle = Excel.XlLineStyle.xlContinuous And
                                    maxLineColor = minLineColor And maxLineWeight = minLineWeight Then
                                      rst = New PredefinedBorder(PredefinedBorderStyle.BORDER_INNER, GetSystemColor(maxLineColor))
                                  End If
                              End If

                              If rst Is Nothing Then rst = New PredefinedBorder(PredefinedBorderStyle.BORDER_CUSTOM)
                          End Sub, sql, parameters.ToArray)
            If rst IsNot Nothing Then Exit For
        Next
        If rst Is Nothing Then rst = New PredefinedBorder(PredefinedBorderStyle.BORDER_CLEAR)
        Return rst
    End Function
    Public Sub Reset(Optional region As RegionType? = Nothing, Optional locale As String = Nothing) Implements IStyleService.Reset
        Dim sql As String = "DELETE FROM STYLE_CUSTOM WHERE REGION=@REGION"
        Dim parameters As New List(Of DbParameter) From {
            New SQLiteParameter("@REGION", DbType.Int32) With {.Value = If(region Is Nothing, RegionType.RT_TABLE_GLOBAL, region)}
                                   }
        If locale IsNot Nothing Then
            sql += " AND LOCALE=@LOCALE"
            parameters.Add(New SQLiteParameter("@LOCALE", DbType.String) With {.Value = locale})
        End If
        ExecuteNonQuery(Sub(rst)
                            logger.Debug(If(region Is Nothing, "All regions were reset.", "Region " & [Enum].GetName(GetType(RegionType), region)) & " was reset.")
                        End Sub, sql, parameters.ToArray)
    End Sub

End Class
