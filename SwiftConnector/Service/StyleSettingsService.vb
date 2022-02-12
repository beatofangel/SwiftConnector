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

Public Class StyleSettingsService
    Inherits SQLiteBaseService
    Implements IStyleSettingsService

    Public Sub Change(style As Style) Implements IStyleSettingsService.Change
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

    ' TODO
    Public Sub BatchChange(styles As Style()) Implements IStyleSettingsService.BatchChange
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

    Public Function Read(style As Style) As Style Implements IStyleSettingsService.Read
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

    Public Function ReadPredefinedBorderStyle(region As RegionType) As PredefinedBorder Implements IStyleSettingsService.ReadPredefinedBorderStyle
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
    Public Sub Reset(Optional region As RegionType? = Nothing, Optional locale As String = Nothing) Implements IStyleSettingsService.Reset
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
