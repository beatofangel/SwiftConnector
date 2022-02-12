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

Public Class ConfigService
    Inherits SQLiteBaseService
    Implements IConfigService

    Private Const NO_LOCALE = "-"

    Public Function Change(cfg As Config) As Boolean Implements IConfigService.Change
        Dim sql As String = "INSERT OR REPLACE INTO CONFIG (PROP,LOCALE,VAL,LASTCHANGE) VALUES (@PROP,@LOCALE,@VAL,@LASTCHANGE)"
        Dim parameters As DbParameter() = {
            New SQLiteParameter("@PROP", DbType.String) With {.Value = cfg.Prop},
            New SQLiteParameter("@LOCALE", DbType.String) With {.Value = cfg.Locale},
            New SQLiteParameter("@VAL", DbType.String) With {.Value = cfg.Val},
            New SQLiteParameter("@LASTCHANGE", DbType.DateTime) With {.Value = cfg.Lastchange}
                                   }
        Dim rst As Boolean = False

        Try
            ExecuteNonQuery(Sub(execRst As Integer)
                                rst = execRst
                            End Sub, sql, parameters)
        Catch ex As Exception
            logger.Error(ex)
        End Try

        Return rst
    End Function

    Public Function Read(cfg As Config) As Config Implements IConfigService.Read
        Dim sql As String = "SELECT PROP,LOCALE,VAL,LASTCHANGE FROM CONFIG WHERE PROP=@PROP AND LOCALE=@LOCALE"
        Dim parameters As DbParameter() = {
            New SQLiteParameter("@PROP", DbType.String) With {.Value = cfg.Prop},
            New SQLiteParameter("@LOCALE", DbType.String) With {.Value = cfg.Locale}
                                   }
        Dim rst As Config = Nothing

        ExecuteReader(Sub(reader)
                          If reader.Read Then
                              rst = New Config
                              rst.Prop = reader.Item(0)
                              rst.Locale = reader.Item(1)
                              rst.Val = reader.Item(2)
                              rst.Lastchange = reader.Item(3)
                          End If
                      End Sub, sql, parameters)
        Return rst
    End Function

    Public Function IsShowColProps() As Boolean Implements IConfigService.IsShowColProps
        Dim rst As Boolean = False
        Dim showColProps = Read(New Config With {
                                        .Prop = "ShowColumnProperties",
                                        .Locale = NO_LOCALE
                                      })
        If showColProps IsNot Nothing Then
            rst = CBool(showColProps.Val)
        End If

        Return rst
    End Function

    Public Function IsAutoFitColumns() As Boolean Implements IConfigService.IsAutoFitColumns
        Dim rst As Boolean = False
        Dim autoFitColumns = Read(New Config With {
                                        .Prop = "AutoFitColumns",
                                        .Locale = NO_LOCALE
                                      })
        If autoFitColumns IsNot Nothing Then
            rst = CBool(autoFitColumns.Val)
        End If

        Return rst
    End Function

    Public Function GetRecordLimit() As Integer Implements IConfigService.GetRecordLimit
        Dim rst As Integer = 0 ' no limit
        Dim recordLimit = Read(New Config With {
                                        .Prop = "RecordLimit",
                                        .Locale = NO_LOCALE
                                      })
        If recordLimit IsNot Nothing Then
            rst = CInt(recordLimit.Val)
        End If

        Return rst
    End Function

    Public Function GetDelMode() As DeleteMode Implements IConfigService.GetDelMode
        Dim rst As DeleteMode = DeleteMode.Truncate
        Dim delMode = Read(New Config With {
                                        .Prop = "DeleteMode",
                                        .Locale = NO_LOCALE
                                      })
        If delMode IsNot Nothing Then
            rst = [Enum].Parse(GetType(DeleteMode), delMode.Val)
        End If

        Return rst
    End Function

    Public Function GetExecMode() As ExecuteMode Implements IConfigService.GetExecMode
        Dim rst As ExecuteMode = ExecuteMode.Normal
        Dim execMode = Read(New Config With {
                                        .Prop = "ExecuteMode",
                                        .Locale = NO_LOCALE
                                      })
        If execMode IsNot Nothing Then
            rst = [Enum].Parse(GetType(ExecuteMode), execMode.Val)
        End If

        Return rst
    End Function

    Function GetVersion() As String Implements IConfigService.GetVersion
        Dim rst As String = ""
        Dim ver = Read(New Config With {
                                    .Prop = "Version",
                                    .Locale = NO_LOCALE
                       })
        If ver IsNot Nothing Then
            rst = ver.Val
        End If

        Return rst
    End Function
End Class
