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
Imports System.Globalization

''' <summary>
''' SQLite读写业务层
''' </summary>
Public Class DatasourceService
    Inherits SQLiteBaseService
    Implements IDatesourceService

    Public Function AddDataSource(ds As DataSource) As Integer Implements IDatesourceService.AddDataSource
        Dim rst As Integer
        Dim sql As String = "INSERT INTO DATASOURCE(ID,CURRENT,TYPE,NAME,DATASOURCE,DATABASE,USERNAME,PASSWORD,MODE,PORT,LASTCHANGE) VALUES (@id,@current,@type,@name,@datasource,@database,@username,@password,@mode,@port,@lastchange)"
        ExecuteNonQuery(Sub(cnt)
                            rst = cnt
                            logger.Debug(cnt & " record inserted.")
                        End Sub, sql, New DbParameter() {
                            New SQLiteParameter("@id", ds.Id),
                            New SQLiteParameter("@current", ds.Current),
                            New SQLiteParameter("@type", ds.Type),
                            New SQLiteParameter("@name", ds.Name),
                            New SQLiteParameter("@datasource", ds.Datasource),
                            New SQLiteParameter("@database", ds.Database),
                            New SQLiteParameter("@username", ds.Username),
                            New SQLiteParameter("@password", ds.Password),
                            New SQLiteParameter("@mode", ds.Mode),
                            New SQLiteParameter("@port", ds.Port),
                            New SQLiteParameter("@lastchange", ds.Lastchange)
                        })

        Return rst
    End Function

    'Public Function ChangeConfig(cfg As Config) As Boolean Implements ISQLiteService.ChangeConfig
    '    Dim sql As String = "INSERT OR REPLACE INTO CONFIG (PROP,LOCALE,VAL,LASTCHANGE) VALUES (@PROP,@LOCALE,@VAL,@LASTCHANGE)"
    '    Dim parameters As DbParameter() = {
    '        New SQLiteParameter("@PROP", DbType.String) With {.Value = cfg.Prop},
    '        New SQLiteParameter("@LOCALE", DbType.String) With {.Value = cfg.Locale},
    '        New SQLiteParameter("@VAL", DbType.String) With {.Value = cfg.Val},
    '        New SQLiteParameter("@LASTCHANGE", DbType.DateTime) With {.Value = cfg.Lastchange}
    '                               }
    '    Dim rst As Boolean = False

    '    Try
    '        ExecuteNonQuery(Sub(execRst As Integer)
    '                            rst = execRst
    '                        End Sub, sql, parameters)
    '    Catch ex As Exception
    '        logger.Error(ex)
    '    End Try

    '    Return rst
    'End Function

    Public Function DeleteDataSource(ds As DataSource) As Integer Implements IDatesourceService.DeleteDataSource
        Dim rst As Integer
        Dim sql As String = "DELETE FROM DATASOURCE WHERE ID = @id"
        ExecuteNonQuery(Sub(cnt)
                            rst = cnt
                            logger.Debug(cnt & " record deleted.")
                        End Sub, sql, New DbParameter() {
                            New SQLiteParameter("@id", ds.Id)
                        })

        Return rst
    End Function

    Public Function EditDataSource(ds As DataSource) As Integer Implements IDatesourceService.EditDataSource
        Dim rst As Integer
        Dim sql As String = "UPDATE DATASOURCE SET TYPE=@type,NAME=@name,DATASOURCE=@datasource,DATABASE=@database,USERNAME=@username,PASSWORD=@password,MODE=@mode,PORT=@port,LASTCHANGE=@lastchange WHERE ID=@id"
        ExecuteNonQuery(Sub(cnt)
                            rst = cnt
                            logger.Debug(cnt & " record updated.")
                        End Sub, sql, New DbParameter() {
                            New SQLiteParameter("@id", ds.Id),
                            New SQLiteParameter("@type", ds.Type),
                            New SQLiteParameter("@name", ds.Name),
                            New SQLiteParameter("@datasource", ds.Datasource),
                            New SQLiteParameter("@database", ds.Database),
                            New SQLiteParameter("@username", ds.Username),
                            New SQLiteParameter("@password", ds.Password),
                            New SQLiteParameter("@mode", ds.Mode),
                            New SQLiteParameter("@port", ds.Port),
                            New SQLiteParameter("@lastchange", ds.Lastchange)
                        })

        Return rst
    End Function

    ''' <summary>
    ''' 查询DataSource表全部数据
    ''' </summary>
    ''' <returns>List(Of <see cref="DataSource"/>)</returns>
    Public Function FindAllDataSource() As List(Of DataSource) Implements IDatesourceService.FindAllDataSource
        Dim sql As String = "SELECT ID,CURRENT,TYPE,NAME,DATASOURCE,IFNULL(DATABASE,'') DATABASE,IFNULL(USERNAME,'') USERNAME,IFNULL(PASSWORD,'') PASSWORD,MODE,IFNULL(PORT,''),LASTCHANGE FROM DATASOURCE ORDER BY TYPE,NAME"
        Dim rst As List(Of DataSource) = Nothing

        Try
            ExecuteReader(Sub(reader As DbDataReader)
                              rst = New List(Of DataSource)
                              While reader.Read
                                  Dim ds = New DataSource() With {
                                                          .Id = reader.Item(0),
                                                          .Current = reader.Item(1),
                                                          .Type = reader.Item(2),
                                                          .Name = reader.Item(3),
                                                          .Datasource = reader.Item(4),
                                                          .Database = reader.Item(5),
                                                          .Username = reader.Item(6),
                                                          .Password = reader.Item(7),
                                                          .Mode = reader.Item(8),
                                                          .Port = reader.Item(9),
                                                          .Lastchange = reader.Item(10)
                                                         }
                                  rst.Add(ds)
                              End While
                          End Sub, sql)
        Catch ex As Exception
            logger.Error(ex)
        End Try

        Return rst
    End Function

    'Public Function ReadConfig(cfg As Config) As Config Implements ISQLiteService.ReadConfig
    '    Dim sql As String = "SELECT PROP,LOCALE,VAL,LASTCHANGE FROM STYLE_CUSTOM WHERE PROP=@PROP AND LOCALE=@LOCALE"
    '    Dim parameters As DbParameter() = {
    '        New SQLiteParameter("@PROP", DbType.String) With {.Value = cfg.Prop},
    '        New SQLiteParameter("@LOCALE", DbType.String) With {.Value = cfg.Locale}
    '                               }
    '    Dim rst As Config = Nothing

    '    ExecuteReader(Sub(reader)
    '                      If reader.Read Then
    '                          rst = New Config
    '                          rst.Prop = reader.Item(0)
    '                          rst.Locale = reader.Item(1)
    '                          rst.Val = reader.Item(2)
    '                          rst.Lastchange = reader.Item(3)
    '                      End If
    '                  End Sub, sql, parameters)
    '    Return rst
    'End Function

    ''' <summary>
    ''' 切换当前数据源到指定id
    ''' </summary>
    ''' <param name="id"></param>
    ''' <returns>Boolean</returns>
    Public Function SwitchDataSourceTo(id As String) As Boolean Implements IDatesourceService.SwitchDataSourceTo
        Dim sqlCancelCurrent As String = "UPDATE DATASOURCE SET CURRENT=FALSE WHERE CURRENT=TRUE"
        Dim sqlSwitch2New As String = "UPDATE DATASOURCE SET CURRENT=TRUE WHERE ID='" & id & "'"
        Dim sqlArr As String() = {sqlCancelCurrent, sqlSwitch2New}
        Dim rst As Boolean = False
        Try
            ExecuteNonQuery(Sub(execRst As List(Of Integer))
                                If execRst.Count > 0 Then
                                    rst = True
                                End If
                            End Sub, sqlArr)
        Catch ex As Exception
            logger.Error(ex)
        End Try

        Return rst
    End Function

    'Public Function IsShowColProps() As Boolean Implements ISQLiteService.IsShowColProps
    '    Dim rst As Boolean = False
    '    Dim showColProps = ReadConfig(New Config With {
    '                                    .Prop = "ShowColumnProperties",
    '                                    .Locale = "-"
    '                                  })
    '    If showColProps IsNot Nothing Then
    '        rst = CBool(showColProps.Val)
    '    End If

    '    Return rst
    'End Function

    'Public Function GetRecordLimit() As Integer Implements ISQLiteService.GetRecordLimit
    '    Dim rst As Integer = 0 ' no limit
    '    Dim recordLimit = ReadConfig(New Config With {
    '                                    .Prop = "RecordLimit",
    '                                    .Locale = "-"
    '                                  })
    '    If recordLimit IsNot Nothing Then
    '        rst = CInt(recordLimit.Val)
    '    End If

    '    Return rst
    'End Function

    'Public Function GetDelMode() As DeleteMode Implements ISQLiteService.GetDelMode
    '    Dim rst As DeleteMode = DeleteMode.Truncate
    '    Dim delMode = ReadConfig(New Config With {
    '                                    .Prop = "DeleteMode",
    '                                    .Locale = "-"
    '                                  })
    '    If delMode IsNot Nothing Then
    '        rst = [Enum].Parse(GetType(DeleteMode), delMode.Val)
    '    End If

    '    Return rst
    'End Function

    ''' <summary>
    ''' for 1.0.0.13 update
    ''' </summary>
    Public Sub Patch_1_0_0_13()
        Dim sqlColumnModeExists = "SELECT NAME FROM PRAGMA_TABLE_INFO('DATASOURCE') WHERE NAME='MODE'"
        Dim columnModeNotFound As Boolean = False
        ExecuteReader(Sub(reader)
                          columnModeNotFound = Not reader.Read
                      End Sub, sqlColumnModeExists)
        Dim sqlAddColumnMode = "ALTER TABLE DATASOURCE ADD MODE INTEGER NOT NULL DEFAULT 0"
        If columnModeNotFound Then
            ExecuteNonQuery(Sub(rst)
                                logger.Debug("add column MODE to table DATASOURCE: " & rst)
                            End Sub, sqlAddColumnMode)
        End If
    End Sub

    Public Sub Patch_1_0_0_15()
        Dim sqlColumnPortExists = "SELECT NAME FROM PRAGMA_TABLE_INFO('DATASOURCE') WHERE NAME='PORT'"
        Dim columnPortNotFound As Boolean = False
        ExecuteReader(Sub(reader)
                          columnPortNotFound = Not reader.Read
                      End Sub, sqlColumnPortExists)
        Dim sqlAddColumnPort = "ALTER TABLE DATASOURCE ADD PORT TEXT"
        If columnPortNotFound Then
            ExecuteNonQuery(Sub(rst)
                                logger.Debug("add column PORT to table DATASOURCE: " & rst)
                            End Sub, sqlAddColumnPort)
        End If
    End Sub
End Class
