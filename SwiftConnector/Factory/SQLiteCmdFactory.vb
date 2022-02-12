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

Public Class SQLiteCmdFactory
    Inherits BaseService
    Implements ISqlCmdFactory

    Private configService As New ConfigService

    Public Function SqlQueryTblDefByTableName(tblName As String) As String Implements ISqlCmdFactory.SqlQueryTblDefByTableName
        Dim sql As String
        sql = "SELECT NAME TABLE_NAME, TYPE TABLE_TYPE, NAME COMMENTS FROM SQLITE_MASTER WHERE TBL_NAME LIKE '%" & tblName.ToUpper & "%' AND TYPE IN ('table', 'view')"

        Return sql
    End Function

    Public Function SqlQueryTblDefByTableId(tblId As String) As String Implements ISqlCmdFactory.SqlQueryTblDefByTableId
        Dim sql As String
        sql = "SELECT NAME TABLE_NAME, TYPE TABLE_TYPE, NAME COMMENTS FROM SQLITE_MASTER WHERE TBL_NAME='" & tblId.ToUpper & "' AND TYPE IN ('table', 'view')"

        Return sql
    End Function

    Public Function SqlQueryColDefByTableId(tblId As String) As String Implements ISqlCmdFactory.SqlQueryColDefByTableId
        Dim sql As String
        sql = "Select NAME COLUMN_NAME, NAME COMMENTS, TYPE DATA_TYPE, 0 DATA_LENGTH, 0 DATA_PRECISION, 0 DATA_SCALE, Case ""NOTNULL"" When 1 Then 'N' ELSE 'Y' END NULLABLE FROM PRAGMA_TABLE_INFO('" & tblId.ToUpper & "')"

        Return sql
    End Function

    Public Function SqlQueryPkDefByTableId(tblId As String) As String Implements ISqlCmdFactory.SqlQueryPkDefByTableId
        Dim sql As String
        sql = "SELECT NAME COLUMN_NAME FROM PRAGMA_TABLE_INFO('" & tblId.ToUpper & "') WHERE PK = 1"

        Return sql
    End Function

    Public Function SqlQueryRecordSqlByTableId(tblId As String) As String Implements ISqlCmdFactory.SqlQueryRecordSqlByTableId
        Dim sql As String
        Dim strRecordLimit As String = Nothing
        Dim recordLimit As Integer = ConfigService.GetRecordLimit
        If recordLimit > 0 Then
            logger.Debug("Record limit set to " & CStr(recordLimit))
            strRecordLimit = "LIMIT " & CStr(recordLimit)
        End If

        sql = "SELECT 'SELECT ' || GROUP_CONCAT(COLUMN_NAME) || ' FROM ' || '" & tblId.ToUpper & " ' || '" & strRecordLimit & "'" & " " &
            "FROM (" & " " &
            "  SELECT CASE TYPE" & " " &
            "         WHEN 'DATETIME' THEN 'strftime(""%Y-%m-%d %H:%M:%f"",' || NAME || ')'" & " " &
            "         ELSE NAME" & " " &
            "         END COLUMN_NAME FROM PRAGMA_TABLE_INFO('" & tblId.ToUpper & "')" & " " &
            ")"

        Return sql
    End Function

    Public Function SqlQueryRowHeaderText(tblId As String) As String Implements ISqlCmdFactory.SqlQueryRowHeaderText
        Throw New NotImplementedException()
    End Function

    Public Function SqlInsertRecord(tblId As String, colNames As List(Of String), Optional prefix As String = "") As String Implements ISqlCmdFactory.SqlInsertRecord
        Dim sql As String
        Dim prefix2 = If(String.IsNullOrEmpty(prefix), "@", prefix)
        sql = String.Format("INSERT INTO {0} ({1}) VALUES ({2})",
                            tblId.ToUpper, String.Join(",", colNames),
                            String.Join(",", colNames.Select(Function(e) prefix2 & e)))
        Return sql
    End Function

    Public Function SqlDeleteTable(tblId As String, pkLikeColNames As List(Of String), Optional prefix As String = "") As String Implements ISqlCmdFactory.SqlDeleteTable
        Dim sql As String
        Dim prefix2 = If(String.IsNullOrEmpty(prefix), "@", prefix)
        Dim whereConditions As New List(Of String)
        For idx = 0 To pkLikeColNames.Count
            whereConditions.AddRange(pkLikeColNames.Select(Function(e) e & "=" & prefix2 & e))
        Next
        sql = String.Format("DELETE FROM {0} WHERE {1}", tblId.ToUpper, String.Join(" AND ", pkLikeColNames.Select(Function(e) e & "=" & prefix2 & e)))

        Return sql
    End Function

    Public Function SqlTruncateTable(tblId As String) As String Implements ISqlCmdFactory.SqlTruncateTable
        Dim sql As String
        sql = "DELETE FROM " & tblId.ToUpper

        Return sql
    End Function
End Class
