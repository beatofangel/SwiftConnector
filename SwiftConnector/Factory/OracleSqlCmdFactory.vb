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


Public Class OracleSqlCmdFactory
    Inherits BaseService
    Implements ISqlCmdFactory

    ' TODO 改善为可配置项
    Public Shared Property DateFormat As String = "YYYY/MM/DD HH24:MI:SS"

    ' TODO 改善为可配置项
    Public Shared Property TimeStampFormat As String = "YYYY/MM/DD HH24:MI:SS.FF"

    ' TODO 改善为可配置项
    Public Shared Property TimeZoneFormat As String = "TZH"

    Private configService As New ConfigService

    Public Function SqlDeleteTable(tblId As String, pkLikeColNames As List(Of String), Optional prefix As String = "") As String Implements ISqlCmdFactory.SqlDeleteTable
        Dim sql As String
        Dim whereConditions As New List(Of String)
        For idx = 0 To pkLikeColNames.Count
            whereConditions.AddRange(pkLikeColNames.Select(Function(e) e & "=" & prefix & e))
        Next
        sql = String.Format("DELETE FROM {0} WHERE {1}", tblId.ToUpper, String.Join(" AND ", pkLikeColNames.Select(Function(e) e & "=" & prefix & e)))

        Return sql
    End Function

    Public Function SqlInsertRecord(tblId As String, colNames As List(Of String), Optional prefix As String = "") As String Implements ISqlCmdFactory.SqlInsertRecord
        Dim sql As String
        sql = String.Format("INSERT INTO {0} ({1}) VALUES ({2})",
                            tblId.ToUpper, String.Join(",", colNames),
                            String.Join(",", colNames.Select(Function(e) prefix & e)))
        Return sql
    End Function

    Public Function SqlQueryColDefByTableId(tblId As String) As String Implements ISqlCmdFactory.SqlQueryColDefByTableId
        Dim sql As String
        sql = "SELECT T1.COLUMN_NAME, T2.COMMENTS, T1.DATA_TYPE, T1.DATA_LENGTH, T1.DATA_PRECISION, T1.DATA_SCALE, T1.NULLABLE FROM USER_TAB_COLUMNS T1 INNER JOIN USER_COL_COMMENTS T2 ON T2.COLUMN_NAME = T1.COLUMN_NAME AND T2.TABLE_NAME = T1.TABLE_NAME WHERE T1.TABLE_NAME = '" & tblId.ToUpper & "' ORDER BY T1.COLUMN_ID"

        Return sql
    End Function

    Public Function SqlQueryPkDefByTableId(tblId As String) As String Implements ISqlCmdFactory.SqlQueryPkDefByTableId
        Dim sql As String
        sql = "SELECT T1.COLUMN_NAME FROM USER_CONS_COLUMNS T1, USER_CONSTRAINTS T2 WHERE T1.CONSTRAINT_NAME = T2.CONSTRAINT_NAME AND T2.CONSTRAINT_TYPE = 'P' AND T2.TABLE_NAME = '" & tblId.ToUpper & "'"

        Return sql
    End Function

    'Public Function SqlQueryRecordByTableId(tblId As String) As String Implements ISqlCmdFactory.SqlQueryRecordByTableId
    '    Throw New NotImplementedException()
    'End Function

    Public Function SqlQueryRecordSqlByTableId(tblId As String) As String Implements ISqlCmdFactory.SqlQueryRecordSqlByTableId
        Dim sql As String
        Dim strRecordLimit As String = Nothing
        Dim recordLimit As Integer = configService.GetRecordLimit
        If recordLimit > 0 Then
            logger.Debug("Record limit set to " & CStr(recordLimit))
            strRecordLimit = " || ' WHERE ROWNUM<=" & CStr(recordLimit) & "'"
        End If
        sql = "SELECT 'SELECT '  || cols  || ' FROM '  || '" & tblId.ToUpper & "'" & strRecordLimit & " " &
          "FROM (SELECT listagg(col,',') within GROUP(ORDER BY column_id) AS cols " &
                "FROM (SELECT COLUMN_ID, " &
                "             CASE " &
                "             WHEN user_tab_columns.data_type = 'RAW' " &
                "             THEN 'RAWTOHEX(' || user_tab_columns.column_name || ')' " &
                "             WHEN user_tab_columns.data_type = 'DATE' " &
                "             THEN 'TO_CHAR(' || user_tab_columns.column_name || ',''" & DateFormat & "'')' " &
                "             WHEN REGEXP_INSTR(user_tab_columns.data_type,'^TIMESTAMP\((\d)\)$',1,1,0,'i')=1 " &
                "             THEN 'TO_CHAR(' || user_tab_columns.column_name || ',''" & TimeStampFormat & "' || REGEXP_SUBSTR(user_tab_columns.data_type, '^TIMESTAMP\((\d)\)$',1,1,'i',1) || ''')' " &
                "             WHEN REGEXP_INSTR(user_tab_columns.data_type,'^TIMESTAMP\((\d)\) WITH TIME ZONE$',1,1,0,'i')=1 " &
                "             THEN 'TO_CHAR(' || user_tab_columns.column_name || ',''" & TimeStampFormat & "' || REGEXP_SUBSTR(user_tab_columns.data_type, '^TIMESTAMP\((\d)\) WITH TIME ZONE$',1,1,'i',1) || ' " & TimeZoneFormat & "' || ''')' " &
                "             WHEN REGEXP_INSTR(user_tab_columns.data_type,'^TIMESTAMP\((\d)\) WITH LOCAL TIME ZONE$',1,1,0,'i')=1 " &
                "             THEN 'TO_CHAR(' || user_tab_columns.column_name || ',''" & TimeStampFormat & "' || REGEXP_SUBSTR(user_tab_columns.data_type, '^TIMESTAMP\((\d)\) WITH LOCAL TIME ZONE$',1,1,'i',1) || ''')' " &
                "             ELSE user_tab_columns.column_name " &
                "             END AS COL, " &
                "     user_tab_columns.data_type " &
                "     FROM user_tab_columns " &
                "WHERE user_tab_columns.table_name = '" & UCase(tblId) & "' " &
          "    ) " &
          ")"

        Return sql
    End Function

    Public Function SqlQueryRowHeaderText(tblId As String) As String Implements ISqlCmdFactory.SqlQueryRowHeaderText
        Throw New NotImplementedException()
    End Function

    Public Function SqlQueryTblDefByTableId(tblId As String) As String Implements ISqlCmdFactory.SqlQueryTblDefByTableId
        Dim sql As String
        sql = "SELECT TABLE_NAME,TABLE_TYPE,COMMENTS FROM USER_TAB_COMMENTS WHERE TABLE_NAME = '" & tblId.ToUpper & "'"

        Return sql
    End Function

    Public Function SqlQueryTblDefByTableName(tblName As String) As String Implements ISqlCmdFactory.SqlQueryTblDefByTableName
        Dim sql As String
        sql = "SELECT TABLE_NAME,TABLE_TYPE,COMMENTS FROM USER_TAB_COMMENTS WHERE UPPER(COMMENTS) LIKE '%' || '" & tblName.ToUpper & "' || '%' AND ROWNUM <= " & CStr(TableCommentLikeQueryRowLimit) & " ORDER BY TABLE_NAME"

        Return sql
    End Function

    Public Function SqlTruncateTable(tblId As String) As String Implements ISqlCmdFactory.SqlTruncateTable
        Dim sql As String
        sql = "TRUNCATE TABLE " & tblId.ToUpper

        Return sql
    End Function

End Class
