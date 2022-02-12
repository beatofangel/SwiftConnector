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


Public Class MySQLCmdFactory
    Inherits BaseService
    Implements ISqlCmdFactory

    ' TODO 改善为可配置项
    Public Shared Property DateFormat As String = "%Y/%m/%d %H:%i:%S"

    ' TODO 改善为可配置项
    Public Shared Property TimeStampFormat As String = "%Y/%m/%d %H:%i:%S.%f"

    Private configService As New ConfigService

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

    Public Function SqlInsertRecord(tblId As String, colNames As List(Of String), Optional prefix As String = "") As String Implements ISqlCmdFactory.SqlInsertRecord
        Dim sql As String
        Dim prefix2 = If(String.IsNullOrEmpty(prefix), "@", prefix)
        sql = String.Format("INSERT INTO {0} ({1}) VALUES ({2})",
                            tblId.ToUpper, String.Join(",", colNames),
                            String.Join(",", colNames.Select(Function(e) prefix2 & e)))
        Return sql
    End Function

    Public Function SqlQueryColDefByTableId(tblId As String) As String Implements ISqlCmdFactory.SqlQueryColDefByTableId
        Dim sql As String
        sql = "SELECT T1.COLUMN_NAME, T1.COLUMN_COMMENT COMMENTS, T1.DATA_TYPE DATA_TYPE, CASE WHEN T1.DATA_TYPE IN ('timestamp','datetime') THEN T1.DATETIME_PRECISION ELSE T1.CHARACTER_MAXIMUM_LENGTH END DATA_LENGTH, T1.NUMERIC_PRECISION DATA_PRECISION, T1.NUMERIC_SCALE DATA_SCALE, CASE T1.IS_NULLABLE WHEN 'YES' THEN 'Y' WHEN 'NO' THEN 'N' END NULLABLE FROM INFORMATION_SCHEMA.COLUMNS T1 WHERE T1.TABLE_SCHEMA = '" & Globals.ThisAddIn.CurDataSource.Database & "' AND T1.TABLE_NAME = '" & tblId.ToUpper & "'ORDER BY T1.ORDINAL_POSITION"

        Return sql
    End Function

    Public Function SqlQueryPkDefByTableId(tblId As String) As String Implements ISqlCmdFactory.SqlQueryPkDefByTableId
        Dim sql As String
        sql = "SELECT T1.COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS T1 WHERE T1.TABLE_SCHEMA = '" & Globals.ThisAddIn.CurDataSource.Database & "' AND T1.TABLE_NAME = '" & tblId.ToUpper & "' AND T1.COLUMN_KEY='PRI'"

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
            strRecordLimit = "LIMIT " & CStr(recordLimit)
        End If

        sql = "SELECT CONCAT('SELECT ', T3.cols, ' FROM ', '" & tblId.ToUpper & " ', '" & strRecordLimit & "')" & " " &
            "FROM (SELECT GROUP_CONCAT(T2.COL ORDER BY T2.ORDINAL_POSITION) AS COLS" & " " &
            "    FROM (SELECT T1.COLUMN_NAME," & " " &
            "                 CASE" & " " &
            "                 WHEN T1.DATA_TYPE IN ('datetime','timestamp')" & " " &
            "                 THEN CASE WHEN DATETIME_PRECISION = 0" & " " &
            "                           THEN CONCAT('DATE_FORMAT(', T1.COLUMN_NAME, ',''" & DateFormat & "'')')" & " " &
            "                           ELSE CONCAT('DATE_FORMAT(', T1.COLUMN_NAME, ',''" & TimeStampFormat & "'')')" & " " &
            "                           END" & " " &
            "                 ELSE T1.COLUMN_NAME" & " " &
            "                 END AS COL," & " " &
            "                 T1.DATA_TYPE," & " " &
            "                 T1.ORDINAL_POSITION" & " " &
            "          FROM INFORMATION_SCHEMA.COLUMNS T1" & " " &
            "          WHERE T1.TABLE_SCHEMA = '" & Globals.ThisAddIn.CurDataSource.Database & "' AND T1.TABLE_NAME = '" & tblId.ToUpper & "') T2" & " " &
            "    ) T3"

        Return sql
    End Function

    Public Function SqlQueryRowHeaderText(tblId As String) As String Implements ISqlCmdFactory.SqlQueryRowHeaderText
        Throw New NotImplementedException()
    End Function

    Public Function SqlQueryTblDefByTableId(tblId As String) As String Implements ISqlCmdFactory.SqlQueryTblDefByTableId
        Dim sql As String
        sql = "SELECT TABLE_NAME, TABLE_TYPE, TABLE_COMMENT FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = '" & Globals.ThisAddIn.CurDataSource.Database & "'AND TABLE_NAME = '" & UCase(tblId) & "'"

        Return sql
    End Function

    Public Function SqlQueryTblDefByTableName(tblName As String) As String Implements ISqlCmdFactory.SqlQueryTblDefByTableName
        Dim sql As String
        sql = "SELECT TABLE_NAME, TABLE_TYPE, TABLE_COMMENT FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = '" & Globals.ThisAddIn.CurDataSource.Database & "'AND UPPER(TABLE_COMMENT) LIKE CONCAT('%', '" & tblName.ToUpper & "', '%') ORDER BY TABLE_NAME"

        Return sql
    End Function

    Public Function SqlTruncateTable(tblId As String) As String Implements ISqlCmdFactory.SqlTruncateTable
        Dim sql As String
        sql = "TRUNCATE TABLE " & tblId.ToUpper

        Return sql
    End Function

End Class
