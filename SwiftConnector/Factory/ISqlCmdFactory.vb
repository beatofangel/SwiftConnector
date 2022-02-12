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

Public Interface ISqlCmdFactory

    ''' <summary>
    ''' SQL:根据表注释获取表定义
    ''' </summary>
    ''' <param name="tblName">表注释</param>
    ''' <returns></returns>
    Function SqlQueryTblDefByTableName(tblName As String) As String

    ''' <summary>
    ''' SQL:根据表名获取表定义
    ''' </summary>
    ''' <param name="tblId">表名</param>
    ''' <returns></returns>
    Function SqlQueryTblDefByTableId(tblId As String) As String

    ''' <summary>
    ''' SQL:根据表名获取表字段定义
    ''' </summary>
    ''' <param name="tblId">表名</param>
    ''' <returns></returns>
    Function SqlQueryColDefByTableId(tblId As String) As String

    ''' <summary>
    ''' SQL:根据表名获取主键定义
    ''' </summary>
    ''' <param name="tblId">表名</param>
    ''' <returns></returns>
    Function SqlQueryPkDefByTableId(tblId As String) As String

    '''' <summary>
    '''' SQL:根据表名获取表数据
    '''' </summary>
    '''' <param name="tblId">表名</param>
    '''' <returns></returns>
    'Function SqlQueryRecordByTableId(tblId As String) As String

    ''' <summary>
    ''' SQL:根据表名获取查询表数据的SQL
    ''' </summary>
    ''' <param name="tblId"></param>
    ''' <returns></returns>
    Function SqlQueryRecordSqlByTableId(tblId As String) As String

    ''' <summary>
    ''' SQL:根据表名获取表类型(表,视图等)
    ''' </summary>
    ''' <param name="tblId"></param>
    ''' <returns></returns>
    Function SqlQueryRowHeaderText(tblId As String) As String

    ''' <summary>
    ''' SQL:根据表名、列名，保存数据
    ''' </summary>
    ''' <param name="tblId">表名</param>
    ''' <param name="colNames">列名</param>
    ''' <param name="prefix">参数前缀</param>
    ''' <returns></returns>
    Function SqlInsertRecord(tblId As String, colNames As List(Of String), Optional prefix As String = "") As String

    ''' <summary>
    ''' SQL:根据表名删除行数据
    ''' </summary>
    ''' <param name="tblId">表名</param>
    ''' <param name="pkLikeColNames">类主键字段名</param>
    ''' <param name="prefix">参数前缀</param>
    ''' <returns></returns>
    Function SqlDeleteTable(tblId As String, pkLikeColNames As List(Of String), Optional prefix As String = "") As String

    ''' <summary>
    ''' SQL:根据表名清空表数据
    ''' </summary>
    ''' <param name="tblId">表名</param>
    ''' <returns></returns>
    Function SqlTruncateTable(tblId As String) As String

End Interface
