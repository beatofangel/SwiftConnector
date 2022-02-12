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

''' <summary>
''' SQLite读写业务层接口
''' </summary>
Public Interface IDatesourceService

#Region "DATASOURCE"
    ''' <summary>
    ''' 查询DataSource表全部数据
    ''' </summary>
    ''' <returns>List(Of <see cref="DataSource"/>)</returns>
    Function FindAllDataSource() As List(Of DataSource)

    ''' <summary>
    ''' 切换当前数据源到指定id
    ''' </summary>
    ''' <param name="id"></param>
    ''' <returns>Boolean</returns>
    Function SwitchDataSourceTo(id As String) As Boolean

    Function AddDataSource(ds As DataSource) As Integer

    Function EditDataSource(ds As DataSource) As Integer

    Function DeleteDataSource(ds As DataSource) As Integer

#End Region

End Interface
