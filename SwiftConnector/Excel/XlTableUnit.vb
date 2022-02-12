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
''' 表单元对象(Excel)
''' </summary>
Public MustInherit Class XlTableUnit
    Implements IXlTableUnit

    Protected Shared ReadOnly StyleService As IStyleService = New StyleService
    Protected Shared ReadOnly TextService As ITextService = New TextService
    Protected Shared ReadOnly DapperService As IDapperService = New DapperService
    Protected Shared ReadOnly ConfigService As IConfigService = New ConfigService

    Protected sqlCmd As ISqlCmdFactory = AbstractSqlCmdFactory.CreateFactory

    Protected dbTypeTranslator As IDbTypeTranslatorFactory = AbstractDbTypeTranslatorFactory.CreateFactory

    Private _regionInheritance As RegionType()
    Public ReadOnly Property RegionInheritance() As RegionType() Implements IXlTableUnit.RegionInheritance
        Get
            If _regionInheritance Is Nothing Then _regionInheritance = GetRegionInheritance(RegionType)
            Return _regionInheritance
        End Get
    End Property

    Private _regionType As RegionType
    Public Property RegionType As RegionType Implements IXlTableUnit.RegionType
        Get
            Return _regionType
        End Get
        Set(value As RegionType)
            _regionType = value
        End Set
    End Property

    Public MustOverride Sub Render(Optional mode As RenderMode = RenderMode.Excel) Implements IXlTableUnit.Render

    Public MustOverride Sub Revoke() Implements IXlTableUnit.Revoke

End Class
