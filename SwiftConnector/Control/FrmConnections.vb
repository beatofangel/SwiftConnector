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

Imports System.Windows.Forms

Public Class FrmConnections
    Public Sub New()

        ' 此调用是设计器所必需的。
        InitializeComponent()

        ' 在 InitializeComponent() 调用之后添加任何初始化。
        Dim maxHeight = Screen.PrimaryScreen.WorkingArea.Height
        Dim h = If(maxHeight < 1024, maxHeight, 1024)
        Me.Size = New Drawing.Size(h / 1024 * 800, h)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.MaximumSize = Screen.PrimaryScreen.WorkingArea.Size
    End Sub

    'Private Sub FrmConnections_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    '    hb.Fragment = "connections"
    'End Sub

    Public Function showDialog2(fragment As String) As DialogResult
        hb.Fragment = fragment
        Return Me.ShowDialog()
    End Function

End Class