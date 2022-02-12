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

Imports System.Collections
Imports System.Diagnostics
Imports System.Runtime.InteropServices

Public Class XlHotKeys
    Implements IDictionary(Of String, XlHotKey)

    Public Delegate Function LowLevelKeyboardProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    Private ReadOnly _proc As LowLevelKeyboardProc = AddressOf HookCallback
    Private Shared _hookID As IntPtr = IntPtr.Zero
    Private Const WH_KEYBOARD As Integer = 2
    Private Const HC_ACTION As Integer = 0
    Private Shared cb As Action(Of String)
    Private Shared ReadOnly kc As New Windows.Forms.KeysConverter

    Private hotKeys As New Dictionary(Of String, XlHotKey)

    Public ReadOnly Property Count As Integer Implements ICollection(Of KeyValuePair(Of String, XlHotKey)).Count
        Get
            Return hotKeys.Count
        End Get
    End Property

    Public ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of KeyValuePair(Of String, XlHotKey)).IsReadOnly
        Get
            Return False
        End Get
    End Property

    Default Public Property Item(key As String) As XlHotKey Implements IDictionary(Of String, XlHotKey).Item
        Get
            Return hotKeys.Item(key)
        End Get
        Set(value As XlHotKey)
            hotKeys.Item(key) = value
        End Set
    End Property

    Public ReadOnly Property Keys As ICollection(Of String) Implements IDictionary(Of String, XlHotKey).Keys
        Get
            Return hotKeys.Keys
        End Get
    End Property

    Public ReadOnly Property Values As ICollection(Of XlHotKey) Implements IDictionary(Of String, XlHotKey).Values
        Get
            Return hotKeys.Values
        End Get
    End Property

    Public Sub Add(item As KeyValuePair(Of String, XlHotKey)) Implements ICollection(Of KeyValuePair(Of String, XlHotKey)).Add
        hotKeys.Add(item.Key, item.Value)
    End Sub

    Public Sub Add(key As String, value As XlHotKey) Implements IDictionary(Of String, XlHotKey).Add
        hotKeys.Add(key, value)
    End Sub

    Public Sub Clear() Implements ICollection(Of KeyValuePair(Of String, XlHotKey)).Clear
        hotKeys.Clear()
    End Sub

    Public Sub CopyTo(array() As KeyValuePair(Of String, XlHotKey), arrayIndex As Integer) Implements ICollection(Of KeyValuePair(Of String, XlHotKey)).CopyTo
        If array Is Nothing Then
            Throw New ArgumentNullException("array")
        End If
        Dim cpy As ICollection(Of KeyValuePair(Of String, XlHotKey)) = hotKeys
        cpy.CopyTo(array, arrayIndex)
    End Sub

    Public Function Contains(item As KeyValuePair(Of String, XlHotKey)) As Boolean Implements ICollection(Of KeyValuePair(Of String, XlHotKey)).Contains
        Return hotKeys.Contains(item)
    End Function

    Public Function ContainsKey(key As String) As Boolean Implements IDictionary(Of String, XlHotKey).ContainsKey
        Return hotKeys.ContainsKey(key)
    End Function

    Public Function GetEnumerator() As IEnumerator(Of KeyValuePair(Of String, XlHotKey)) Implements IEnumerable(Of KeyValuePair(Of String, XlHotKey)).GetEnumerator
        Return hotKeys.GetEnumerator()
    End Function

    Public Function Remove(item As KeyValuePair(Of String, XlHotKey)) As Boolean Implements ICollection(Of KeyValuePair(Of String, XlHotKey)).Remove
        Return hotKeys.Remove(item.Key)
    End Function

    Public Function Remove(key As String) As Boolean Implements IDictionary(Of String, XlHotKey).Remove
        Return hotKeys.Remove(key)
    End Function

    Public Function TryGetValue(key As String, ByRef value As XlHotKey) As Boolean Implements IDictionary(Of String, XlHotKey).TryGetValue
        Return hotKeys.TryGetValue(key, value)
    End Function

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return hotKeys.GetEnumerator()
    End Function

    Public Sub Bind(callback As Action(Of String))
        cb = callback
        SetHook()
    End Sub

    Public Sub Unbind()
        ReleaseHook()
        cb = Nothing
    End Sub

    Private Function DetectHotKey(k As Windows.Forms.Keys) As Boolean
        For Each hk As XlHotKey In hotKeys.Values
            If hk.ContainsKey(kc.ConvertToString(k)) Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Sub SetHook()
        _hookID = NativeMethods.SetWindowsHookEx(WH_KEYBOARD, _proc, IntPtr.Zero, NativeMethods.GetCurrentThreadId())
    End Sub

    Private Sub ReleaseHook()
        NativeMethods.UnhookWindowsHookEx(_hookID)
    End Sub

    Private Function HookCallback(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
        Dim PreviousStateBit As Integer = 31
        Dim KeyWasAlreadyPressed As Boolean = False
        Dim bitmask As Int64 = CLng(Math.Pow(2, (PreviousStateBit - 1)))

        Try

            If nCode < 0 Then
                Return CInt(NativeMethods.CallNextHookEx(_hookID, nCode, wParam, lParam))
            Else
                If nCode = HC_ACTION Then
                    Dim keyData As Windows.Forms.Keys = CType(wParam, Windows.Forms.Keys)
                    KeyWasAlreadyPressed = (CLng(lParam) And bitmask) > 0
                    'If (IsKeyDown(Windows.Forms.Keys.ControlKey) Or IsKeyDown(Windows.Forms.Keys.ShiftKey)) And DetectHotKey(keyData) And Not KeyWasAlreadyPressed Then
                    If IsKeyDown(Windows.Forms.Keys.ControlKey) And DetectHotKey(keyData) And Not KeyWasAlreadyPressed Then
                        Dim xlHwnd = Globals.ThisAddIn.Application.Hwnd
                        Dim topHwnd = NativeMethods.GetForegroundWindow().ToInt32
                        ' 仅当excel窗口为当前窗口时 
                        If xlHwnd = topHwnd And Not IsEditing() Then
                            Dim keyList As New List(Of String)
                            If IsKeyDown(Windows.Forms.Keys.ControlKey) Then keyList.Add(KEY_CTRL)
                            If IsKeyDown(Windows.Forms.Keys.ShiftKey) Then keyList.Add(KEY_SHIFT)
                            keyList.Add(kc.ConvertToString(keyData))
                            cb.Invoke(String.Join("", keyList))
                        Else
                            Return CInt(NativeMethods.CallNextHookEx(_hookID, nCode, wParam, lParam))
                        End If

                        Return 1
                    End If
                End If
                Return CInt(NativeMethods.CallNextHookEx(_hookID, nCode, wParam, lParam))
            End If

        Catch ex As Exception
            'System.Windows.Forms.MessageBox.Show(ex.Message)
            Return CInt(NativeMethods.CallNextHookEx(_hookID, nCode, wParam, lParam))
        End Try
    End Function

    Private Shared Function IsKeyDown(ByVal keys As Windows.Forms.Keys) As Boolean
        Return (NativeMethods.GetKeyState(CInt(keys)) And &H8000) = &H8000
    End Function

    Friend NotInheritable Class NativeMethods

        <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Public Shared Function SetWindowsHookEx(ByVal idHook As Integer, ByVal lpfn As LowLevelKeyboardProc, ByVal hMod As IntPtr, ByVal dwThreadId As UInteger) As IntPtr
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Public Shared Function UnhookWindowsHookEx(ByVal hhk As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Public Shared Function CallNextHookEx(ByVal hhk As IntPtr, ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
        End Function

        <DllImport("kernel32.dll")>
        Public Shared Function GetCurrentThreadId() As UInteger
        End Function

        <DllImport("user32.dll")>
        Public Shared Function GetKeyState(ByVal nVirtKey As Integer) As Short
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function GetForegroundWindow() As IntPtr
        End Function
    End Class
End Class
