Imports System.Text
Imports System.Runtime.InteropServices
Module WinApi
    Public Declare Function IsWindowEnabled Lib "user32.dll" (ByVal hwnd As Integer) As Long
    Public Declare Function IsWindowVisible Lib "user32.dll" (ByVal hwnd As Integer) As Long
    Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As StringBuilder, ByVal cch As Long) As Long
    Public Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
    Public Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    Public Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As StringBuilder, ByVal nMaxCount As Long) As Long
    Public Declare Function EnumChildWindows Lib "user32.dll" Alias "EnumChildWindows" (ByVal hWndParent As Long, ByVal callback As EnumWindowsCallback, ByVal lParam As Integer) As Integer
    Public Declare Sub SetWindowPos Lib "User32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
    Public Declare Function EnableWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
    Public Declare Function SetForegroundWindow Lib "user32" Alias "SetForegroundWindow" (ByVal hwnd As Long) As Long
    Public Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
    Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Object) As Long
    Public Declare Function GetParent Lib "user32.dll" Alias "GetParent" (ByVal hwnd As Integer) As Integer
    Public Declare Function GetWindow Lib "user32.dll" Alias "GetWindow" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
    Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Integer, ByVal nIndex As Integer) As Integer

    Public Structure POINTAPI
        Public X As Long
        Public Y As Long
    End Structure

    Public Function WinFromCursor() As Long
        Dim s As POINTAPI
        GetCursorPos(s)
        Return WindowFromPoint(s.X, s.Y)
    End Function

    Public Delegate Function EnumWindowsCallback(ByVal hWnd As Integer, ByVal lParam As Integer) As Boolean
    Public Declare Function EnumWindows Lib "user32.dll" Alias "EnumWindows" (ByVal callback As EnumWindowsCallback, ByVal lParam As Integer) As Integer
    <DllImport("user32.dll", EntryPoint:="EnumWindows", SetLastError:=True, CharSet:=CharSet.Ansi, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> Public Function EnumWindowsDllImport(ByVal callback As EnumWindowsCallback, ByVal lParam As Integer) As Integer
    End Function
   
End Module
