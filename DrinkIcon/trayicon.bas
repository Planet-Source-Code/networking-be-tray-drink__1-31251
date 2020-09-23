Attribute VB_Name = "TrayIcon"
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hiCon As Long
        szTip As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1

Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4

Public Const WM_MOUSEMOVE = &H200
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDOWN = &H204

Global TrayIcon As NOTIFYICONDATA

Public Sub AddToTray(frm As Form, ToolTip As String, Icon, Optional Update As Boolean = False)

    ' add icon to tray, add Form_Mousemove event as callback function
    On Error Resume Next
    TrayIcon.cbSize = Len(TrayIcon)
    TrayIcon.hwnd = frm.hwnd
    TrayIcon.szTip = ToolTip & vbNullChar
    TrayIcon.hiCon = Icon
    TrayIcon.uID = vbNull
    TrayIcon.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    TrayIcon.uCallbackMessage = WM_MOUSEMOVE
    
    If Update Then
        Shell_NotifyIcon NIM_MODIFY, TrayIcon
    Else
        Shell_NotifyIcon NIM_ADD, TrayIcon
    End If
End Sub

Public Sub RemoveFromTray()
    Shell_NotifyIcon NIM_DELETE, TrayIcon
End Sub

