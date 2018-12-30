Attribute VB_Name = "modTray"
'���� : modTray.bas
'
'2008/07/19 ����

Option Explicit
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeout As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type

Public Const NOTIFYICON_VERSION = 3       'V5 style taskbar
Public Const NOTIFYICON_OLDVERSION = 0    'Win95 style taskbar

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIM_SETFOCUS = &H3
Public Const NIM_SETVERSION = &H4

Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIF_STATE = &H8
Public Const NIF_INFO = &H10

Public Const NIS_HIDDEN = &H1
Public Const NIS_SHAREDICON = &H2

Public Const NIIF_NONE = &H0
Public Const NIIF_WARNING = &H2
Public Const NIIF_ERROR = &H3
Public Const NIIF_INFO = &H1
Public Const NIIF_GUID = &H4

Public myData As NOTIFYICONDATA


Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Public Const TRAY_CALLBACK = (&H400 + 1001&)
Public Const GWL_WNDPROC = -4

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206

Public Enum WIN_STATUS
    STA_MIN
    STA_NORMAL
End Enum

Public glWinRet As Long
Public OrgWinRet As Long
Public Status As WIN_STATUS
Public MyForm As Form
Public Function CallbackMsgs(ByVal wHwnd As Long, ByVal wMsg As Long, ByVal wp_id As Long, ByVal lp_id As Long) As Long
On Error Resume Next
    If wMsg = TRAY_CALLBACK Then
        With MyForm
            Select Case CLng(lp_id)
                Case WM_RBUTTONUP
                    .PopupMenu .mnuTray, , , , .mnushow
                Case WM_LBUTTONUP
                    If frmSkin.Visible = False Then
                    frmSkin.WindowState = vbNormal
                    frmSkin.Show
                    UnloadTray
                    End If
                    'If .WindowState = vbMinimized Then
                    '  .Visible = True
                    '    .WindowState = vbNormal
                    'Else
                    '    .WindowState = vbMinimized
                    '    .Visible = False
                    'End If
            End Select
        End With
    End If
    CallbackMsgs = CallWindowProc(glWinRet, wHwnd, wMsg, wp_id, lp_id)
End Function

Public Function ShowTip(ByVal TipTitle As String, ByVal TipContent As String, TipIco As Integer)
With myData
    .szInfoTitle = TipTitle & vbNullChar
    .szInfo = TipContent & vbNullChar
    .dwInfoFlags = TipIco
    
End With
Shell_NotifyIcon NIM_MODIFY, myData
End Function
Public Function CreatTray(ByRef TheForm As Form, TipMove As String, TipTitle As String, TipContent As String, TipIco As Long)
Set MyForm = TheForm
OrgWinRet = GetWindowLong(MyForm.hWnd, GWL_WNDPROC)
With myData
    .cbSize = Len(myData)
    .hWnd = MyForm.hWnd
    .uID = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_INFO Or NIF_MESSAGE
    .uCallbackMessage = TRAY_CALLBACK
    .hIcon = MyForm.Icon
    .szTip = TipMove & vbNullChar
    .dwState = 0
    .dwStateMask = 0
    .szInfoTitle = TipTitle & vbNullChar
    .szInfo = TipContent & vbNullChar
    .dwInfoFlags = TipIco
    .uTimeout = 10000
End With
Shell_NotifyIcon NIM_ADD, myData
glWinRet = SetWindowLong(MyForm.hWnd, GWL_WNDPROC, AddressOf CallbackMsgs)
End Function
Public Function UnloadTray()
Shell_NotifyIcon NIM_DELETE, myData
Call SetWindowLong(MyForm.hWnd, GWL_WNDPROC, OrgWinRet)
End Function
