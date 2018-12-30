VERSION 5.00
Begin VB.Form frmTray 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   4680
   Icon            =   "frmTray.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.PictureBox P 
      Height          =   375
      Left            =   2400
      Picture         =   "frmTray.frx":3AFA
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   480
      Width           =   375
   End
   Begin VB.Menu mnuTray 
      Caption         =   "trayMenu"
      Begin VB.Menu mnushow 
         Caption         =   "열기(&Open)"
      End
      Begin VB.Menu mnuref 
         Caption         =   "새로 고침(&Refresh)"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "정보(&About)"
      End
      Begin VB.Menu mnuquit 
         Caption         =   "종료(&Quit)"
      End
   End
End
Attribute VB_Name = "frmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type NOTIFYICONDATA
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
  uTimeoutOrVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
End Type
  

Private Const NIIF_WARNING = 2
Private Const NIIF_ERROR = 3
Private Const NIIF_INFO = 1

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Dim SysTrayT As NOTIFYICONDATA
Public Function GoTray()
Call CreatTray(frmTray, "SKKU Notice Checker", "SKKU Notice Checker", "트레이로 이동하였습니다", 1)
        'Shell_NotifyIcon &H0, SysTrayT
End Function
Private Sub Form_Resize()
With Me
    If Status = STA_NORMAL And .WindowState = vbMinimized And .Visible = True Then
        .Visible = False
        Status = STA_MIN
        Else
        Status = STA_NORMAL
    End If
End With
End Sub
Private Sub mabout_Click()
End Sub
Public Function TrayChk(fn As Integer, fn2 As Integer)

Call ShowTip("SKKU Notice Checker", "새로운 공지사항" & vbCrLf & "성균관대학교 : " & fn & vbCrLf & "소프트웨어학과 : " & fn2, 1)

End Function
Private Sub mopen_Click()
End Sub

Private Sub mquit_Click()
End
End Sub

Private Sub mnuabout_Click()
MsgBox "SKKU Notice Checker" & vbCrLf & "Made By Sanop(류성희)" & vbCrLf & "COPYRIGHT ⓒ 2018 Sanop(RYU SUNG HEE). ALL RIGHTS RESERVED"

End Sub

Private Sub mnuquit_Click()
Unload frmSkin
End Sub

Private Sub mnuref_Click()
frmSkin.getAllNotice
End Sub

Private Sub mnushow_Click()
frmSkin.WindowState = vbNormal
frmSkin.Show
UnloadTray
End Sub
