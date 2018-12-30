VERSION 5.00
Begin VB.Form FrmSetting 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  '고정 도구 창
   Caption         =   " Setting"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.ListBox lstfilt 
      Height          =   1320
      ItemData        =   "FrmSetting.frx":0000
      Left            =   360
      List            =   "FrmSetting.frx":0002
      TabIndex        =   6
      Top             =   2400
      Width           =   5535
   End
   Begin VB.CheckBox chk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "알림 제외 키워드 설정(준비중)"
      Enabled         =   0   'False
      Height          =   615
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   4575
   End
   Begin VB.TextBox TxtTime 
      Height          =   270
      Left            =   1680
      TabIndex        =   4
      Text            =   "10"
      Top             =   1500
      Width           =   615
   End
   Begin VB.CheckBox chk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "자동 새로 고침                    분 간격"
      Height          =   615
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   4575
   End
   Begin VB.CheckBox chk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "실행 시 트레이로 자동 이동"
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4575
   End
   Begin VB.CheckBox chk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "시작 프로그램에 등록"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "  COPYRIGHT ⓒ 2017 UNSEED(RYU SUNG HEE). ALL RIGHTS RESERVED"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   4080
      Width           =   6495
   End
End
Attribute VB_Name = "FrmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAs_Click()

End Sub

Private Sub chkTray_Click()


End Sub

Private Sub chk_Click(Index As Integer)
If chk(2).Value = 1 Then
frmSkin.reftim = TxtTime
frmSkin.TmrRef.Enabled = True
Else
frmSkin.TmrRef.Enabled = False
End If
End Sub

Private Sub form_load()
Dim ff
ff = FreeFile
Dim sset(2) As String
If Dir(App.Path & "\setting.ini") = vbNullString Then
Open App.Path & "\setting.ini" For Output As ff
Print #ff, "AutoStart : 0" & vbCrLf & "AutoToTray : 0" & vbCrLf & "AutoRefresh : 10"
Close ff
End If

Open App.Path & "\setting.ini" For Input As ff
Do Until EOF(ff)
Line Input #ff, sset(i)
If i = 2 Then
If Split(sset(i), " : ")(1) <> 0 Then
TxtTime.Text = Split(sset(i), " : ")(1)
frmSkin.reftim = TxtTime.Text
chk(i).Value = 1
Else
chk(i).Value = 0
frmSkin.TmrRef.Enabled = False
End If
Else
FrmSetting.chk(i) = Split(sset(i), " : ")(1)
i = i + 1
End If
Loop
Close ff
End Sub

Private Sub form_unload(cancel As Integer)
ff = FreeFile
If TxtTime.Text = 0 Or TxtTime.Text = vbNullString Then TxtTime.Text = 10
If chk(2).Value = 1 Then frmSkin.reftim = TxtTime.Text
Dim Result As Integer
    If chk(0).Value = 1 Then
Startup
    
Else
StartUPDelete
    
End If


Open App.Path & "\setting.ini" For Output As ff

Print #ff, "AutoStart : " & chk(0).Value & vbCrLf & "AutoToTray : " & chk(1).Value & vbCrLf & "AutoRefresh : " & chk(2).Value * CInt(TxtTime.Text) ' & vbCrLf & "Filtering Notices : " & vbCrLf
Close ff
End Sub

