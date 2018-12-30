VERSION 5.00
Begin VB.UserControl frmSkinV8 
   BackColor       =   &H0000FF00&
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2205
   ControlContainer=   -1  'True
   ScaleHeight     =   2205
   ScaleWidth      =   2205
   Begin prjFrmSkinV8.LunaButton cmdMinimize 
      Height          =   405
      Left            =   990
      TabIndex        =   4
      Top             =   0
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   714
      Picture         =   "frmskinV8.ctx":0000
      MouseOverPicture=   "frmskinV8.ctx":027E
      MouseDownPicture=   "frmskinV8.ctx":04E5
   End
   Begin prjFrmSkinV8.LunaButton cmdUnload 
      Height          =   390
      Left            =   1830
      TabIndex        =   0
      Top             =   0
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   688
      Picture         =   "frmskinV8.ctx":076A
      MouseOverPicture=   "frmskinV8.ctx":09E8
      MouseDownPicture=   "frmskinV8.ctx":0C63
   End
   Begin VB.Label lblDrag 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   0
      TabIndex        =   3
      Top             =   -15
      Width           =   2160
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   30
      TabIndex        =   1
      Top             =   105
      Width           =   885
   End
   Begin VB.Label lblCaptionShadow 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '투명
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   45
      TabIndex        =   2
      Top             =   135
      Width           =   885
   End
   Begin VB.Image Image2 
      Height          =   405
      Left            =   1410
      Picture         =   "frmskinV8.ctx":0EE8
      Top             =   0
      Width           =   300
   End
   Begin VB.Image DownRight 
      Height          =   525
      Left            =   1680
      Picture         =   "frmskinV8.ctx":1084
      Top             =   1680
      Width           =   525
   End
   Begin VB.Image Down 
      Height          =   525
      Left            =   525
      Picture         =   "frmskinV8.ctx":1113
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1200
   End
   Begin VB.Image DownLeft 
      Height          =   525
      Left            =   0
      Picture         =   "frmskinV8.ctx":119F
      Top             =   1680
      Width           =   525
   End
   Begin VB.Image Right 
      Height          =   1200
      Left            =   1680
      Picture         =   "frmskinV8.ctx":122D
      Stretch         =   -1  'True
      Top             =   525
      Width           =   525
   End
   Begin VB.Image Center 
      Height          =   1215
      Left            =   525
      Picture         =   "frmskinV8.ctx":12DA
      Stretch         =   -1  'True
      Top             =   525
      Width           =   1200
   End
   Begin VB.Image Left 
      Height          =   1200
      Left            =   0
      Picture         =   "frmskinV8.ctx":1367
      Stretch         =   -1  'True
      Top             =   525
      Width           =   525
   End
   Begin VB.Image Up 
      Height          =   525
      Left            =   525
      Picture         =   "frmskinV8.ctx":1414
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1200
   End
   Begin VB.Image UpLeft 
      Height          =   525
      Left            =   0
      Picture         =   "frmskinV8.ctx":14E7
      Top             =   0
      Width           =   525
   End
   Begin VB.Image UpRight 
      Height          =   525
      Left            =   1680
      Picture         =   "frmskinV8.ctx":1582
      Top             =   0
      Width           =   525
   End
End
Attribute VB_Name = "frmSkinV8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim CName As String
Dim frm As Form
Option Explicit
'//폼드래그를위해 선언
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
'//폼드래그를위해 선언 끝


Private Sub cmdMinimize_Click()
frm.WindowState = 1
End Sub

Private Sub cmdUnload_Click()
Unload frm
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
      Set frm = Parent
CName = PropBag.ReadProperty("Caption", UserControl.Name)
    lblCaption = CName
    lblCaptionShadow = CName
End Sub

Private Sub UserControl_Resize()
Dim vWidth As Integer
Dim vHeight As Integer
Dim vLeft As Integer
Dim vTop As Integer

'vWidth = Width - (UpLeft.Width + UpRight.Width)
'Up.Width = vWidth: Center.Width = vWidth: Down.Width = vWidth
'vLeft = Width - UpRight.Width
'UpRight.Left = vLeft: Right.Left = vLeft: DownRight.Left = vLeft



vWidth = Width - UpLeft.Width - UpRight.Width
Up.Width = vWidth: Center.Width = vWidth: Down.Width = vWidth

vHeight = Height - UpLeft.Height - DownLeft.Height
Left.Height = vHeight: Center.Height = vHeight: Right.Height = vHeight

vLeft = UpLeft.Width + Up.Width
UpRight.Left = vLeft: Right.Left = vLeft: DownRight.Left = vLeft:

vTop = UpLeft.Height + Left.Height
DownLeft.Top = vTop: Down.Top = vTop: DownRight.Top = vTop:


cmdMinimize.Left = Width - 1215
Image2.Left = Width - 795
cmdUnload.Left = Width - 375
lblCaption.Top = lblCaption.Top + 5
lblCaption.Width = Width - 1320
lblCaptionShadow.Width = Width - 1320
lblDrag.Width = Width
End Sub

Private Sub lblDrag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lngReturnValue As Long

 If Button = 1 Then
  Call ReleaseCapture
  lngReturnValue = SendMessage(frm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Caption", CName, Empty)
End Sub


Public Property Get Caption() As String
Caption = CName
End Property

Public Property Let Caption(Str As String)
CName = Str
    lblCaption = CName
    lblCaptionShadow = CName
End Property



