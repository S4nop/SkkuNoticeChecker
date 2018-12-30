VERSION 5.00
Begin VB.UserControl LunaButton 
   Appearance      =   0  '���
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   870
   ForeColor       =   &H00000000&
   LockControls    =   -1  'True
   PaletteMode     =   4  '����
   ScaleHeight     =   77
   ScaleMode       =   3  '�ȼ�
   ScaleWidth      =   58
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   390
      Top             =   60
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   0
      Left            =   60
      Top             =   60
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   1
      Left            =   60
      Top             =   330
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   2
      Left            =   60
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   3
      Left            =   60
      Top             =   870
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "LunaButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Programmed by Chun Dong Hyuk

Option Explicit

'Mouse Over �̺�Ʈ�� ���� API�� �����Ѵ�.
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

'API ���� ���� ����ü
Private Type PointAPI
    x As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'����� ���� ��Ʈ�ѿ��� ����� �̺�Ʈ�� �����Ѵ�.
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseOver()
Public Event MouseExit()

'��ư MouseOver ������ ���� �÷��� ����
Private ButtonMouseOverFlag As Integer

'��Ʈ���� �ڵ��� ��ȯ�Ѵ�.
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd

End Property

'��Ʈ���� �⺻ �̹����� ��ȯ�Ѵ�.
Public Property Get Picture() As Picture
    Set Picture = UserControl.Image1(0).Picture
    
End Property

'��Ʈ���� �⺻ �̹����� �����Ѵ�.
Public Property Set Picture(ByVal NewPicture As Picture)
    Set UserControl.Image1(0).Picture = NewPicture
    PropertyChanged "Picture"
    UserControl.Picture = Image1(0).Picture
    
End Property

'��Ʈ���� MouseOver �̹����� ��ȯ�Ѵ�.
Public Property Get MouseOverPicture() As Picture
    Set MouseOverPicture = UserControl.Image1(1).Picture

End Property

'��Ʈ���� MouseOver �̹����� �����Ѵ�.
Public Property Set MouseOverPicture(ByVal NewPicture As Picture)
    Set UserControl.Image1(1).Picture = NewPicture
    PropertyChanged "MouseOverPicture"
    
End Property

'��Ʈ���� Disable �̹����� ��ȯ�Ѵ�.
Public Property Get DisablePicture() As Picture
    Set DisablePicture = UserControl.Image1(3).Picture

End Property

'��Ʈ���� Disable �̹����� �����Ѵ�.
Public Property Set DisablePicture(ByVal NewPicture As Picture)
    Set UserControl.Image1(3).Picture = NewPicture
    PropertyChanged "DisablePicture"
    
End Property

'��Ʈ���� MouseDown �̹����� ��ȯ�Ѵ�.
Public Property Get MouseDownPicture() As Picture
    Set MouseDownPicture = UserControl.Image1(2).Picture

End Property

'��Ʈ���� MouseDown �̹����� �����Ѵ�.
Public Property Set MouseDownPicture(ByVal NewPicture As Picture)
    Set UserControl.Image1(2).Picture = NewPicture
    PropertyChanged "MouseDownPicture"
    
End Property

'��Ʈ���� Enabled �Ӽ��� ��ȯ�Ѵ�.
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled

End Property

'��Ʈ���� Enabled �Ӽ��� �����Ѵ�.
Public Property Let Enabled(ByVal Flag As Boolean)
    UserControl.Enabled = Flag
    If Flag = True Then
      UserControl.Picture = Image1(0).Picture
    Else
      UserControl.Picture = Image1(3).Picture
    End If
    PropertyChanged "Enabled"
    
End Property

'MouseOver �̺�Ʈ ó���� ���� Ÿ�̸Ӹ� ����Ѵ�.
Private Sub Timer1_Timer()
    On Error Resume Next
    Dim WindowPointAPI As PointAPI
    Dim WindowRect As RECT
    GetCursorPos WindowPointAPI
    GetWindowRect UserControl.hWnd, WindowRect
    If (WindowRect.Left <= WindowPointAPI.x And WindowRect.Right >= WindowPointAPI.x And WindowRect.Top <= WindowPointAPI.Y And WindowRect.Bottom >= WindowPointAPI.Y) Then
      If ButtonMouseOverFlag = 1 Then Exit Sub
      UserControl.Picture = Image1(1).Picture
      ButtonMouseOverFlag = 1
      RaiseEvent MouseOver
    Else
      UserControl.Picture = Image1(0).Picture
      ButtonMouseOverFlag = 0
      Timer1.Enabled = False
      RaiseEvent MouseExit
    End If

End Sub

'��Ʈ�� Ŭ���� Ŭ���̺�Ʈ�� �߻��Ѵ�.
Private Sub UserControl_Click()
    RaiseEvent Click

End Sub

'�ʱ� ��Ʈ�� ũ�⸦ �����Ѵ�.
Private Sub UserControl_InitProperties()
    UserControl.Height = 255
    UserControl.Width = 255

End Sub

'��Ʈ�� MouseDown �̺�Ʈ�� ó���Ѵ�.
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Timer1.Enabled = False
    UserControl.Picture = Image1(2).Picture
    RaiseEvent MouseDown(Button, Shift, x, Y)

End Sub

'��Ʈ�� MouseMove �̺�Ʈ�� ó���Ѵ�.
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If ButtonMouseOverFlag = 0 Then
      Timer1.Enabled = True
    End If

End Sub

'��Ʈ�� MouseUp �̺�Ʈ�� ó���Ѵ�.
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    UserControl.Picture = Image1(1).Picture
    Timer1.Enabled = True
    RaiseEvent MouseUp(Button, Shift, x, Y)

End Sub

'��Ʈ�� �Ӽ��� �о �����Ѵ�.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    Set DisablePicture = PropBag.ReadProperty("DisablePicture", Nothing)
    Set MouseOverPicture = PropBag.ReadProperty("MouseOverPicture", Nothing)
    Set MouseDownPicture = PropBag.ReadProperty("MouseDownPicture", Nothing)
    Enabled = PropBag.ReadProperty("Enabled", True)
    
End Sub

'��Ʈ�� �Ӽ��� �����Ѵ�.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Picture", Picture, Nothing
    PropBag.WriteProperty "DisablePicture", DisablePicture, Nothing
    PropBag.WriteProperty "MouseOverPicture", MouseOverPicture, Nothing
    PropBag.WriteProperty "MouseDownPicture", MouseDownPicture, Nothing
    PropBag.WriteProperty "Enabled", Enabled, True
    
End Sub


