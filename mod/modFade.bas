Attribute VB_Name = "modFade"
Option Explicit
Public Const LWA_ALPHA As Long = &H2
Public Const GCL_STYLE As Long = -26&
Public Const GWL_STYLE As Long = -16&
Public Const GWL_EXSTYLE As Long = -20
Public Const ws_ex_layered As Long = &H80000
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Function FadeIN(frm As Form)
   Const timeFadeIn As Long = 6 '점점 불투명해지는 간격(밀리초)
   Dim lastTimer As Single, i As Long

 

   '페이드 인(Fade In) 효과
   frm.Enabled = False
   Call SetWindowLong(frm.hWnd, GWL_EXSTYLE, GetWindowLong(frm.hWnd, GWL_EXSTYLE) Or ws_ex_layered)
   Call SetLayeredWindowAttributes(frm.hWnd, 0, 0, LWA_ALPHA)
   frm.Show
   For i = 1 To 255 Step 7
      lastTimer = Timer
      Do While Timer < lastTimer + (timeFadeIn / 1000)
         DoEvents
      Loop
      Call SetLayeredWindowAttributes(frm.hWnd, 0, i, LWA_ALPHA)
      DoEvents
   Next
   Call SetWindowLong(frm.hWnd, GWL_EXSTYLE, GetWindowLong(frm.hWnd, GWL_EXSTYLE) Xor ws_ex_layered)
   frm.Enabled = True
End Function


Function FadeOUT(frm As Form)
'For i = 1 To 255 Step 7

   Const timeFadeIn As Long = 6 '점점 불투명해지는 간격(밀리초)
   Dim lastTimer As Single, i As Long

 

   '페이드 아웃(Fade Out) 효과
   'frm.Enabled = False
   Call SetWindowLong(frm.hWnd, GWL_EXSTYLE, GetWindowLong(frm.hWnd, GWL_EXSTYLE) Or ws_ex_layered)
   Call SetLayeredWindowAttributes(frm.hWnd, 0, 0, LWA_ALPHA)
   'frm.Show
   For i = 255 To 0 Step -7
      lastTimer = Timer
      Do While Timer < lastTimer + (timeFadeIn / 1000)
         DoEvents
      Loop
      Call SetLayeredWindowAttributes(frm.hWnd, 0, i, LWA_ALPHA)
      DoEvents
   Next
   Call SetWindowLong(frm.hWnd, GWL_EXSTYLE, GetWindowLong(frm.hWnd, GWL_EXSTYLE) Xor ws_ex_layered)
   frm.Enabled = True
   
   
End Function

