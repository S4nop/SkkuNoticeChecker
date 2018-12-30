VERSION 5.00
Begin VB.Form frmSkin 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  '없음
   Caption         =   "SKKU Notice Checker v.Beta"
   ClientHeight    =   6900
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   6495
   FillColor       =   &H000080FF&
   Icon            =   "frmSkin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows 기본값
   Begin VB.ListBox lstSoftData 
      Appearance      =   0  '평면
      Height          =   5430
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   6255
   End
   Begin prjFrmSkinV8.frmSkinV8 UserControl11 
      Height          =   6915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   12197
      Caption         =   "SKKU Notice Checker v2.0.0"
      Begin VB.ComboBox cboLst 
         Appearance      =   0  '평면
         Height          =   300
         ItemData        =   "frmSkin.frx":3AFA
         Left            =   120
         List            =   "frmSkin.frx":3B04
         TabIndex        =   6
         Text            =   "::성균관대학교::"
         Top             =   480
         Width           =   6255
      End
      Begin prjFrmSkinV8.jcbutton cmds 
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   6360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         ButtonStyle     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   0
         Caption         =   "Setting"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjFrmSkinV8.jcbutton cmdQ 
         Height          =   375
         Left            =   5280
         TabIndex        =   3
         Top             =   6360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         ButtonStyle     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   0
         Caption         =   "Quit"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjFrmSkinV8.jcbutton cmdr 
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   6360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         ButtonStyle     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   0
         Caption         =   "Refresh"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin VB.Timer TmrRef 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   1680
         Top             =   1320
      End
      Begin VB.ListBox lstSkkuData 
         Appearance      =   0  '평면
         Height          =   5430
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   6255
      End
      Begin VB.Label lbl 
         BackStyle       =   0  '투명
         Caption         =   "Sanop(rshtiger@g.skku.edu)"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   6450
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************
'* Copyright ⓒ 2018 by Sanop(rshtiger@naver.com), All rights reserved.
'*
'* Permission is hereby granted, free of charge, to any person
'* obtaining a copy of this software and associated documentation
'* files (the “Software”), to deal in the Software without
'* restriction, including without limitation the rights to use,
'* copy, modify, merge, publish,distribute, sublicense, and/or sell
'* copies of the Software, and to permit persons to whom the
'* Software is furnished to do so, subject to the following
'* conditions:
'*
'* The above copyright notice and this permission notice shall be
'* included in all copies or substantial portions of the Software.
'*
'* THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND,
'* EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
'* OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
'* NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
'* HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
'* WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
'* FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
'* OTHER DEALINGS IN THE SOFTWARE.
'*
'* Project name : Skku Notice Checker
'*
'* Written by   : Ryu, Sung Hee
'*                Department of Software
'*                Sunkyunkwan University
'*
'* Verseion     : 2.0.0
'*
'* J.Nakim's form skin is used.
'*************************************************************************************************
Const nowver = "2.0.0"
Dim Skku_Link(50) As String
Dim Soft_Link(50) As String
Dim Shadow As clsShadow
Dim W As New WinHttp.WinHttpRequest
Public svSkku As String, chkSkku As String
Public svSoft As String, chkSoft As String
Public tmplink As String, reftim As Integer
Dim locklist As Boolean
Dim firstrun As Boolean
Dim lastchk As Boolean
Dim lastnum As String
Public Function getAllNotice()
Dim fn As Integer, fn2 As Integer
fn = GetSkkuNotice()
fn2 = GetSoftNotice()

If firstrun = False And fn = 0 Then Exit Function
lastchk = True
If Me.WindowState = vbMinimized Then
    lastchk = False
Else
    svSkku = chkSkku
    svSoft = chkSoft
End If
firstrun = False

If Me.Visible = True Then
    MsgBox "새로운 공지사항" & vbCrLf & "성균관대학교 : " & fn & vbCrLf & "소프트웨어학과 : " & fn2
Else
    Call frmTray.TrayChk(fn, fn2)
End If

End Function

Public Function GetSkkuNotice() As Integer
Dim i As Integer, n As Integer, fn As Integer
On Error GoTo err
chk = False
lstSkkuData.Clear

Dim chked As Boolean
Dim strtmp As String, lnum As String

chked = False
For j = 1 To 5
    W.Open "GET", "https://www.skku.edu/skku/campus/skk_comm/notice01.do?mode=list&&articleLimit=10&article.offset=" & j & "0"
    W.Send
    W.WaitForResponse
    
    Dim fullD As String, listd() As String
    fullD = W.ResponseText
    Clipboard.SetText fullD
    
    listd = Split(fullD, "<td class=""left"">")
    
    For i = 1 To UBound(listd) - 1 'Step -1
        lnum = Split(Split(listd(i), "<a href=""?mode=view&amp;articleNo=")(1), "&amp;article.offset=" & j & "0&amp;articleLimit=10"" class="""">")(0)
        If i = 1 And j = 1 Then chkSkku = lnum
        If lnum = svSkku Then
        fn = n
        chked = True
        End If
        lstSkkuData.AddItem "-----------------------------------------------------------------------------------------------------------------------------------------"
        strtmp = Split(Split(listd(i), "class="""">")(1), "</a>")(0)
        strtmp = Trim(Mid(strtmp, 126, Len(strtmp) - 60))
        If chked = False Then
        lstSkkuData.AddItem Split(Split(listd(i), "<td>")(1), "</td>")(0) & " || " & "<New Notice> " & strtmp
        Else
        lstSkkuData.AddItem Split(Split(listd(i), "<td>")(1), "</td>")(0) & " || " & strtmp
        End If
        Skku_Link(n) = Split(Split(listd(i), "<a href=""?mode=view&amp;articleNo=")(1), "&amp;article.offset=" & j & "0&amp;articleLimit=10"" class="""">")(0)
        n = n + 1
    Next i
Next j

If svSkku = "" Then fn = n
GetSkkuNotice = fn
Exit Function
err:
lstSoftData.AddItem "데이터를 받아오는 데 실패하였습니다."
End Function

Public Function GetSoftNotice() As Integer
Dim AllData As String, splitedD() As String, isNotice As Boolean, chked As Boolean, i As Integer, j As Integer, fn As Integer
Dim json() As String, id As String, title As String, category As String, time As String
On Error GoTo err
W.Open "POST", "http://cs.skku.edu/ajax/board/list/notice"
W.SetRequestHeader "Accept-Language", "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7"
W.SetRequestHeader "Host", "cs.skku.edu"
W.SetRequestHeader "Referer", "http://cs.skku.edu/news/notice/list"
W.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.220 Whale/1.3.53.4 Safari/537.36"
W.Send
W.WaitForResponse

AllData = W.ResponseText
splitedD = Split(AllData, "{")

isNotice = True
i = 1
Do While (isNotice)
    i = i + 1
    If InStr(splitedD(i), "공지") = 0 Then isNotice = False
    'If MsgBox(InStr(splitedD(i), "공지") & vbCrLf & splitedD(i), vbYesNo) = vbYes Then End
Loop
chked = False
For j = i To i + 45
    json = Split(splitedD(j), ",")
    id = Split(json(0), ":")(1)
    title = Split(Split(json(1), ":")(1), """")(1)
    category = Split(Split(json(2), ":")(1), """")(1)
    time = Split(Split(json(3), ":")(1), """")(1)
    Soft_Link(j - i) = id
    If id = svSoft Then
        chked = True
        fn = j - i
    End If
    If j = i Then chkSoft = id
    lstSoftData.AddItem "-----------------------------------------------------------------------------------------------------------------------------------------"
    If chked Then
        lstSoftData.AddItem time & " " & category & " || " & title
    Else
        lstSoftData.AddItem time & " " & category & " || <New notice> " & title
    End If
Next j

If svSoft = "" Then fn = 45
GetSoftNotice = fn
Exit Function
err:
lstSoftData.AddItem "데이터를 받아오는 데 실패하였습니다."
End Function
Public Function GetSkkuHTML(n As Integer)
On Error GoTo err
'MsgBox Skku_Link(n)
W.Open "GET", "https://www.skku.edu/skku/campus/skk_comm/notice01.do?mode=view&articleNo=" & Skku_Link(n)
W.Send
W.WaitForResponse
tmplink = Split(Split(Split(W.ResponseText, "<dl class=""board-write-box board-write-box-v03"">")(1), "</dl>")(0), "<dt class=""hide replyNone"">게시글 내용</dt>")(1)
If InStr(W.ResponseText, "javascript:downLoad") Then tmplink = "<a href=https://www.skku.edu/skku/campus/skk_comm/notice01.do?mode=view&articleNo=" & Skku_Link(n) & ">※※첨부파일이 있습니다. 웹에서 확인해 주세요※※</a><br/>----------------------------------------------------------------<br/>" & tmplink
frmBrswer.Show vbModal, Me
Exit Function
err:
MsgBox "데이터를 받아오는 중 에러가 발생했습니다. 다시 시도해 주세요"
End Function

Public Function GetSoftHTML(n As Integer)
On Error GoTo err
W.Open "POST", "http://cs.skku.edu/ajax/board/view/notice/" & Soft_Link(n)
W.SetRequestHeader "Accept-Language", "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7"
W.SetRequestHeader "Host", "cs.skku.edu"
W.SetRequestHeader "Referer", "http://cs.skku.edu/news/notice/view/" & Soft_Link(n)
W.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.220 Whale/1.3.53.4 Safari/537.36"
W.Send
W.WaitForResponse

tmplink = Split(Split(W.ResponseText, """text"":""")(1), """,""fixDate")(0)
tmplink = Replace(tmplink, "\n", vbCrLf)
If InStr(Split(W.ResponseText, "fixDate")(1), "files"":[]") = 0 Then tmplink = "<a href=http://cs.skku.edu/ajax/board/view/notice/" & Soft_Link(n) & ">※※첨부파일이 있습니다. 웹에서 확인해 주세요※※</a><br/>----------------------------------------------------------------<br/>" & tmplink
frmBrswer.Show vbModal, Me
Exit Function
err:
MsgBox "데이터를 받아오는 중 에러가 발생했습니다. 다시 시도해 주세요"
End Function

Private Sub cboLst_Click()
If cboLst.Text = "::성균관대학교::" Then
    lstSkkuData.Visible = True
    lstSoftData.Visible = False
Else
    lstSkkuData.Visible = False
    lstSoftData.Visible = True
    
End If
End Sub

Private Sub cmdQ_Click()
Unload Me
End Sub

Private Sub cmdr_Click()
getAllNotice
End Sub

Private Sub cmds_Click()
FrmSetting.Show vbModal, Me
End Sub

Private Sub Command1_Click()
MsgBox reftim
End Sub

Private Sub Form_Initialize()
FadeIN Me
Load FrmSetting

If FrmSetting.chk(1) = 1 Then
Me.Hide
frmTray.GoTray
End If
getAllNotice
End Sub

Private Sub form_load()
'//그림자소스
On Error Resume Next
ff = FreeFile
If Dir(App.Path & "\tmp.dat") = vbNullString Then
s = 0
Else
Open App.Path & "\tmp.dat" For Input As ff
Line Input #ff, svSkku
Line Input #ff, svSoft
Close ff
End If
    Set Shadow = New clsShadow
    Call Shadow.Shadow(Me)
    Shadow.Color = vbBlack
    Shadow.Depth = 4
'//그림자소스끝
'//앰피제로님의 아이콘 적용소스
'ApplyIcon hWnd

'W.Open "GET", ""
'W.Send
'W.WaitForResponse
firstrun = True
'Dim upinfo As String, uptmp As String
'uptmp = Replace(Replace(Replace(Split(Split(W.ResponseText, "<div class=""tt_article_useless_p_margin""><p>")(1), "</p><div")(0), "&gt;", ">"), "&quot", """"), "&amp;", "&")
'Clipboard.Clear
'Clipboard.SetText uptmp
'upinfo = cDecode(uptmp, "skku@ncsft17")
'If Split(Split(upinfo, vbCrLf)(0), " : ")(1) <> nowver Then
'If MsgBox("새로운 버전이 발견되었습니다. 새로운 버전을 다운로드 받으시겠습니까?", vbYesNo) = vbYes Then Shell "explorer " & Split(Split(upinfo, vbCrLf)(1), " : ")(1)
'End If
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then
Me.Hide
frmTray.GoTray
Else
lastchk = True
If firstrun = False Then s = chknow
End If
End Sub

Public Sub form_unload(cancel As Integer)
FadeOUT Me
ff = FreeFile
Open App.Path & "\tmp.dat" For Output As ff
Print #ff, chkSkku & vbCrLf & chkSoft
Close ff
Unload frmTray
Unload FrmSetting
'Cancel = 1
End Sub

Private Sub jcbutton1_Click()
    MsgBox "JN frmSkin V8"
End Sub

Private Sub lstSkkuData_DblClick()
If locklist = True Then Exit Sub
If lstSkkuData.List(lstSkkuData.ListIndex) = "-------------------------------------------------------------------------------------------------------------------------------" Then Exit Sub
locklist = True
GetSkkuHTML ((lstSkkuData.ListIndex - 1) / 2)
locklist = False
End Sub

Private Sub lstSoftData_DblClick()
If locklist = True Then Exit Sub
If lstSoftData.List(lstSoftData.ListIndex) = "-------------------------------------------------------------------------------------------------------------------------------" Then Exit Sub
locklist = True
GetSoftHTML ((lstSoftData.ListIndex - 1) / 2)
locklist = False
End Sub

Private Sub TmrRef_Timer()
Static i As Integer
i = i + 1
If i = reftim Then
getAllNotice
i = 0
End If

End Sub
