VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************************************
'* clsShadow - SelfSub based sample that creates a form shadow.
'*
'* Originally by Paul Caton
'* SelfSub contributions made by LaVolpe and Tai Chi Minh Ralph Eastwood
'* Minor edits and improvements by Joe Jordan (March 6, 2009)
'*
'*
'* Copyright free, use and abuse as you see fit.
'*
'* v1.0 Re-write of the SelfSub/WinSubHook-2 submission to Planet Source Code............ 20060322
'* v1.1 Shadow color property added...................................................... 20060406
'* v1.2 SelfSub updated to LaVolpe's thunks, removed unused code
'*      Added Windows Vista, Windows 7 Support........................................... 20090306
'*************************************************************************************************

Option Explicit

'------------------------------------------------------------------------------------------------------------------------------------
' SelfSub Items used in all thunks (SelfHook and SelfCallback have been removed, see URL above for full version)
'------------------------------------------------------------------------------------------------------------------------------------
  ' Local variables/constants
    Private z_scFunk            As Collection   'hWnd/thunk-address collection; initialized as needed
    Private Const IDX_INDEX     As Long = 2     'index of the subclassed hWnd OR hook type
    Private Const IDX_PREVPROC  As Long = 9     'Thunk data index of the original WndProc
    Private Const IDX_BTABLE    As Long = 11    'Thunk data index of the Before table for messages
    Private Const IDX_ATABLE    As Long = 12    'Thunk data index of the After table for messages
    Private Const IDX_CALLBACKORDINAL As Long = 36 ' Ubound(callback thunkdata)+1, index of the callback

  ' Declarations:
    Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
    Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
    Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
    Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
    Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
    Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    
    Private Enum eThunkType
        SubclassThunk = 0
    End Enum

    '-Selfsub specific declarations----------------------------------------------------------------------------
    Private Enum eMsgWhen                                                   'When to callback
      MSG_BEFORE = 1                                                        'Callback before the original WndProc
      MSG_AFTER = 2                                                         'Callback after the original WndProc
      MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                            'Callback before and after the original WndProc
    End Enum
    
    ' see ssc_Subclass for complete listing of indexes and what they relate to
    Private Const IDX_PARM_USER As Long = 13    'Thunk data index of the User-defined callback parameter data index
    Private Const IDX_UNICODE   As Long = 107   'Must be UBound(subclass thunkdata)+1; index for unicode support
    Private Const MSG_ENTRIES   As Long = 32    'Number of msg table entries. Set to 1 if using ALL_MESSAGES for all subclassed windows
    Private Const ALL_MESSAGES  As Long = -1    'All messages will callback
    
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
    Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hWnd As Long) As Long
    Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'------------------------------------------------------------------------------------------------------------------------------------


Private Const DEF_DEPTH             As Long = 7
Private Const DEF_TRANSPARENCY      As Long = 140

Private Const AC_SRC_OVER           As Long = &H0
Private Const AC_SRC_ALPHA          As Long = &H1
Private Const BITSPIXEL             As Long = 12
Private Const SIZE_RESTORED         As Long = 0
Private Const SW_HIDE               As Long = 0
Private Const SW_SHOWNOACTIVATE     As Long = 4
Private Const SWP_HIDEWINDOW        As Long = &H80
Private Const SWP_SHOWWINDOW        As Long = &H40
Private Const ULW_ALPHA             As Long = &H2

Private Const WM_DISPLAYCHANGE      As Long = &H7E
Private Const WM_THEMECHANGED       As Long = &H31A
Private Const WM_WINDOWPOSCHANGED   As Long = &H47
Private Const WM_SIZE               As Long = &H5
Private Const ws_ex_layered         As Long = &H80000
Private Const WS_EX_NOPARENTNOTIFY  As Long = &H4
Private Const WS_EX_TRANSPARENT     As Long = &H20
Private Const WS_POPUP              As Long = &H80000000

Private Type tBGRA
    Blue                            As Byte
    Green                           As Byte
    Red                             As Byte
    Alpha                           As Byte
End Type

Private Type tBITMAPINFOHEADER
  biSize                            As Long
  biWidth                           As Long
  biHeight                          As Long
  biPlanes                          As Integer
  biBitCount                        As Integer
  biCompression                     As Long
  biSizeImage                       As Long
  biXPelsPerMeter                   As Long
  biYPelsPerMeter                   As Long
  biClrUsed                         As Long
  biClrImportant                    As Long
End Type

Private Type tBLENDFUNCTION
  BlendOp                           As Byte
  BlendFlags                        As Byte
  SourceConstantAlpha               As Byte
  AlphaFormat                       As Byte
End Type

Private Type tOSVERSIONINFO
  dwOSVersionInfoSize               As Long
  dwMajorVersion                    As Long
  dwMinorVersion                    As Long
  dwBuildNumber                     As Long
  dwPlatformId                      As Long
  szCSDVersion                      As String * 128
End Type

Private Type tPOINT
  x                                 As Long
  Y                                 As Long
End Type

Private Type tSAFEARRAYBOUND
  cElements                         As Long
  lLbound                           As Long
End Type

Private Type tSAFEARRAY2D
  cDims                             As Integer
  fFeatures                         As Integer
  cbElements                        As Long
  cLocks                            As Long
  pvData                            As Long
  Bounds(0 To 1)                    As tSAFEARRAYBOUND
End Type

Private Type tSIZE
  cX                                As Long
  cY                                As Long
End Type

Private Type tWINDOWPOS
  hWnd                              As Long
  hWndInsertAfter                   As Long
  x                                 As Long
  Y                                 As Long
  cX                                As Long
  cY                                As Long
  Flags                             As Long
End Type

Private m_Color                     As Long                                 'Private shadow color property value
Private m_Depth                     As Long                                 'Private shadow depth property value
Private m_Transparency              As Long                                 'Private shadow transparency property value

Private bIsLayered                  As Boolean                              'Layered windows supported
Private bIsLuna                     As Boolean                              'Luna theme?
Private bIsXP                       As Boolean                              'Windows XP (or greater)?
Private bLastShow                   As Boolean                              'The previous show state
Private cX                          As Long                                 'Width
Private cY                          As Long                                 'Height
Private hWndBt                      As Long                                 'Bottom shadow window handle
Private hWndRt                      As Long                                 'Right shadow window handle
Private hWndForm                    As Long                                 'Parent window handle
Private wp                          As tWINDOWPOS                           'Parent window position

Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As tBITMAPINFOHEADER, ByVal un As Long, ByRef lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateWindowExA Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As Long, ByVal cchMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetThemeDocumentationProperty Lib "uxtheme.dll" (ByVal pszThemeName As Long, ByVal pszPropertyName As Long, ByVal pszValueBuff As Long, ByVal cchMaxValChars As Long) As Long
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As tOSVERSIONINFO) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal hdcDest As Long, ptDst As Any, pSize As Any, ByVal hDCSrc As Long, ptSrc As Any, ByVal crKey As Long, pBlend As Any, ByVal dwFlags As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Sub Class_Initialize()                                              'Class initialization
  CheckEnvironment
  m_Depth = DEF_DEPTH
  m_Transparency = DEF_TRANSPARENCY
End Sub

Private Sub Class_Terminate()
  ssc_Terminate                                                              'Terminate all subclassing
  
  If hWndRt <> 0 Then
    DestroyWindow hWndRt
    DestroyWindow hWndBt
  End If
End Sub

'Shadow depth property
Public Property Get Color() As Long
  Color = m_Color
End Property
'
Public Property Let Color(ByVal newValue As Long)
  Const COLOR_SYS_MASK As Long = &H80000000
  
  If (newValue And COLOR_SYS_MASK) Then                                       'If the system color bit is set
    newValue = GetSysColor(newValue Xor COLOR_SYS_MASK)                       'Get the translated system color
  End If
  
  If newValue <> m_Color Then
    m_Color = newValue
    DisplayShadows
  End If
End Property

'Shadow depth property
Public Property Get Depth() As Long
  Depth = m_Depth
End Property
'
Public Property Let Depth(ByVal newValue As Long)
  If newValue <> m_Depth Then
    m_Depth = newValue
    DisplayShadows
  End If
End Property

'Shadow transparency property
Public Property Get Transparency() As Byte
  Transparency = CByte(m_Transparency)
End Property
'
Public Property Let Transparency(ByVal newValue As Byte)
  If newValue <> m_Transparency Then
    m_Transparency = CLng(newValue)
    DisplayShadows
  End If
End Property

'Set the form to be shadowed
Public Function Shadow(frm As Form) As Boolean
  If hWndForm <> 0 Then
    zError "Shadow", "Only a single form per cShadow instance is allowed"
    Exit Function
  End If

  If bIsLayered Then
    hWndForm = frm.hWnd
  
    If ssc_Subclass(hWndForm) Then
      ssc_AddMsg hWndForm, MSG_BEFORE, WM_WINDOWPOSCHANGED
      ssc_AddMsg hWndForm, MSG_BEFORE, WM_SIZE
      ssc_AddMsg hWndForm, MSG_BEFORE, WM_DISPLAYCHANGE
      
      If bIsXP Then
        ssc_AddMsg hWndForm, MSG_BEFORE, WM_THEMECHANGED
      End If
      
      CreateWindows
      Shadow = True
    End If
  End If
End Function

'Validate the OS and color depth
Private Sub CheckEnvironment()
  Dim OSV As tOSVERSIONINFO

  bIsLayered = False

  With OSV
    .dwOSVersionInfoSize = Len(OSV)                                         'Set the length element
    GetVersionExA OSV                                                       'Fill the type with OS version info

    If .dwPlatformId = 2 Then                                               'If it's an NT based OS
      If .dwMajorVersion >= 5 Then                                           'If the major version is 5 or greater, OS supports transparency
        If .dwMinorVersion > 0 Then
          bIsXP = True
          bIsLuna = IsLuna                                                  'Determine if the Luna theme is active
        End If

        bIsLayered = (GetDeviceCaps(GetDC(0), BITSPIXEL) >= 16)             'Ensure we have enough screen colors
      End If
    End If
  End With
End Sub

'Create the shadow windows
Private Sub CreateWindows()
  Const EX_STYLE As Long = ws_ex_layered Or WS_EX_TRANSPARENT Or WS_EX_NOPARENTNOTIFY

  hWndRt = CreateWindowExA(EX_STYLE, "#32770", vbNullString, WS_POPUP, 0, 0, 0, 0, hWndForm, 0, App.hInstance, 0)
  hWndBt = CreateWindowExA(EX_STYLE, "#32770", vbNullString, WS_POPUP, 0, 0, 0, 0, hWndForm, 0, App.hInstance, 0)
End Sub

'Display the right and bottom shadows
Private Sub DisplayShadows()
  If bIsLayered Then
    If IsWindowVisible(hWndForm) <> 0 Then
      With wp
        DisplayShadowSub .x + .cX, .Y + m_Depth, m_Depth, .cY, True
        DisplayShadowSub .x + m_Depth, .Y + .cY, .cX - m_Depth, m_Depth, False
      End With
    End If
  End If
End Sub

'Display the content of the specified shadow window
Private Sub DisplayShadowSub(ByVal x As Long, ByVal Y As Long, cX As Long, cY As Long, ByVal bRight As Boolean)
  Dim DC        As Long
  Dim iX        As Long
  Dim iY        As Long
  Dim hDib      As Long
  Dim hWin      As Long
  Dim nAlpha    As Long
  Dim aPixels() As Long
  Dim pBmpBits  As Long
  Dim pt0       As tPOINT
  Dim pt        As tPOINT
  Dim sz        As tSIZE
  Dim bs        As tBLENDFUNCTION
  Dim bmpHeader As tBITMAPINFOHEADER
  Dim SafeArray As tSAFEARRAY2D

  DC = CreateCompatibleDC(0)                                                'Get a screen compatible memory dc

  With bmpHeader                                                            'Initialize a bitmap header
    .biSize = Len(bmpHeader)                                                'Bitmap header size
    .biWidth = cX                                                           'Bitmap/window pixel width
    .biHeight = cY                                                          'Bitmap/window pixel height
    .biPlanes = 1                                                           'Graphics planes
    .biBitCount = 32                                                        '32bits per pixel BGRA (Blue, Green, Red, Alpha)
    .biSizeImage = cX * cY * 4                                              'Memory size, width * height * 32bit
  End With

  hDib = CreateDIBSection(DC, bmpHeader, 0, pBmpBits, 0, 0)                 'Create a device independant bitmap as per the header, compatible with the dc (compatible with the screen)

  With SafeArray                                                            'Construct a VB safearray header that matches the specs of the bitmap
    .cbElements = 4                                                         '4 bytes per element - 32bits per pixel
    .cDims = 2                                                              'We'll treat the pixels as a two dimensional (x, y) array
    .pvData = pBmpBits                                                      'The data pointer points to the bitmap data (pixels)
    .Bounds(0).lLbound = 0                                                  'Lowest bound will be 0
    .Bounds(0).cElements = cY                                               'The number of elements
    .Bounds(1).lLbound = 0                                                  'Lowest bound will be 0
    .Bounds(1).cElements = cX                                               'The number of elements
  End With

  CopyMemory ByVal VarPtrArray(aPixels()), VarPtr(SafeArray), 4             'Copy the address of our safearray over the address of aPixels() safearray

  If bRight Then
    hWin = hWndRt
  Else
    hWin = hWndBt
  End If

  If bRight Then
    For iY = 0 To cY - 1
      If iY < cX Then
        nAlpha = (255 * iY) \ cX
      ElseIf iY >= (cY - cX) Then
        nAlpha = ((cY - iY) * 255) \ cX
      Else
        nAlpha = 255
      End If

      For iX = 0 To cX - 1
        aPixels(iX, iY) = MakeBGRA((nAlpha * (cX - iX)) \ cX)
      Next iX
    Next iY
  Else
    For iX = 0 To cX - 1
      If iX < cY Then
        nAlpha = (255 * iX) \ cY
      Else
        nAlpha = 255
      End If

      For iY = 0 To cY - 1
        aPixels(iX, iY) = MakeBGRA((nAlpha * iY) \ cY)
      Next iY
    Next iX
  End If

  If bRight Then
    If bIsLuna Then
      On Error Resume Next                                                  'Protect against shadow depths less than 5
      aPixels(cX - 1, cY - 1) = 0
      aPixels(cX - 2, cY - 1) = 0
      aPixels(cX - 3, cY - 1) = 0
      aPixels(cX - 4, cY - 1) = 0
      aPixels(cX - 5, cY - 1) = 0

      aPixels(cX - 1, cY - 2) = 0
      aPixels(cX - 2, cY - 2) = 0
      aPixels(cX - 3, cY - 2) = 0

      aPixels(cX - 1, cY - 3) = 0
      aPixels(cX - 2, cY - 3) = 0

      aPixels(cX - 1, cY - 4) = 0
      aPixels(cX - 1, cY - 5) = 0
      On Error GoTo 0
    End If
  End If

  CopyMemory ByVal VarPtrArray(aPixels()), 0&, 4
  
  With bs                                                                   'Setup the blend function
    .AlphaFormat = AC_SRC_ALPHA                                             'Use the alpha channel for individual pixel transparency
    .BlendFlags = 0
    .BlendOp = AC_SRC_OVER                                                  'Alpha overlay
    .SourceConstantAlpha = m_Transparency                                   'Alpha transparency for overall transparency
  End With

  pt.x = x                                                                  'Setup the window position and size data
  pt.Y = Y
  sz.cX = cX
  sz.cY = cY

  hDib = SelectObject(DC, hDib)                                             'Select the bitmap into the memory display context
  UpdateLayeredWindow hWin, DC, pt, sz, DC, pt0, 0, bs, ULW_ALPHA           'Do the layered update
  SelectObject DC, hDib                                                     'Trash the bitmap
  DeleteDC DC                                                               'Delete the memory display context
End Sub

'Return whether the Luna theme is active
Private Function IsLuna() As Boolean
  Dim hLib   As Long
  Dim nPos   As Long
  Dim sTheme As String
  Dim sName  As String

  hLib = LoadLibraryA("uxtheme.dll")

  If hLib <> 0 Then
    sTheme = String$(255, 0)

    GetCurrentThemeName StrPtr(sTheme), Len(sTheme), 0, 0, 0, 0
    nPos = InStr(1, sTheme, vbNullChar)

    If nPos > 0 Then
      sTheme = Left$(sTheme, nPos - 1)
      sName = String$(255, 0)

      GetThemeDocumentationProperty StrPtr(sTheme), StrPtr("ThemeName"), StrPtr(sName), Len(sName)
      nPos = InStr(1, sName, vbNullChar)

      If nPos > 0 Then
        sName = Left$(sName, nPos - 1)
        bIsLuna = (StrComp(sName, "Luna", vbTextCompare) = 0)
      End If
    End If

    FreeLibrary hLib
  End If
End Function

'Pre-multiply the shadow color with the passed alpha value. This is needed to get nice looking colors according to MSDN.
Private Function MakeBGRA(ByVal Alpha As Byte) As Long
  Dim fFactor As Double
  Dim BGRA    As tBGRA
    
  fFactor = CDbl(Alpha) / 255#                                              'Calculate the factor
  
'Note that nColor is in RGB format, part of this process is to convert to BGRA format
  With BGRA                                                                 'Blue, Green, Red, Alpha
    .Blue = ((m_Color And &HFF0000) \ &H10000) * fFactor                     'Factor the blue component
    .Green = ((m_Color And &HFF00&) \ &H100&) * fFactor                      'Factor the green component
    .Red = (m_Color And &HFF) * fFactor                                      'Factor the red component
    .Alpha = Alpha                                                          'Store the alpha value
  End With
  
'Copy the BGRA type to long
  CopyMemory MakeBGRA, BGRA, 4
End Function

'Show/hide the shadow windows
Private Sub Show(ByVal bShow As Boolean, Optional ByVal bForce As Boolean = False)
  If Not bForce Then
    If bLastShow = bShow Then
      Exit Sub
    End If
  End If

  bLastShow = bShow

  If bShow Then
    ShowWindow hWndRt, SW_SHOWNOACTIVATE
    ShowWindow hWndBt, SW_SHOWNOACTIVATE
  Else
    ShowWindow hWndRt, SW_HIDE
    ShowWindow hWndBt, SW_HIDE
  End If
End Sub

'Size/position the shadows
Private Sub SizePos()
  With wp
    If .Flags And SWP_HIDEWINDOW Then                                       'If the parent form is being hidden
      Show False                                                            'Hide the shadow windows also

    Else
      If .cX <> cX Then                                                     'If the parent's width has changed
        cX = .cX                                                            'Store the new width
        DisplayShadowSub .x + m_Depth, .Y + .cY, .cX - m_Depth, m_Depth, False
      End If

      If .cY <> cY Then                                                     'If the parent's height has changed
        cY = .cY                                                            'Store the new height
        DisplayShadowSub .x + .cX, .Y + m_Depth, m_Depth, .cY, True
      End If

      MoveWindow hWndRt, .x + .cX, .Y + m_Depth, m_Depth, .cY, False
      MoveWindow hWndBt, .x + m_Depth, .Y + .cY, .cX - m_Depth, m_Depth, False

      If (.Flags And SWP_SHOWWINDOW) Then
        Show True
      End If
    End If
  End With
End Sub


'-SelfSub code------------------------------------------------------------------------------------
'-The following routines are exclusively for the ssc_Subclass routines----------------------------
Private Function ssc_Subclass(ByVal lng_hWnd As Long, _
                    Optional ByVal lParamUser As Long = 0, _
                    Optional ByVal nOrdinal As Long = 1, _
                    Optional ByVal oCallback As Object = Nothing, _
                    Optional ByVal bIdeSafety As Boolean = True, _
                    Optional ByRef bUnicode As Boolean = False, _
                    Optional ByVal bIsAPIwindow As Boolean = False) As Boolean 'Subclass the specified window handle

    '*************************************************************************************************
    '* lng_hWnd   - Handle of the window to subclass
    '* lParamUser - Optional, user-defined callback parameter
    '* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
    '* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety - Optional, enable/disable IDE safety measures. There is not reason to set this to False
    '* bUnicode - Optional, if True, Unicode API calls should be made to the window vs ANSI calls
    '*            Parameter is byRef and its return value should be checked to know if ANSI to be used or not
    '* bIsAPIwindow - Optional, if True DestroyWindow will be called if IDE ENDs
    '*****************************************************************************************
    '** Subclass.asm - subclassing thunk
    '**
    '** Paul_Caton@hotmail.com
    '** Copyright free, use and abuse as you see fit.
    '**
    '** v2.0 Re-write by LaVolpe, based mostly on Paul Caton's original thunks....... 20070720
    '** .... Reorganized & provided following additional logic
    '** ....... Unsubclassing only occurs after thunk is no longer recursed
    '** ....... Flag used to bypass callbacks until unsubclassing can occur
    '** ....... Timer used as delay mechanism to free thunk memory afer unsubclassing occurs
    '** .............. Prevents crash when one window subclassed multiple times
    '** .............. More END safe, even if END occurs within the subclass procedure
    '** ....... Added ability to destroy API windows when IDE terminates
    '** ....... Added auto-unsubclass when WM_NCDESTROY received
    '*****************************************************************************************
    ' Subclassing procedure must be declared identical to the one at the end of this class (Sample at Ordinal #1)

    Dim z_Sc(0 To IDX_UNICODE) As Long                 'Thunk machine-code initialised here
    
    Const SUB_NAME      As String = "ssc_Subclass"     'This routine's name
    Const CODE_LEN      As Long = 4 * IDX_UNICODE + 4  'Thunk length in bytes
    Const PAGE_RWX      As Long = &H40&                'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&              'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&              'Release allocated memory flag
    Const GWL_WNDPROC   As Long = -4                   'SetWindowsLong WndProc index
    Const WNDPROC_OFF   As Long = &H60                 'Thunk offset to the WndProc execution address
    Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1)) 'Bytes to allocate per thunk, data + code + msg tables
    
  ' This is the complete listing of thunk offset values and what they point/relate to.
  ' Those rem'd out are used elsewhere or are initialized in Declarations section
  
  'Const IDX_RECURSION  As Long = 0     'Thunk data index of callback recursion count
  'Const IDX_SHUTDOWN   As Long = 1     'Thunk data index of the termination flag
  'Const IDX_INDEX      As Long = 2     'Thunk data index of the subclassed hWnd
   Const IDX_EBMODE     As Long = 3     'Thunk data index of the EbMode function address
   Const IDX_CWP        As Long = 4     'Thunk data index of the CallWindowProc function address
   Const IDX_SWL        As Long = 5     'Thunk data index of the SetWindowsLong function address
   Const IDX_FREE       As Long = 6     'Thunk data index of the VirtualFree function address
   Const IDX_BADPTR     As Long = 7     'Thunk data index of the IsBadCodePtr function address
   Const IDX_OWNER      As Long = 8     'Thunk data index of the Owner object's vTable address
  'Const IDX_PREVPROC   As Long = 9     'Thunk data index of the original WndProc
   Const IDX_CALLBACK   As Long = 10    'Thunk data index of the callback method address
  'Const IDX_BTABLE     As Long = 11    'Thunk data index of the Before table
  'Const IDX_ATABLE     As Long = 12    'Thunk data index of the After table
  'Const IDX_PARM_USER  As Long = 13    'Thunk data index of the User-defined callback parameter data index
   Const IDX_DW         As Long = 14    'Thunk data index of the DestroyWinodw function address
   Const IDX_ST         As Long = 15    'Thunk data index of the SetTimer function address
   Const IDX_KT         As Long = 16    'Thunk data index of the KillTimer function address
   Const IDX_EBX_TMR    As Long = 20    'Thunk code patch index of the thunk data for the delay timer
   Const IDX_EBX        As Long = 26    'Thunk code patch index of the thunk data
  'Const IDX_UNICODE    As Long = xx    'Must be UBound(subclass thunkdata)+1; index for unicode support
    
    Dim z_ScMem       As Long           'Thunk base address
    Dim nAddr         As Long
    Dim nID           As Long
    Dim nMyID         As Long
    Dim bIDE          As Boolean

    If IsWindow(lng_hWnd) = 0 Then      'Ensure the window handle is valid
        zError SUB_NAME, "Invalid window handle"
        Exit Function
    End If
    
    nMyID = GetCurrentProcessId                         'Get this process's ID
    GetWindowThreadProcessId lng_hWnd, nID              'Get the process ID associated with the window handle
    If nID <> nMyID Then                                'Ensure that the window handle doesn't belong to another process
        zError SUB_NAME, "Window handle belongs to another process"
        Exit Function
    End If
    
    If oCallback Is Nothing Then Set oCallback = Me     'If the user hasn't specified the callback owner
    
    nAddr = zAddressOf(oCallback, nOrdinal)             'Get the address of the specified ordinal method
    If nAddr = 0 Then                                   'Ensure that we've found the ordinal method
        zError SUB_NAME, "Callback method not found"
        Exit Function
    End If
        
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
    
    If z_ScMem <> 0 Then                                  'Ensure the allocation succeeded
    
      If z_scFunk Is Nothing Then Set z_scFunk = New Collection 'If this is the first time through, do the one-time initialization
      On Error GoTo CatchDoubleSub                              'Catch double subclassing
        z_scFunk.Add z_ScMem, "h" & lng_hWnd                    'Add the hWnd/thunk-address to the collection
      On Error GoTo 0
      
   'z_Sc (0) thru z_Sc(17) are used as storage for the thunks & IDX_ constants above relate to these thunk positions which are filled in below
    z_Sc(18) = &HD231C031: z_Sc(19) = &HBBE58960: z_Sc(21) = &H21E8F631: z_Sc(22) = &HE9000001: z_Sc(23) = &H12C&: z_Sc(24) = &HD231C031: z_Sc(25) = &HBBE58960: z_Sc(27) = &H3FFF631: z_Sc(28) = &H75047339: z_Sc(29) = &H2873FF23: z_Sc(30) = &H751C53FF: z_Sc(31) = &HC433913: z_Sc(32) = &H53FF2274: z_Sc(33) = &H13D0C: z_Sc(34) = &H18740000: z_Sc(35) = &H875C085: z_Sc(36) = &H820443C7: z_Sc(37) = &H90000000: z_Sc(38) = &H87E8&: z_Sc(39) = &H22E900: z_Sc(40) = &H90900000: z_Sc(41) = &H2C7B8B4A: z_Sc(42) = &HE81C7589: z_Sc(43) = &H90&: z_Sc(44) = &H75147539: z_Sc(45) = &H6AE80F: z_Sc(46) = &HD2310000: z_Sc(47) = &HE8307B8B: z_Sc(48) = &H7C&: z_Sc(49) = &H7D810BFF: z_Sc(50) = &H8228&: z_Sc(51) = &HC7097500: z_Sc(52) = &H80000443: z_Sc(53) = &H90900000: z_Sc(54) = &H44753339: z_Sc(55) = &H74047339: z_Sc(56) = &H2473FF3F: z_Sc(57) = &HFFFFFC68
    z_Sc(58) = &H2475FFFF: z_Sc(59) = &H811453FF: z_Sc(60) = &H82047B: z_Sc(61) = &HC750000: z_Sc(62) = &H74387339: z_Sc(63) = &H2475FF07: z_Sc(64) = &H903853FF: z_Sc(65) = &H81445B89: z_Sc(66) = &H484443: z_Sc(67) = &H73FF0000: z_Sc(68) = &H646844: z_Sc(69) = &H56560000: z_Sc(70) = &H893C53FF: z_Sc(71) = &H90904443: z_Sc(72) = &H10C261: z_Sc(73) = &H53E8&: z_Sc(74) = &H3075FF00: z_Sc(75) = &HFF2C75FF: z_Sc(76) = &H75FF2875: z_Sc(77) = &H2473FF24: z_Sc(78) = &H891053FF: z_Sc(79) = &H90C31C45: z_Sc(80) = &H34E30F8B: z_Sc(81) = &H1078C985: z_Sc(82) = &H4C781: z_Sc(83) = &H458B0000: z_Sc(84) = &H75AFF228: z_Sc(85) = &H90909023: z_Sc(86) = &H8D144D8D: z_Sc(87) = &H8D503443: z_Sc(88) = &H75FF1C45: z_Sc(89) = &H2C75FF30: z_Sc(90) = &HFF2875FF: z_Sc(91) = &H51502475: z_Sc(92) = &H2073FF52: z_Sc(93) = &H902853FF: z_Sc(94) = &H909090C3: z_Sc(95) = &H74447339: z_Sc(96) = &H4473FFF7
    z_Sc(97) = &H4053FF56: z_Sc(98) = &HC3447389: z_Sc(99) = &H89285D89: z_Sc(100) = &H45C72C75: z_Sc(101) = &H800030: z_Sc(102) = &H20458B00: z_Sc(103) = &H89145D89: z_Sc(104) = &H81612445: z_Sc(105) = &H4C4&: z_Sc(106) = &H1862FF00

    ' cache callback related pointers & offsets
      z_Sc(IDX_EBX) = z_ScMem                                                 'Patch the thunk data address
      z_Sc(IDX_EBX_TMR) = z_ScMem                                             'Patch the thunk data address
      z_Sc(IDX_INDEX) = lng_hWnd                                              'Store the window handle in the thunk data
      z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                   'Store the address of the before table in the thunk data
      z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
      z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback owner's object address in the thunk data
      z_Sc(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data
      z_Sc(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data
      
      ' validate unicode request & cache unicode usage
      If bUnicode Then bUnicode = (IsWindowUnicode(lng_hWnd) <> 0&)
      z_Sc(IDX_UNICODE) = bUnicode                                            'Store whether the window is using unicode calls or not
      
      ' get function pointers for the thunk
      If bIdeSafety = True Then                                               'If the user wants IDE protection
          Debug.Assert zInIDE(bIDE)
          If bIDE = True Then z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode", bUnicode) 'Store the EbMode function address in the thunk data
                                                        '^^ vb5 users, change vba6 to vba5
      End If
      If bIsAPIwindow Then                                                    'If user wants DestroyWindow sent should IDE end
          z_Sc(IDX_DW) = zFnAddr("user32", "DestroyWindow", bUnicode)
      End If
      z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree", bUnicode)           'Store the VirtualFree function address in the thunk data
      z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", bUnicode)        'Store the IsBadCodePtr function address in the thunk data
      z_Sc(IDX_ST) = zFnAddr("user32", "SetTimer", bUnicode)                  'Store the SetTimer function address in the thunk data
      z_Sc(IDX_KT) = zFnAddr("user32", "KillTimer", bUnicode)                 'Store the KillTimer function address in the thunk data
      
      If bUnicode Then
          z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcW", bUnicode)      'Store CallWindowProc function address in the thunk data
          z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongW", bUnicode)       'Store the SetWindowLong function address in the thunk data
          RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                    'Copy the thunk code/data to the allocated memory
          z_Sc(IDX_PREVPROC) = SetWindowLongW(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Set the new WndProc, return the address of the original WndProc
      Else
          z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA", bUnicode)      'Store CallWindowProc function address in the thunk data
          z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA", bUnicode)       'Store the SetWindowLong function address in the thunk data
          RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                    'Copy the thunk code/data to the allocated memory
          z_Sc(IDX_PREVPROC) = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Set the new WndProc, return the address of the original WndProc
      End If
      If z_Sc(IDX_PREVPROC) = 0 Then                                          'Ensure the new WndProc was set correctly
          zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
          GoTo ReleaseMemory
      End If
      'Store the original WndProc address in the thunk data
      RtlMoveMemory z_ScMem + IDX_PREVPROC * 4, VarPtr(z_Sc(IDX_PREVPROC)), 4&
      ssc_Subclass = True                                                     'Indicate success
      
    Else
        zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
        
    End If

 Exit Function                                                                'Exit ssc_Subclass
    
CatchDoubleSub:
 zError SUB_NAME, "Window handle is already subclassed"
      
ReleaseMemory:
      VirtualFree z_ScMem, 0, MEM_RELEASE                                     'ssc_Subclass has failed after memory allocation, so release the memory
      
End Function

'Terminate all subclassing
Private Sub ssc_Terminate()
    ' can be made public, can be removed & zTerminateThunks can be called instead
    zTerminateThunks SubclassThunk
End Sub

''UnSubclass the specified window handle
'Private Sub ssc_UnSubclass(ByVal lng_hWnd As Long)
'    ' can be made public, can be removed & zUnthunk can be called instead
'    zUnThunk lng_hWnd, SubclassThunk
'End Sub

'Add the message value to the window handle's specified callback table
Private Sub ssc_AddMsg(ByVal lng_hWnd As Long, ByVal When As eMsgWhen, ParamArray Messages() As Variant)
    
    Dim z_ScMem       As Long                                   'Thunk base address
    
    z_ScMem = zMap_VFunction(lng_hWnd, SubclassThunk)           'Ensure that the thunk hasn't already released its memory
    If z_ScMem Then
      Dim M As Long
      For M = LBound(Messages) To UBound(Messages)
        Select Case VarType(Messages(M))                        ' ensure no strings, arrays, doubles, objects, etc are passed
        Case vbByte, vbInteger, vbLong
            If When And MSG_BEFORE Then                         'If the message is to be added to the before original WndProc table...
              If zAddMsg(Messages(M), IDX_BTABLE, z_ScMem) = False Then 'Add the message to the before table
                When = (When And Not MSG_BEFORE)
              End If
            End If
            If When And MSG_AFTER Then                          'If message is to be added to the after original WndProc table...
              If zAddMsg(Messages(M), IDX_ATABLE, z_ScMem) = False Then 'Add the message to the after table
                When = (When And Not MSG_AFTER)
              End If
            End If
        End Select
      Next
    End If
End Sub

'Add the message to the specified table of the window handle
Private Function zAddMsg(ByVal uMsg As Long, ByVal nTable As Long, ByVal z_ScMem As Long) As Boolean
      Dim nCount As Long                                                        'Table entry count
      Dim nBase  As Long
      Dim i      As Long                                                        'Loop index
    
      zAddMsg = True
      nBase = zData(nTable, z_ScMem)                                            'Map zData() to the specified table
      
      If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
        nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
      Else
        
        nCount = zData(0, nBase)                                                'Get the current table entry count
        For i = 1 To nCount                                                     'Loop through the table entries
          If zData(i, nBase) = 0 Then                                           'If the element is free...
            zData(i, nBase) = uMsg                                              'Use this element
            GoTo Bail                                                           'Bail
          ElseIf zData(i, nBase) = uMsg Then                                    'If the message is already in the table...
            GoTo Bail                                                           'Bail
          End If
        Next i                                                                  'Next message table entry
    
        nCount = i                                                             'On drop through: i = nCount + 1, the new table entry count
        If nCount > MSG_ENTRIES Then                                           'Check for message table overflow
          zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
          zAddMsg = False
          GoTo Bail
        End If
        
        zData(nCount, nBase) = uMsg                                            'Store the message in the appended table entry
      End If
    
      zData(0, nBase) = nCount                                                 'Store the new table entry count
Bail:
End Function

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long, ByVal z_ScMem As Long)
      Dim nCount As Long                                                        'Table entry count
      Dim nBase  As Long
      Dim i      As Long                                                        'Loop index
    
      nBase = zData(nTable, z_ScMem)                                            'Map zData() to the specified table
    
      If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
        zData(0, nBase) = 0                                                     'Zero the table entry count
      Else
        nCount = zData(0, nBase)                                                'Get the table entry count
        
        For i = 1 To nCount                                                     'Loop through the table entries
          If zData(i, nBase) = uMsg Then                                        'If the message is found...
            zData(i, nBase) = 0                                                 'Null the msg value -- also frees the element for re-use
            GoTo Bail                                                           'Bail
          End If
        Next i                                                                  'Next message table entry
        
       ' zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
      End If
Bail:
End Sub

'-The following routines are used for each of the three types of thunks ----------------------------

'Maps zData() to the memory address for the specified thunk type
Private Function zMap_VFunction(vFuncTarget As Long, _
                                vType As eThunkType, _
                                Optional oCallback As Object, _
                                Optional bIgnoreErrors As Boolean) As Long
    
    Dim thunkCol As Collection
    Dim colID As String
    Dim z_ScMem       As Long         'Thunk base address
    
    If vType = SubclassThunk Then
        Set thunkCol = z_scFunk
        colID = "h" & vFuncTarget
    Else
        zError "zMap_Vfunction", "Invalid thunk type passed"
        Exit Function
    End If
    
    If thunkCol Is Nothing Then
        zError "zMap_VFunction", "Thunk hasn't been initialized"
    Else
        If thunkCol.Count Then
            On Error GoTo Catch
            z_ScMem = thunkCol(colID)               'Get the thunk address
            If IsBadCodePtr(z_ScMem) Then z_ScMem = 0&
            zMap_VFunction = z_ScMem
        End If
    End If
    Exit Function                                               'Exit returning the thunk address
    
Catch:
    ' error ignored when zUnThunk is called, error handled there
    If Not bIgnoreErrors Then zError "zMap_VFunction", "Thunk type for " & vType & " does not exist"
End Function

' sets/retrieves data at the specified offset for the specified memory address
Private Property Get zData(ByVal nIndex As Long, ByVal z_ScMem As Long) As Long
  RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal z_ScMem As Long, ByVal nValue As Long)
  RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4
End Property

'Error handler
Private Sub zError(ByRef sRoutine As String, ByVal sMsg As String)
  ' Note. These two lines can be rem'd out if you so desire. But don't remove the routine
'  App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
'  MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String, ByVal asUnicode As Boolean) As Long
  If asUnicode Then
    zFnAddr = GetProcAddress(GetModuleHandleW(StrPtr(sDLL)), sProc)         'Get the specified procedure address
  Else
    zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                 'Get the specified procedure address
  End If
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
  ' ^^ FYI VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
    ' Note: used both in subclassing and hooking routines
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim i     As Long                                                         'Loop index
  Dim j     As Long                                                         'Loop limit
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, i, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form method
      If Not zProbe(nAddr + &H710, i, bSub) Then                            'Probe for a PropertyPage method
        If Not zProbe(nAddr + &H7A4, i, bSub) Then                          'Probe for a UserControl method
            Exit Function                                                   'Bail...
        End If
      End If
    End If
  End If
  
  i = i + 4                                                                 'Bump to the next entry
  j = i + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While i < j
    RtlMoveMemory VarPtr(nAddr), i, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    i = i + 4                                                               'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Do                                                             'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

Private Function zInIDE(ByRef bIDE As Boolean) As Boolean
    ' only called in IDE, never called when compiled
    bIDE = True
    zInIDE = bIDE
End Function

Private Sub zUnThunk(ByVal thunkID As Long, ByVal vType As eThunkType, Optional ByVal oCallback As Object)

    ' thunkID, depends on vType:
    '   - Subclassing:  the hWnd of the window subclassed
    On Error Resume Next ' In case there is a collection problem

    Const IDX_SHUTDOWN  As Long = 1
    
    Dim z_ScMem       As Long                       'Thunk base address
    
    z_ScMem = zMap_VFunction(thunkID, vType, oCallback, True)
    Select Case vType
    Case SubclassThunk
        If z_ScMem Then                                 'Ensure that the thunk hasn't already released its memory
            zData(IDX_SHUTDOWN, z_ScMem) = 1            'Set the shutdown indicator
            zDelMsg ALL_MESSAGES, IDX_BTABLE, z_ScMem   'Delete all before messages
            zDelMsg ALL_MESSAGES, IDX_ATABLE, z_ScMem   'Delete all after messages
        End If
        z_scFunk.Remove "h" & thunkID                   'Remove the specified thunk from the collection
    End Select

End Sub

Private Sub zTerminateThunks(ByVal vType As eThunkType)

    ' Terminates all thunks of a specific type
    ' Any subclassing, hooking, recurring callbacks should have already been canceled

    Dim i As Long
    Dim oCallback As Object
    Dim thunkCol As Collection
    Dim z_ScMem       As Long                           'Thunk base address
    Const INDX_OWNER As Long = 0
    
    Select Case vType
    Case SubclassThunk
        Set thunkCol = z_scFunk
    Case Else
        Exit Sub
    End Select
    
    If Not (thunkCol Is Nothing) Then                 'Ensure that hooking has been started
      With thunkCol
        For i = .Count To 1 Step -1                   'Loop through the collection of hook types in reverse order
          z_ScMem = .Item(i)                          'Get the thunk address
          If IsBadCodePtr(z_ScMem) = 0 Then           'Ensure that the thunk hasn't already released its memory
            Select Case vType
                Case SubclassThunk
                    zUnThunk zData(IDX_INDEX, z_ScMem), SubclassThunk    'Unsubclass
            End Select
          End If
        Next i                                        'Next member of the collection
      End With
      Set thunkCol = Nothing                         'Destroy the hook/thunk-address collection
    End If


End Sub

'- ordinal #1, example of a subclassing procedure used with ssc_Subclass
Private Sub myWndProc(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc.
'*              Not applicable with After messages
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter. Change vartype as needed (i.e., Object, UDT, etc)
'*************************************************************************************************
    
  Dim tLayered As Boolean

  If uMsg = WM_WINDOWPOSCHANGED Then                                        'Select the message number
    CopyMemory wp, ByVal lParam, Len(wp)                                    'Copy the WINDOWPOS data
    SizePos                                                                 'Position shadows

  ElseIf uMsg = WM_SIZE Then                                                'Parent form has been minimized/restored/maximized
    If wParam = SIZE_RESTORED Then                                          'If the parent has been restored
      If IsWindowVisible(hWndForm) = 1 Then                                 'If the parent is visible
        Show True                                                           'Show shadows
      End If
    End If

  ElseIf uMsg = WM_DISPLAYCHANGE Then                                       'The display settings have been changed
    tLayered = bIsLayered                                                   'Save the the existing layered status
    CheckEnvironment                                                        'Check OS, color depth etc.
    If tLayered Then                                                        'If we were layer enabled
      If Not bIsLayered Then                                                'But not anymore
        DestroyWindow hWndBt                                                'Get rid of the shadow
        DestroyWindow hWndRt                                                'Get rid of the shadow
      End If
     Else
      If bIsLayered Then                                                    'If we're now layered but previously weren't
        CreateWindows                                                       'Creates the shadow windows
        DisplayShadows                                                      'Display the shadows
        If IsWindowVisible(hWndForm) Then                                   'If the parent is visible
          Show True, True                                                   'Display the shadows
        End If
      End If
    End If

  ElseIf uMsg = WM_THEMECHANGED Then                                        'Windows theme has changed
    CheckEnvironment                                                        'Check OS, color depth etc.
    DisplayShadows                                                          'Display the shadows

  End If
    
' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub



