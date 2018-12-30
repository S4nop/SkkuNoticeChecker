Attribute VB_Name = "ModReg"
Option Explicit

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const KEY_WRITE = ((&H20000 Or &H2 Or &H4) And (Not &H100000))

Function RegGetSectionValueName(ByVal SubKey As String) As String
    Dim hKey As Long
    Dim m, ByteLen As Long
    Dim KeyValueName As String
    Dim KeyValue As String
    
    RegOpenKey HKEY_LOCAL_MACHINE, SubKey, hKey
    
    ByteLen = 255&: m = 0&
    KeyValueName = String$(255, 0)
    KeyValue = String$(255, 0) '이거 안하니까 튕긴다 ㄷㄷ
    
    Do Until Not RegEnumValue(hKey, m, KeyValueName, 255, 0&, ByVal 0&, KeyValue, ByteLen) = 0&
        RegGetSectionValueName = IIf(RegGetSectionValueName = "", Split(KeyValueName, vbNullChar, 2)(0), RegGetSectionValueName & "|" & Split(KeyValueName, vbNullChar, 2)(0))
        ByteLen = 255&: m = m + 1
        KeyValueName = String$(255, 0): KeyValue = String$(255, 0)
    Loop
    RegCloseKey hKey
End Function
Function SHRegWriteString(ByVal SubKey As String, ByVal ValueName As String, ByVal szData As String) As Integer
    Dim hKey As Long
    Dim bData() As Byte
    
    bData = StrConv(szData, vbFromUnicode)
    'REG_OTPION_NON_VOLATILE = &H0
    If Not RegCreateKeyEx(HKEY_LOCAL_MACHINE, _
                            SubKey, _
                            0, _
                            0, _
                            0, _
                            KEY_WRITE, _
                            ByVal 0&, _
                            hKey, _
                            0) = 0& Then: Exit Function
                            
    SHRegWriteString = IIf(RegSetValueEx(hKey, _
                            ValueName, _
                            0, _
                            1, _
                            bData(0), _
                            UBound(bData) + 2) = 0, 1, 0)
    RegCloseKey hKey
End Function
Function SHRegDelValue(ByVal SubKey As String, ByVal ValueName As String) As Integer
    Dim hKey As Long
    
    If Not RegOpenKey(HKEY_LOCAL_MACHINE, SubKey, hKey) = 0 Then: Exit Function
    SHRegDelValue = IIf(RegDeleteValue(hKey, ValueName) = 0, 1, 0)
    RegCloseKey hKey
End Function



