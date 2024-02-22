Attribute VB_Name = "modRegistry"
Option Explicit


'**************************************
'Windows API/Global Declarations for :Re
'     ad/Write to the Registry
'**************************************
Private Declare Function SystemParametersInfo Lib "user32" Alias _
   "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam _
   As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long


Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" ( _
    ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, ByRef phkResult As Long) As Long


Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" ( _
    ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
    ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long

Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" ( _
    ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
    ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long


Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" ( _
    ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" ( _
    ByVal hKey As Long, ByVal lpValueName As String) As Long


Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" ( _
    ByVal hKey As Long, ByVal lpSubKey As String) As Long

' Reg Key ROOT Types...
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const REG_OPTION_NON_VOLATILE = 0
Public Const SYNCHRONIZE = &H100000
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_WRITE = &H20006
Public Const KEY_ALL_ACCESS = &H2003F
Public Const KEY_READ = _
((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
' Registry location
Public Const gREGKEYLocation = "SOFTWARE\Your Company Name\Your App Name\Your Current Version"
Public Const gREGKEYXPos = "XPos"
Public Const gREGKEYYPos = "YPos"
Public Const gREGKEYWidth = "Width"
Public Const gREGKEYHeight = "Height"
Public Const gREGKEYWindowState = "WindowState"
Public Const ERROR_SUCCESS = 0&
'
Public Const SPIF_UPDATEINIFILE = &H1
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPIF_SENDWININICHANGE = &H2



Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, _
    ByRef KeyVal As String) As Boolean
    Dim i As Long
    Dim rc As Long
    Dim hKey As Long
    Dim KeyValType As Long
    Dim tmpVal As String
    Dim KeyValSize As Long
    
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    
    tmpVal = String$(1024, 0)
    KeyValSize = 1024
    
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    

    If KeyValSize > 0 Then
        If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then
            tmpVal = Left(tmpVal, KeyValSize - 1)
        Else
            tmpVal = Left(tmpVal, KeyValSize)
        End If
        
    
    
        Select Case KeyValType
            Case REG_DWORD
    
    
            For i = Len(tmpVal) To 1 Step -1
                KeyVal = KeyVal + Format(Hex(Asc(Mid(tmpVal, i, 1))), "00")
            Next
            KeyVal = Format$("&h" + KeyVal)
            Case REG_SZ
            KeyVal = tmpVal
        End Select
    End If
GetKeyValue = True
rc = RegCloseKey(hKey)
Exit Function

GetKeyError:
GetKeyValue = False
rc = RegCloseKey(hKey)
End Function


Public Function SetKeyValue(KeyRoot As Long, KeyName As String, lType As Long, _
SubKeyRef As String, KeyVal As Variant) As Boolean
    Dim rc As Long
    Dim hKey As Long
    
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
    
    If (rc <> ERROR_SUCCESS) Then
        Call RegCreateKey(KeyRoot, KeyName, hKey)
    End If

    Select Case lType
        Case REG_SZ
        rc = RegSetValueEx(hKey, SubKeyRef, 0&, REG_SZ, ByVal CStr(KeyVal & Chr$(0)), Len(KeyVal))
        Case REG_BINARY
        rc = RegSetValueEx(hKey, SubKeyRef, 0&, REG_BINARY, ByVal CStr(KeyVal & Chr$(0)), Len(KeyVal))
        Case REG_DWORD
        rc = RegSetValueEx(hKey, SubKeyRef, 0&, REG_DWORD, CLng(KeyVal), 4)
    End Select
If (rc <> ERROR_SUCCESS) Then GoTo SetKeyError

SetKeyValue = True
rc = RegCloseKey(hKey)

Exit Function
SetKeyError:
KeyVal = ""
SetKeyValue = False
rc = RegCloseKey(hKey)
End Function


Public Function DeleteRegValue(KeyName As String, SubKeyRef As String) As Boolean
    Dim rc As Long
    Dim hKey As Long
    rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, KeyName, 0, KEY_ALL_ACCESS, hKey)
    If (rc <> ERROR_SUCCESS) Then GoTo DeleteKeyError
    rc = RegDeleteValue(hKey, SubKeyRef)
    If (rc <> ERROR_SUCCESS) Then GoTo DeleteKeyError
    DeleteRegValue = True
    Exit Function
DeleteKeyError:
    DeleteRegValue = False
    
End Function


Public Function DeleteRegKey(KeyName As String) As Boolean
    
    Dim rc As Long
    'All sub keys must be deleted for this t
    '     o work.
    'If you create key under your original k
    '     ey, you
    'need to delete it forst.
    rc = RegDeleteKey(HKEY_LOCAL_MACHINE, KeyName)
    DeleteRegKey = IIf(rc = ERROR_SUCCESS, True, False)
End Function

        

Public Function SysParInfo(ByVal uAction As Long, ByVal uParam _
   As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
    SysParInfo = SystemParametersInfo(uAction, uParam, lpvParam, fuWinIni)
End Function

'Simple write routine for the registry
Public Sub WriteRegistry(ByVal Group As Long, ByVal Section As String, ByVal Key As String, NewVal As String)
    Dim lResult   As Long
    lResult = SetKeyValue(Group, Section, REG_SZ, Key, NewVal)
End Sub

Public Sub ReadRegistry(ByVal Group As Long, ByVal Section As String, ByVal Key As String, KeyVal As String)
    Dim lResult   As Long
    lResult = GetKeyValue(Group, Section, Key, KeyVal)
End Sub




