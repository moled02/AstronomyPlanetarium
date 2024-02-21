VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpecialFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
'Project:   Astronomie
'Module:    clsSpecialFolder
'Doel:      Opvragen van (user-)specifieke folders
'
'Wijzigingen
'dd/mm/jjjj door omschrijving
'---------- ---- --------------------------------------------------------------
'27/01/2004 SMA  Initiële versie
'26/02/2004 SMA  Added Temporary folder (no longer "c:\temp" for WinXP)
'19-03-2004 SMA  Database folder niet in "Mijn Documenten" maar via registry
'04-01-2006 DMO  Migratie Sybase
'******************************************************************************

Option Explicit
Option Base 0

' PRIVATE CONSTANTS / VARIABLES
'----------------------------------------
Private Const m_MOD = "clsOscarSpecialFolder"
'
Private m_hWnd   As Long    'Typically set to 0 (for GUI when a dialup connection is needed)
Private m_objFSO As FileSystemObject
'
' API constants
Private Enum CSIDL_VALUES
    CSIDL_DESKTOP = &H0
    CSIDL_INTERNET = &H1
    CSIDL_PROGRAMS = &H2
    CSIDL_CONTROLS = &H3
    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_BITBUCKET = &HA
    CSIDL_STARTMENU = &HB
    CSIDL_MYDOCUMENTS = &HC
    CSIDL_MYMUSIC = &HD
    CSIDL_MYVIDEO = &HE
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_TEMPLATES = &H15
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_APPDATA = &H1A
    CSIDL_PRINTHOOD = &H1B
    CSIDL_LOCAL_APPDATA = &H1C
    CSIDL_ALTSTARTUP = &H1D
    CSIDL_COMMON_ALTSTARTUP = &H1E
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
    CSIDL_COMMON_APPDATA = &H23
    CSIDL_WINDOWS = &H24
    CSIDL_SYSTEM = &H25
    CSIDL_PROGRAM_FILES = &H26
    CSIDL_MYPICTURES = &H27
    CSIDL_PROFILE = &H28
    CSIDL_SYSTEMX86 = &H29
    CSIDL_PROGRAM_FILESX86 = &H2A
    CSIDL_PROGRAM_FILES_COMMON = &H2B
    CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
    CSIDL_COMMON_TEMPLATES = &H2D
    CSIDL_COMMON_DOCUMENTS = &H2E
    CSIDL_COMMON_ADMINTOOLS = &H2F
    CSIDL_ADMINTOOLS = &H30
    CSIDL_CONNECTIONS = &H31
    CSIDL_COMMON_MUSIC = &H35
    CSIDL_COMMON_PICTURES = &H36
    CSIDL_COMMON_VIDEO = &H37
    CSIDL_RESOURCES = &H38
    CSIDL_RESOURCES_LOCALIZED = &H39
    CSIDL_COMMON_OEM_LINKS = &H3A
    CSIDL_CDBURN_AREA = &H3B
    CSIDL_COMPUTERSNEARME = &H3D
    CSIDL_FLAG_PER_USER_INIT = &H800
    CSIDL_FLAG_NO_ALIAS = &H1000
    CSIDL_FLAG_DONT_VERIFY = &H4000
    CSIDL_FLAG_CREATE = &H8000
    CSIDL_FLAG_MASK = &HFF00
End Enum

Private Const SHGFP_TYPE_CURRENT = &H0  'Current path for user
Private Const SHGFP_TYPE_DEFAULT = &H1  'Default path

Private Const MAX_LENGTH = 260          'Buffer length for SHGetFolderPath

'Return values for SHGetFolderPath
Private Const S_OK = &H0                'Success
Private Const S_FALSE = &H1             'Folder is valid but does not exist
Private Const E_INVALIDARG = &H80070057 'Invalid CSIDL value

Private Declare Function SHGetFolderPath Lib "shfolder.dll" _
   Alias "SHGetFolderPathA" _
  (ByVal hwndOwner As Long, _
   ByVal nFolder As Long, _
   ByVal hToken As Long, _
   ByVal dwReserved As Long, _
   ByVal lpszPath As String) As Long

' PROPERTIES      PROPERTIES       PROPERTIES      PROPERTIES      PROPERTIES
'--------------------------------------------------------------------------
Public Property Let SetWindowHandle(hWnd As Long)
    ' Might not be neccesary
    m_hWnd = hWnd
End Property


Public Property Get TemporaryFolder() As String
    On Error GoTo Err_Handler
    Const cSUB As String = "TemporaryFolder"
    
    Dim sFolder As String

   
    ' try "C:\Documents and Settings\<user>\Local Settings\Temp"
    If sFolder = vbNullString Then
        sFolder = SpecialFolder(CSIDL_LOCAL_APPDATA)
        If sFolder <> vbNullString Then
            With m_objFSO
                sFolder = .GetAbsolutePathName(.BuildPath(sFolder, "..\Temp"))
                If Not .FolderExists(sFolder) Then
                    'Try to create
                    On Error Resume Next
                    .CreateFolder sFolder
                    On Error GoTo Err_Handler
                    sFolder = IIf(.FolderExists(sFolder), sFolder, vbNullString)
                End If
                sFolder = sFolder & "\Astronomie"
                If Not .FolderExists(sFolder) Then
                    'Try to create
                    On Error Resume Next
                    .CreateFolder sFolder
                    On Error GoTo Err_Handler
                    sFolder = IIf(.FolderExists(sFolder), sFolder, vbNullString)
                End If
            End With
        End If
    End If

    TemporaryFolder = sFolder
    Exit Property

Err_Handler:
    
End Property



' PRIVATE      PRIVATE       PRIVATE      PRIVATE      PRIVATE
'--------------------------------------------------------------------------
Private Function SpecialFolder(csidl As Long) As String
    Dim sBuffer As String
    
    sBuffer = Space(MAX_LENGTH)
    If SHGetFolderPath(m_hWnd, csidl, 0, SHGFP_TYPE_CURRENT, sBuffer) = S_OK Then
        'Returned C-string is null terminated
        SpecialFolder = Left(sBuffer, InStr(1, sBuffer, Chr(0)) - 1)
    Else
        SpecialFolder = vbNullString
    End If
End Function

' Initialize
Private Sub Class_Initialize()
  ' Create a new FileSystemObject
  Set m_objFSO = New FileSystemObject
End Sub

' Terminate
Private Sub Class_Terminate()
  ' Dispose the FileSystemObject
  Set m_objFSO = Nothing
End Sub



