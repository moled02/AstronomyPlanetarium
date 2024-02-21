VERSION 5.00
Begin VB.Form FrmSysTray 
   BorderStyle     =   0  'None
   ClientHeight    =   1950
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   2400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   2400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox attFlash1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   720
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   1320
      Width           =   540
   End
   Begin VB.PictureBox attFlash2 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   1320
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   1320
      Width           =   540
   End
   Begin VB.PictureBox errFlash2 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   1320
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   720
      Width           =   540
   End
   Begin VB.PictureBox errFlash1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   720
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   720
      Width           =   540
   End
   Begin VB.PictureBox Default 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   120
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox Flash2 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   1320
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   120
      Width           =   540
   End
   Begin VB.PictureBox Flash1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   720
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   120
      Width           =   540
   End
   Begin VB.Timer TmrFlash 
      Interval        =   500
      Left            =   1920
      Top             =   120
   End
   Begin VB.Menu mPopupMenu 
      Caption         =   "&PopupMenu"
      Begin VB.Menu mAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mMaximize 
         Caption         =   "Ma&ximize"
      End
      Begin VB.Menu mRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mMinimize 
         Caption         =   "&Minimize"
      End
      Begin VB.Menu mSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "FrmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public WithEvents FSys As Form
Attribute FSys.VB_VarHelpID = -1
Public Event Click(ClickWhat As String)
Public Event TIcon(F As Form)

Private nid As NOTIFYICONDATA
Private LastWindowState As Integer

Public Property Let Tooltip(Value As String)
    nid.szTip = Value & vbNullChar
End Property

Public Property Get Tooltip() As String
    Tooltip = nid.szTip
End Property

Public Property Let Interval(Value As Integer)
    TmrFlash.Interval = Value
    UpdateIcon NIM_MODIFY
End Property

Public Property Get Interval() As Integer
    Interval = TmrFlash.Interval
End Property

Public Property Let TrayIcon(Value)
    TmrFlash.Enabled = False
    On Error Resume Next
    ' Value can be a picturebox, image, form or string
    Select Case TypeName(Value)
      Case "PictureBox", "Image"
          Me.Icon = Value.Picture
          TmrFlash.Enabled = False
          RaiseEvent TIcon(Me)
      Case "String"
          If (UCase(Value) = "DEFAULT") Then
            TmrFlash.Enabled = False
            Me.Icon = Default.Picture
            RaiseEvent TIcon(Me)
          ElseIf (UCase(Value) = "FLASH_ERROR") Then
            Flash1.Picture = errFlash1.Picture
            Flash2.Picture = errFlash2.Picture
            TmrFlash.Enabled = True
            Me.Icon = Flash1.Picture
            RaiseEvent TIcon(Me)
          ElseIf (UCase(Value) = "FLASH_ATTENTION") Then
            Flash1.Picture = attFlash1.Picture
            Flash2.Picture = attFlash2.Picture
            TmrFlash.Enabled = True
            Me.Icon = Flash1.Picture
            RaiseEvent TIcon(Me)
          Else
            ' String is filename; load icon from picture file.
            TmrFlash.Enabled = False
            Me.Icon = LoadPicture(Value)
            RaiseEvent TIcon(Me)
          End If
      Case Else
          ' It's a form ?
          Me.Icon = Value.Icon
          RaiseEvent TIcon(Me)
   End Select
   UpdateIcon NIM_MODIFY
End Property

Private Sub Form_Load()
    Me.Icon = Default
    RaiseEvent TIcon(Me)
    Me.Visible = False
    TmrFlash.Enabled = False
    'Tooltip = "Plaatselijke sterrentijd = " & frmOscarMonitor.stBar.Panels(2).Text
    mAbout.Caption = "About " & App.Title
    UpdateIcon NIM_ADD
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Result As Long
    Dim msg As Long
   
    ' The Form_MouseMove is intercepted to give systray mouse events.
    If Me.ScaleMode = vbPixels Then
        msg = x
    Else
        msg = x / Screen.TwipsPerPixelX
    End If
      
    Select Case msg
      Case WM_RBUTTONUP
          ' Popup menu: selectively enable items dependent on context.
          Select Case FSys.Visible
            Case True
              Select Case FSys.WindowState
                  Case vbMaximized
                      mMaximize.Enabled = False
                      mMinimize.Enabled = True
                      mRestore.Enabled = False
                  Case vbNormal
                      mMaximize.Enabled = True
                      mMinimize.Enabled = True
                      mRestore.Enabled = False
                  Case vbMinimized
                      mMaximize.Enabled = True
                      mMinimize.Enabled = False
                      mRestore.Enabled = True
                  Case Else
                      mMaximize.Enabled = True
                      mMinimize.Enabled = True
                      mRestore.Enabled = True
               End Select
            Case Else
                mRestore.Enabled = True
                mMaximize.Enabled = True
                mMinimize.Enabled = False
          End Select
         
          RaiseEvent Click("RBUTTONUP")
          SetForegroundWindow Me.hWnd 'So the menu will close after losing focus
          PopupMenu mPopupMenu
      Case WM_LBUTTONDBLCLK
          RaiseEvent Click("LBUTTONDBLCLK")
          mRestore_Click
      'Case WM_RBUTTONDBLCLK
      '    RaiseEvent Click("RBUTTONDBLCLK")
      'Case WM_RBUTTONDOWN
      '    RaiseEvent Click("RBUTTONDOWN")
      'Case WM_LBUTTONDOWN
      '    RaiseEvent Click("LBUTTONDOWN")
      'Case WM_LBUTTONUP
      '    RaiseEvent Click("LBUTTONUP")
      'Case WM_MBUTTONDBLCLK
      '    RaiseEvent Click("MBUTTONDBLCLK")
      'Case WM_MBUTTONDOWN
      '    RaiseEvent Click("MBUTTONDOWN")
      'Case WM_MBUTTONUP
      '    RaiseEvent Click("MBUTTONUP")
      'Case WM_MOUSEMOVE
      '    RaiseEvent Click("MOUSEMOVE")
      'Case Else
      '    RaiseEvent Click("OTHER....: " & Format$(msg))
    End Select
End Sub

Private Sub FSys_Resize()
    ' Event generated my main form. WindowState is stored in LastWindowState, so that
    ' it may be re- set when the menu item "Restore" is selected.
    If (FSys.WindowState <> vbMinimized) Then LastWindowState = FSys.WindowState
End Sub

Private Sub FSys_Unload(Cancel As Integer)
    ' Important: remove icon from tray, and unload this form when
    ' the main form is unloaded.
    UpdateIcon NIM_DELETE
    Unload Me
End Sub

Private Sub mAbout_Click()
    FSys.About
End Sub

Private Sub mMaximize_Click()
    FSys.WindowState = vbMaximized
    FSys.Show
End Sub

Private Sub mMinimize_Click()
    FSys.WindowState = vbMinimized
End Sub

Public Sub mExit_Click()
    Unload FSys
End Sub

Private Sub mRestore_Click()
    ' Don't "restore"  FSys is visible and not minimized.
    If (FSys.Visible And FSys.WindowState <> vbMinimized) Then
        SetForegroundWindow FSys.hWnd
        Exit Sub
    End If
    ' Restore LastWindowState
    FSys.WindowState = LastWindowState
    FSys.Visible = True
    SetForegroundWindow FSys.hWnd
End Sub

Private Sub UpdateIcon(Value As Long)
    ' Used to add, modify and delete icon.
    With nid
      .cbSize = Len(nid)
      .hWnd = Me.hWnd
      .uID = vbNull
      .uFlags = NIM_DELETE Or NIF_TIP Or NIM_MODIFY
      .uCallbackMessage = WM_MOUSEMOVE
      .hIcon = Me.Icon
    End With
    Shell_NotifyIcon Value, nid
End Sub

Public Sub MeQueryUnload(ByRef F As Form, Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
      ' Cancel by setting Cancel = 1, minimize and hide main window.
      Cancel = 1
      F.WindowState = vbMinimized
      F.Hide
    End If
End Sub

Public Sub MeResize(ByRef F As Form)
    Select Case F.WindowState
      Case vbNormal, vbMaximized
         ' Store LastWindowState
         LastWindowState = F.WindowState
      Case vbMinimized
         F.Hide
    End Select
End Sub

Private Sub TmrFlash_Timer()
    ' Change icon.
    Static LastIconWasFlash1 As Boolean
    LastIconWasFlash1 = Not LastIconWasFlash1
    Select Case LastIconWasFlash1
      Case True
         Me.Icon = Flash2
      Case Else
         Me.Icon = Flash1
    End Select
    RaiseEvent TIcon(Me)
    UpdateIcon NIM_MODIFY
End Sub
