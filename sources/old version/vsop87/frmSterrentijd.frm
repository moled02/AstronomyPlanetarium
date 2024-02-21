VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSterrentijd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plaatselijke Sterrentijd"
   ClientHeight    =   2295
   ClientLeft      =   195
   ClientTop       =   765
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLatitude 
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Text            =   "52.05"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtLongitude 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Text            =   "-05.08"
      Top             =   360
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   2025
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4436
            MinWidth        =   176
            Text            =   "Plaatselijke sterrentijd"
            TextSave        =   "Plaatselijke sterrentijd"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1614
            MinWidth        =   1605
            TextSave        =   "22-9-2006"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   291
            MinWidth        =   282
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrInterval 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4680
      Top             =   840
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "&Minimize"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Westerlengte"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Noorderbreedte"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Line lnMenu2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   6915
      Y1              =   35
      Y2              =   35
   End
   Begin VB.Line lnMenu1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   6915
      Y1              =   20
      Y2              =   20
   End
End
Attribute VB_Name = "frmSterrentijd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Dim WithEvents FormSys As FrmSysTray  'SystemTray functionality
Attribute FormSys.VB_VarHelpID = -1
Dim gAutoPos As New clsAutoPositioner 'AutoPos functionality
Dim gRedisplayingFirst As Boolean     'Signals first page of history is being refetched

'Registrypath
Private Const Pi = 3.14159265358979
Private Const Pi2 = Pi * 2

'Keep track of minutes passed
Private gMinuteCounter As Integer


Private Sub CmdExit_Click()
    'Unload Me
    WindowState = vbMinimized
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    About
End Sub

Public Sub About()
    MsgBox _
        "Berekenen van de plaatselijke sterrentijd." & vbNewLine & _
        "Ontwikkeld door Dominique Molenkamp." & vbNewLine & _
        "Version " & App.Major & "." & App.Minor & "." & App.Revision & "  Copyright © 2006", _
        vbInformation, "About " & App.Title
End Sub

Private Function max(i1, i2) As Integer
    max = IIf(i1 > i2, i1, i2)
End Function

Private Sub startTimer()
    With tmrInterval
        .Enabled = False
        .Interval = 100 'Event every minute
        .Enabled = True
        gMinuteCounter = 0
    End With
    updateTimerStatus
End Sub

Private Sub stopTimer()
    tmrInterval.Enabled = False
    stBar.Panels(1).Enabled = False
End Sub

Private Sub Form_Terminate()
Unload Me
End Sub

Private Sub tmrInterval_Timer()
    updateTimerStatus
End Sub
    
Private Sub updateTimerStatus()
Dim dattijd As Date
    Dim iNext As Integer
    dattijd = Now()
'    dattijd = #1/1/1986 11:43:00 AM#
    stBar.Panels(2).Text = Format(PlaatselijkeSterrentijd(dattijd), "hh:mm:ss")
    stBar.Panels(4).Text = Format(dattijd, "hh:mm:ss")
    FormSys.Tooltip = "Plaatselijke sterrentijd = " & stBar.Panels(2).Text
    FormSys.TrayIcon = "DEFAULT"
End Sub

Private Sub updateConnectionStatus(msg As String)
    stBar.Panels(1).Enabled = True
    stBar.Panels(1).Text = msg
End Sub



Private Function formatTime(iSeconds As Integer) As String
    If iSeconds = 0 Then
        formatTime = "< 1 second"
    ElseIf iSeconds = 1 Then
        formatTime = "1 second"
    ElseIf iSeconds < 60 Then
        formatTime = CStr(iSeconds) & " seconds"
    ElseIf iSeconds < 3600 Then
        formatTime = Format(iSeconds \ 60, "0#") & "m " & Format(iSeconds Mod 60, "0#") & "s"
    Else
        formatTime = Format(iSeconds \ 3600, "0#") & ":" & Format((iSeconds Mod 3600) \ 60, "0#") & ":" & Format(iSeconds Mod 60, "0#")
    End If
End Function


Private Function getRegItem(sSubKey As String, sItem As String) As String
End Function



Private Sub Form_Load()
    Dim sLastChar As String

    'For the SystemTray functionality
    Set FormSys = New FrmSysTray
    Load FormSys
    Set FormSys.FSys = Me
    
    'Auto resize/position
    'Line below menu
'    gAutoPos.AddAssignment Me.lnMenu1, Me, DELTA_WIDTH_RIGHT
    'gAutoPos.AddAssignment Me.lnMenu2, Me, DELTA_WIDTH_RIGHT
    'URL
'    gAutoPos.AddAssignment txtURL, Me, DELTA_WIDTH_RIGHT
    'Buttons
'    gAutoPos.AddAssignment CmdExit, Me, RELATIVE_POS_BOTTOM + RELATIVE_POS_RIGHT
'    gAutoPos.AddAssignment cmdCheckNow, Me, RELATIVE_POS_BOTTOM + RELATIVE_POS_RIGHT
'    gAutoPos.AddAssignment cmdReset, Me, RELATIVE_POS_BOTTOM + RELATIVE_POS_RIGHT
'    gAutoPos.AddAssignment optWinNt, Me, RELATIVE_POS_BOTTOM
'    gAutoPos.AddAssignment optWinXP, Me, RELATIVE_POS_BOTTOM
'    'TabStrip
'    gAutoPos.AddAssignment TabStrip, Me, DELTA_HEIGHT_BOTTOM + DELTA_WIDTH_RIGHT
    'Frames
'    gAutoPos.AddAssignment frmCheckDB, Me, DELTA_HEIGHT_BOTTOM + DELTA_WIDTH_RIGHT
'    gAutoPos.AddAssignment frmDownload, Me, DELTA_HEIGHT_BOTTOM + DELTA_WIDTH_RIGHT
'    gAutoPos.AddAssignment frmLog, Me, DELTA_HEIGHT_BOTTOM + DELTA_WIDTH_RIGHT
'    gAutoPos.AddAssignment frmPage, Me, DELTA_HEIGHT_BOTTOM + DELTA_WIDTH_RIGHT
'    gAutoPos.AddAssignment frmRegistry, Me, DELTA_HEIGHT_BOTTOM + DELTA_WIDTH_RIGHT
'    'Inside frame frmLog
'    gAutoPos.AddAssignment txtLog, frmLog, DELTA_HEIGHT_BOTTOM + DELTA_WIDTH_RIGHT
    'inside frame frmCheckDB
'    gAutoPos.AddAssignment txtDbLog, frmCheckDB, DELTA_HEIGHT_BOTTOM + DELTA_WIDTH_RIGHT
'    gAutoPos.AddAssignment cmdCheckNow, frmCheckDB, RELATIVE_POS_BOTTOM + RELATIVE_POS_RIGHT
'    gAutoPos.AddAssignment cmdDbCheck, frmCheckDB, RELATIVE_POS_BOTTOM + RELATIVE_POS_RIGHT
'    gAutoPos.AddAssignment cmdDBinfo, frmCheckDB, RELATIVE_POS_BOTTOM + RELATIVE_POS_RIGHT
'    gAutoPos.AddAssignment cmdDbDownload, frmCheckDB, RELATIVE_POS_BOTTOM + RELATIVE_POS_RIGHT
'    gAutoPos.AddAssignment txtDbURL, frmCheckDB, DELTA_WIDTH_RIGHT
'    gAutoPos.AddAssignment txtDbDestination, frmCheckDB, DELTA_WIDTH_RIGHT
'    gAutoPos.AddAssignment prgsBar, frmCheckDB, RELATIVE_POS_BOTTOM + DELTA_WIDTH_RIGHT
    'inside frame frmDownload
'    gAutoPos.AddAssignment txtDownloadedFile, frmDownload, DELTA_HEIGHT_BOTTOM + DELTA_WIDTH_RIGHT
'    gAutoPos.AddAssignment lblDownloadDate, frmDownload, RELATIVE_POS_RIGHT
    'Inside frame frmPage
'    gAutoPos.AddAssignment WebBrowser, frmPage, DELTA_HEIGHT_BOTTOM + DELTA_WIDTH_RIGHT
'    gAutoPos.AddAssignment cmdBrowse, frmPage, RELATIVE_POS_RIGHT
'    gAutoPos.AddAssignment cmdBrowseNext, frmPage, RELATIVE_POS_RIGHT
'    gAutoPos.AddAssignment cmdBrowsePrev, frmPage, RELATIVE_POS_RIGHT
'    gAutoPos.AddAssignment txtPageURL, frmPage, DELTA_WIDTH_RIGHT
'    'Inside frame frmRegistry
'    gAutoPos.AddAssignment txtRegistrySettings, frmRegistry, DELTA_HEIGHT_BOTTOM + DELTA_WIDTH_RIGHT
'    gAutoPos.AddAssignment cmdGetRegistry, frmRegistry, RELATIVE_POS_BOTTOM + RELATIVE_POS_RIGHT
'    gAutoPos.AddAssignment cmdEditReg, frmRegistry, RELATIVE_POS_BOTTOM + RELATIVE_POS_RIGHT
'    gAutoPos.AddAssignment lblRegistryPath, frmRegistry, RELATIVE_POS_BOTTOM + DELTA_WIDTH_RIGHT
    
    'Start timer
    startTimer
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    FormSys.MeQueryUnload Me, Cancel, UnloadMode
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmError
End Sub

Private Sub Form_Resize()
    FormSys.MeResize Me
    gAutoPos.RefreshPositions
End Sub

Private Sub FormSys_TIcon(F As Form)
    Me.Icon = F.Icon
    If frmError.Visible Then frmError.picError = F.Icon
End Sub

Private Sub FlashErrorIcon()
    FormSys.TrayIcon = "Flash_Error"
End Sub

Private Sub FlashAttentionIcon()
    FormSys.TrayIcon = "Flash_Attention"
End Sub

Private Sub FlashIconOff()
    FormSys.TrayIcon = "Default"
End Sub
Private Function PlaatselijkeSterrentijd(dattijd As Date)
' bereken juliaansedatum
Dim A As Integer, B As Integer, m As Integer, j As Integer
Dim D As Double, JD As Double, T As Double
Dim Theta As Double
Dim SiderealTime As Double
Dim nLongitude As Double
Dim JD_ZT As Double, JD_WT As Double

D = Day(dattijd) + TimeSerial(Hour(dattijd), Minute(dattijd), Second(dattijd))
If (Month(dattijd) > 2) Then
    j = Year(dattijd)
    m = Month(dattijd)
    
Else
    j = Year(dattijd) - 1
    m = Month(dattijd) + 12
End If

A = Int(j / 100)

If (Year(dattijd) < 1582) _
Or ((Year(dattijd) = 1582) And ((Month(dattijd) < 10) _
Or ((Month(dattijd) = 10) And (Day(dattijd <= 4))))) Then
    B = 0
Else
    B = 2 - A + Int(A / 4)
End If
JD = Int(365.25 * (j + 4716)) + Int(30.6001 * (m + 1)) + D + B - 1524.5
'correctie voor zomertijd/wintertijd
Call Zomertijd_Wintertijd(Year(dattijd), JD_ZT, JD_WT)
If JD >= JD_ZT And JD < JD_WT Then
    JD = JD - 2 / 24
Else
    JD = JD - 1 / 24
End If

'Bereken T (juliaanse eeuwen)
T = (JD - 2451545) / 36525
'bereken sid.time Greenwich
Theta = T * (360.98564736629 * 36525 + T * (0.000387933 - T / 38710000))
SiderealTime = modpi2((280.46061837 + Theta) * Pi / 180) * 12 / Pi
'Corrigeer voor de oosterlengte. Dit moet er afgetrokken worden
'wat inhoudt dat er ongeveer 20 min. opgeteld worden, want het is -5.08 oosterlengte
SiderealTime = SiderealTime - 4 * ReadDMS(txtLongitude) / 60
'en op scherm plaatsen
'stBar.Panels(2).Text = Format(SiderealTime / 24, "hh:mm:ss")
PlaatselijkeSterrentijd = SiderealTime / 24
End Function
'(*****************************************************************************)
'(* Name:    modpi2                                                           *)
'(* Type:    Function                                                         *)
'(* Purpose: reduce an angle to the interval (0, 2pi).                        *)
'(*****************************************************************************)

Function modpi2(x As Double) As Double
x = x - Int(x / Pi2) * Pi2
If (x < 0) Then
  x = x + Pi2
End If
modpi2 = x
End Function

Function ReadDMS(s As String) As Double
Dim D As Long, m As Long, sign As Long
Dim Angle As Double, ss   As Double
    Angle = Val(s)
    If (Angle < 0) Then
        sign = -1
        Angle = -Angle
    Else
        sign = 1
    End If
    D = Int(Angle)
    Angle = (Angle - D) * 100
    m = Int(Angle)
    ss = (Angle - m) * 100 + 0.00001
               '{ Otherwise we might get 59.999... seconds }
    ReadDMS = sign * (D + m / 60# + ss / 3600#)
End Function

