VERSION 5.00
Begin VB.Form frmSaturnusDiagram 
   BackColor       =   &H80000009&
   Caption         =   "Saturn diagram"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   HasDC           =   0   'False
   Icon            =   "frmSaturnusDiagram.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   403
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   605
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picDiagram 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      Height          =   6000
      Left            =   0
      ScaleHeight     =   396
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   596
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   9000
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Label1"
      Height          =   375
      Index           =   7
      Left            =   600
      TabIndex        =   8
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Label1"
      Height          =   375
      Index           =   6
      Left            =   600
      TabIndex        =   7
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Label1"
      Height          =   375
      Index           =   5
      Left            =   600
      TabIndex        =   6
      Top             =   3930
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Label1"
      Height          =   375
      Index           =   4
      Left            =   600
      TabIndex        =   5
      Top             =   3180
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Label1"
      Height          =   375
      Index           =   3
      Left            =   600
      TabIndex        =   4
      Top             =   2430
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Label1"
      Height          =   375
      Index           =   2
      Left            =   600
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Label1"
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   930
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Label1"
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   180
      Width           =   1575
   End
End
Attribute VB_Name = "frmSaturnusDiagram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MERGEPAINT = &HBB0226
Private Const SRCAND = &H8800C6
Private Const SRCCOPY = &HCC0020

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Type Tgeg
    x1 As Double
    x2 As Double
    x3 As Double
    x4 As Double
End Type
Private blnDoorgaan As Boolean

Sub teken(ByVal JD As Double)

Dim I As Long
Dim maand As String, dag As String
Dim k As Double
Dim u(6) As Double
Dim r(6) As Double
Dim x(6) As Double
Dim Y(6) As Double
Dim color As Long
Dim JD_ZT As Double, JD_WT As Double
Dim ddate As tDatum
Dim SaturnB As Double

Dim TimeZone As Double

Cls
For I = -3 To 4
  ddate = JDNaarKalender(JD + I)
  maand = Format(ddate.mm, "00")
  dag = Format(ddate.DD, "00")
  Label1(I + 3).Caption = dag + "-" + maand + "-" + Format(ddate.jj, "0000")
Next
Me.Refresh
TimeZone = TijdCorrectie(JD, JD_ZT, JD_WT)
Font.Name = "Arial"
Font.Size = 10
Font.Bold = True
  Call TekenRechthoek(0, 0, 600, 400)
  Call TekenRechthoek(145, 19, 550, 370)
  Call TekenRechthoek(146, 20, 549, 369)
  For I = -3 To 3
    Line (145, 70 + 50 * (I + 2))-(339, 70 + 50 * (I + 2))
    Line (351, 70 + 50 * (I + 2))-(550, 70 + 50 * (I + 2))
    If I < 4 Then
      DrawStyle = vbDot
      Line (145, 70 + Int(50 * (I + 2.5)))-(339, 70 + Int(50 * (I + 2.5)))
      Line (351, 70 + Int(50 * (I + 2.5)))-(550, 70 + Int(50 * (I + 2.5)))
      DrawStyle = vbSolid
    End If
  Next

  Line (339, 20)-(339, 370)
  Line (351, 20)-(351, 370)
  For I = -33 To 33
    If I Mod 5 = 0 Then
      Line (345 + 6 * I, 360)-(345 + 6 * I, 370)
    Else
      Line (345 + 6 * I, 365)-(345 + 6 * I, 370)
    End If

    If I Mod 5 = 0 Then
      Line (345 + 6 * I, 20)-(345 + 6 * I, 30)
    Else
      Line (345 + 6 * I, 20)-(345 + 6 * I, 25)
    End If
  Next
  k = -3#
  Do While k < 4# And blnDoorgaan
    DoEvents
    Call SaturnusGeg(JD + k, u(), r(), x(), Y(), SaturnB)
    For I = 1 To 6
      x(I) = x(I) / 7.5  'is niet helemaal juist: diam. Sat/2
      Y(I) = Y(I) / 7.5
'      If Not ((x(i) * x(i) + y(i) * y(i) < 1)) Then
        x(I) = -x(I)
        If I = 1 Then PSet (345 - Int(6 * x(1)), 70 + Int(50 * (k + 2))), vbBlack
        If I = 2 Then PSet (345 - Int(6 * x(2)), 70 + Int(50 * (k + 2))), vbGreen
        If I = 3 Then PSet (345 - Int(6 * x(3)), 70 + Int(50 * (k + 2))), vbBlue
        If I = 4 Then PSet (345 - Int(6 * x(4)), 70 + Int(50 * (k + 2))), vbRed
        If I = 5 Then PSet (345 - Int(6 * x(5)), 70 + Int(50 * (k + 2))), vbCyan
        If I = 6 Then PSet (345 - Int(6 * x(6)), 70 + Int(50 * (k + 2))), vbMagenta
'      Else
'        MsgBox "ja"
'      End If
    Next
    k = k + 0.005
  Loop
End Sub

Private Sub TekenRechthoek(x1, y1, x2, y2)
Line (x1, y1)-(x2, y1)
Line (x2, y1)-(x2, y2)
Line (x2, y2)-(x1, y2)
Line (x1, y2)-(x1, y1)
End Sub

Private Sub Form_Activate()

Dim ddate As tDatum
Dim JD  As Double
Dim weeknr As Long
Dim JD_ZT As Double, JD_WT As Double
Dim cPad As String

'  cPad = App.Path
  
Dim objspecialfolder As New clsSpecialFolder

cPad = objspecialfolder.TemporaryFolder
  
  #If FRANS Then
    cPad = Trim(InputBox("Donnée chemin d'accès", "Fabrique graphique oscillation de satellites de Saturne", cPad))
  #Else
    cPad = Trim(InputBox("Give exportpath", "Diagram of the satellites of Saturn", cPad))
  #End If
  
  If Right(cPad, 1) = "\" Then
    cPad = Left(cPad, Len(cPad) - 1)
  End If
  ddate.jj = frmPlanets.Year
  weeknr = ddate.jj * 100 + 1
  Call WeekDate(weeknr, ddate)
  JD = KalenderNaarJD(ddate)
  ddate.jj = frmPlanets.Year + 1
  weeknr = ddate.jj * 100 + 1
  Call WeekDate(weeknr, ddate)
  jde = KalenderNaarJD(ddate)
  blnDoorgaan = True
  Do While JD <= jde + 0.3 And blnDoorgaan
    DoEvents
    ddate = JDNaarKalender(JD)
    weeknr = WeekOfYear(ddate)
    Call BepaalZT_WT(ddate.jj, JD_ZT, JD_WT)
    Call teken(JD + 3#)
   
'    BitBlt Picture1.hDC, 0, 0, 596, 396, _
'    Me.hDC, 0, 0, SRCCOPY
'    Picture1.Refresh

    BitBlt picDiagram.hDC, _
    0, 0, 5000, 5000, _
    Me.hDC, 0, 0, SRCCOPY
'    BitBlt picDiagram.hDC, _
    0, 0, picDiagram.ScaleWidth, picDiagram.ScaleHeight, _
    Me.hDC, 0, 0, SRCCOPY
    picDiagram.Refresh
    Call SavePicture(picDiagram.Image, cPad + "\S" + Format(weeknr, "000000") + ".bmp")
    JD = JD + 7
  Loop
  Unload Me
End Sub


Private Sub Form_Load()
Dim I As Long
#If FRANS Then
    Me.Caption = "Graphique Saturne"
#End If
For I = 0 To 7: Label1(I).Caption = "": Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
blnDoorgaan = False
End Sub

