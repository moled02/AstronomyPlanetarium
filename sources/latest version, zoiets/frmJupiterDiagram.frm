VERSION 5.00
Begin VB.Form frmJupiterDiagram 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Jupiter diagram"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   FillColor       =   &H00FFFFFF&
   HasDC           =   0   'False
   Icon            =   "frmJupiterDiagram.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   403
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   605
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picDiagram 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FillColor       =   &H00FFFFFF&
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
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   7
      Left            =   600
      TabIndex        =   8
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   6
      Left            =   600
      TabIndex        =   7
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   600
      TabIndex        =   6
      Top             =   3930
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   600
      TabIndex        =   5
      Top             =   3180
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   600
      TabIndex        =   4
      Top             =   2430
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   600
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   930
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   180
      Width           =   1575
   End
End
Attribute VB_Name = "frmJupiterDiagram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MERGEPAINT = &HBB0226
Private Const SRCAND = &H8800C6
Private Const SRCCOPY = &HCC0020

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Type Tgeg
    x1 As Double
    x2 As Double
    x3 As Double
    x4 As Double
End Type


Private Sub JupiterPhysEphLo(ByVal JD As Double, ByRef DD As Double, ByRef B As Double, ByRef Psi As Double, ByRef DeclS As Double, ByRef DeclE As Double)

Dim D As Double
Dim v As Double, M As Double, n As Double, j As Double, A As Double, k As Double, lambda As Double
Dim Re As Double, Rj     As Double '{Radius vector of Earth, Jupiter}
Dim Rej     As Double ' {distance Earth-Jupiter}
  
  D = JD - 2451545#
  v = (172.74 + 0.00111588 * D) * DToR
  M = (357.529 + 0.9856003 * D) * DToR
  n = (20.02 + 0.0830853 * D + 0.329 * Sin(v)) * DToR
  j = (66.115 + 0.9025179 * D - 0.329 * Sin(v)) * DToR
  A = (1.915 * Sin(M) + 0.02 * Sin(2 * M)) * DToR
  B = (5.555 * Sin(n) - 0.168 * Sin(2 * n)) * DToR
  k = j + A - B
  Re = 1.00014 - 0.1671 * Cos(M) - 0.00014 * Cos(2 * M)
  Rj = 5.20872 - 0.25208 * Cos(M) - 0.00611 * Cos(2 * n)
  Rej = Sqr(Re * Re + Rj * Rj - 2 * Re * Rj * Cos(k))
  Psi = asin(Sin(k) * Re / Rej)
  DD = D - Rej / 173
  lambda = (34.35 + 0.083091 * D + 0.329 * Sin(v)) * DToR
  DeclS = (3.12 * Sin(lambda + 42.8 * DToR)) * DToR
  DeclE = DeclS - (2.22 * Sin(Psi) * Cos(lambda + 22 * DToR)) * DToR
  DeclE = DeclE - (1.3 * (Rj - Rej) / Rej * Sin(lambda - 100.5 * DToR)) * DToR
End Sub

Private Sub JupiterGeg(ByVal JD As Double, ByRef u As Tgeg, ByRef r As Tgeg, ByRef X As Tgeg, ByRef Y As Tgeg)
Dim g As Double, H As Double, DD As Double, B As Double, Psi As Double, yScalefactor As Double, DeclE As Double
Dim i As Long

Call JupiterPhysEphLo(JD, DD, B, Psi, yScalefactor, DeclE)

    u.x1 = (163.8067 + 203.4058643 * DD) * DToR + Psi - B
    u.x2 = (358.4108 + 101.2916334 * DD) * DToR + Psi - B
    u.x3 = (5.7129 + 50.2345179 * DD) * DToR + Psi - B
    u.x4 = (224.8151 + 21.4879801 * DD) * DToR + Psi - B

'    {Principal perturbations}
    g = (331.18 + 50.310488 * DD) * DToR
    H = (87.4 + 21.569231 * DD) * DToR
    u.x1 = u.x1 + 0.473 * Sin(2 * (u.x1 - u.x2)) * DToR
    u.x2 = u.x2 + 1.065 * Sin(2 * (u.x2 - u.x3)) * DToR
    u.x3 = u.x3 + 0.165 * Sin(g) * DToR
    u.x4 = u.x4 + 0.841 * Sin(H) * DToR
    r.x1 = 5.9073 - 0.0244 * Cos(2 * (u.x1 - u.x2))
    r.x2 = 9.3991 - 0.0882 * Cos(2 * (u.x2 - u.x3))
    r.x3 = 14.9924 - 0.0216 * Cos(g)
    r.x4 = 26.3699 - 0.1935 * Cos(H)
    yScalefactor = Sin(DeclE)
    X.x1 = r.x1 * Sin(u.x1)
    X.x2 = r.x2 * Sin(u.x2)
    X.x3 = r.x3 * Sin(u.x3)
    X.x4 = r.x4 * Sin(u.x4)
    Y.x1 = r.x1 * Cos(u.x1)
    Y.x2 = r.x2 * Cos(u.x2)
    Y.x3 = r.x3 * Cos(u.x3)
    Y.x4 = r.x4 * Cos(u.x4)

End Sub

Sub teken(ByVal JD As Double)

Dim i As Long
Dim maand As String, dag As String
Dim k As Double
Dim u As Tgeg, r As Tgeg, X As Tgeg, Y As Tgeg
Dim color As Long
Dim JD_ZT As Double, JD_WT As Double
Dim ddate As tDatum

Dim TimeZone As Double

Cls
For i = -3 To 4
  ddate = JDNaarKalender(JD + i)
  maand = Format(ddate.MM, "00")
  dag = Format(ddate.DD, "00")
  Label1(i + 3).Caption = dag + "-" + maand + "-" + Format(ddate.jj, "0000")
Next
Me.Refresh
TimeZone = TijdCorrectie(JD, JD_ZT, JD_WT)
Font.Name = "Arial"
Font.size = 10
Font.Bold = True
  Call TekenRechthoek(0, 0, 600, 400)
  Call TekenRechthoek(145, 19, 550, 370)
  Call TekenRechthoek(146, 20, 549, 369)
  For i = -3 To 3
    Line (145, 70 + 50 * (i + 2))-(339, 70 + 50 * (i + 2))
    Line (351, 70 + 50 * (i + 2))-(550, 70 + 50 * (i + 2))
    If i < 4 Then
      DrawStyle = vbDot
      Line (145, 70 + Int(50 * (i + 2.5)))-(339, 70 + Int(50 * (i + 2.5)))
      Line (351, 70 + Int(50 * (i + 2.5)))-(550, 70 + Int(50 * (i + 2.5)))
      DrawStyle = vbSolid
    End If
  Next

  Line (339, 20)-(339, 370)
  Line (351, 20)-(351, 370)
  For i = -33 To 33
    If i Mod 5 = 0 Then
      Line (345 + 6 * i, 360)-(345 + 6 * i, 370)
    Else
      Line (345 + 6 * i, 365)-(345 + 6 * i, 370)
    End If

    If i Mod 5 = 0 Then
      Line (345 + 6 * i, 20)-(345 + 6 * i, 30)
    Else
      Line (345 + 6 * i, 20)-(345 + 6 * i, 25)
    End If
  Next
  k = -3#
  Do While k < 4#
    Call JupiterGeg(JD + k, u, r, X, Y)
    u.x1 = modpi2(u.x1 + Pi)
    u.x2 = modpi2(u.x2 + Pi)
    u.x3 = modpi2(u.x3 + Pi)
    u.x4 = modpi2(u.x4 + Pi)
    If Not (((u.x1 < Pi / 4) Or (u.x1 > 1.75 * Pi)) And (X.x1 * X.x1 < 1)) Then
       PSet (345 - Int(6 * X.x1), 70 + Int(50 * (k + 2))), vbBlack
    End If
    If Not (((u.x2 < Pi / 4) Or (u.x2 > 1.75 * Pi)) And (X.x2 * X.x2 < 1)) Then
       PSet (345 - Int(6 * X.x2), 70 + Int(50 * (k + 2))), vbBlue
    End If
    If Not (((u.x3 < Pi / 4) Or (u.x3 > 1.75 * Pi)) And (X.x3 * X.x3 < 1)) Then
       PSet (345 - Int(6 * X.x3), 70 + Int(50 * (k + 2))), vbGreen
    End If
    If Not (((u.x4 < Pi / 4) Or (u.x4 > 1.75 * Pi)) And (X.x4 * X.x4 < 1)) Then
       PSet (345 - Int(6 * X.x4), 70 + Int(50 * (k + 2))), vbRed
    End If
    k = k + 0.01
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

Dim objspecialfolder As New clsSpecialFolder

cPad = objspecialfolder.TemporaryFolder


  'cPad = App.Path
  
  #If FRANS Then
    cPad = Trim(InputBox("Donnée chemin d'accès", "Fabrique graphique oscillation de satellites de Jupiter", cPad))
  #Else
    cPad = Trim(InputBox("Give exportpath", "Diagram of the satellites of Jupiter", cPad))
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
  Do While JD <= jde + 0.3
    DoEvents
    ddate = JDNaarKalender(JD)
    weeknr = WeekOfYear(ddate)
    Call BepaalZT_WT(ddate.jj, JD_ZT, JD_WT)
    Call teken(JD + 3#)
   
'    BitBlt Picture1.hDC, 0, 0, 596, 396, _
'    Me.hDC, 0, 0, SRCCOPY
'    Picture1.Refresh
    BitBlt picDiagram.hdc, _
    0, 0, 5000, 5000, _
    Me.hdc, 0, 0, SRCCOPY
'    BitBlt picDiagram.hDC, _
    0, 0, picDiagram.ScaleWidth, picDiagram.ScaleHeight, _
    Me.hDC, 0, 0, SRCCOPY
    picDiagram.Refresh
    Call SavePicture(picDiagram.Image, cPad + "\J" + Format(weeknr, "000000") + ".bmp")
    JD = JD + 7
  Loop
  Unload Me
End Sub


Private Sub Form_Load()
Dim i As Long
#If FRANS Then
    Me.Caption = "Graphique Jupiter"
#End If
For i = 0 To 7: Label1(i).Caption = "": Next
End Sub

