VERSION 5.00
Begin VB.Form frmSaturnus 
   BackColor       =   &H8000000A&
   Caption         =   "Moons of Saturn"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14010
   Icon            =   "frmSaturnus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   612
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   934
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9840
      Top             =   240
   End
   Begin VB.CommandButton Command3 
      Caption         =   "S&tep"
      Height          =   375
      Left            =   7200
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Start"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   8145
      Left            =   240
      ScaleHeight     =   539
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   907
      TabIndex        =   2
      Top             =   960
      Width           =   13665
   End
   Begin VB.PictureBox picHidden 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   8085
      Left            =   240
      ScaleHeight     =   539
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   907
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   13605
   End
End
Attribute VB_Name = "frmSaturnus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Schaal As Double
Private JD As Double
Private Const MERGEPAINT = &HBB0226
Private Const SRCAND = &H8800C6
Private Const SRCCOPY = &HCC0020

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private lDoorgaan As Boolean
Private nStap As Long

Sub teken(ByVal JD As Double)
Dim a1 As Double
Dim a2 As Double
Dim k As Double
Dim nFactor As Double
Dim u(6) As Double
Dim r(6) As Double
Dim x(6) As Double
Dim y(6) As Double
Dim SaturnB As Double
Dim Xstart As Double
Dim Ystart As Double
Xstart = Int(picHidden.ScaleWidth / 2)
Ystart = Int(picHidden.ScaleHeight / 2)

picHidden.Cls
picHidden.Line (0, Ystart)-(Xstart * 2, Ystart)
picHidden.Line (Xstart, 0)-(Xstart, Ystart * 2)
  k = 0
  nFactor = Int(picHidden.ScaleWidth / 350) * Schaal
  Call SaturnusGeg(JD + k, u(), r(), x(), y(), SaturnB)
  For I = 1 To 6
      x(I) = -x(I)
      'y(i) = -y(i)
      u(I) = u(I) - 360 * Int(u(I) / 360) - 180 't.b.v. controle voor of achter Saturnus
  Next
    
    ' eerste loop tekenen saturnusmaantjes: eerst de maantjes die achter de planeet staan
    For I = 1 To 6
     If (I = 1) Or (I = 2) Or (I = 3) Or (I = 4) Or (I = 5) Or (I = 6) Then
        If Abs(u(I)) > 90 Then
            If Schaal > 1.5 Then
                picHidden.Circle (Xstart - Int(nFactor * x(I)), Ystart + Int(nFactor * (y(I)))), 2
            ElseIf Schaal > 0.7 Then
                picHidden.Circle (Xstart - Int(nFactor * x(I)), Ystart + Int(nFactor * (y(I)))), 1
            Else
                picHidden.PSet (Xstart - Int(nFactor * x(I)), Ystart + Int(nFactor * (y(I)))), 0
            End If
        End If
     End If
    Next
    'DoEvents

' Het tekenen is in 2 delen gesplitst as  de onderkant en de bovenkant
'  Dit om het mogelijk te maken ringen voor en achter de planeet te tekenen}
If Sin(SaturnB) <= 0 Then
   a1 = Int(nFactor * SaturnRingData.aAxis / 2)
   a2 = Int(nFactor * SaturnRingData.bAxis / 2)
   picHidden.Circle (Xstart, Ystart), a1, RGB(255, 0, 0), 0, Pi, a2 / a1
   a1 = Int(nFactor * SaturnRingData.ioaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.iobAxis / 2)
   picHidden.Circle (Xstart, Ystart), a1, 0, 0, Pi, a2 / a1
   a1 = Int(nFactor * SaturnRingData.oiaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.oibAxis / 2)
   picHidden.Circle (Xstart, Ystart), a1, 0, 0, Pi, a2 / a1

   a1 = Int(nFactor * SaturnRingData.iiaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.iibAxis / 2)
   picHidden.Circle (Xstart, Ystart), a1, RGB(255, 0, 0), 0, Pi, a2 / a1
   
   a1 = Int(nFactor * SaturnRingData.idaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.idbAxis / 2)
   picHidden.Circle (Xstart, Ystart), a1, RGB(100, 0, 0), 0, Pi, a2 / a1
Else
   a1 = Int(nFactor * SaturnRingData.aAxis / 2)
   a2 = Int(nFactor * SaturnRingData.bAxis / 2)
   picHidden.Circle (Xstart, Ystart), a1, RGB(255, 0, 0), Pi, 2 * Pi, a2 / a1

   a1 = Int(nFactor * SaturnRingData.ioaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.iobAxis / 2)
   picHidden.Circle (Xstart, Ystart), a1, 0, Pi, 2 * Pi, a2 / a1

   a1 = Int(nFactor * SaturnRingData.oiaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.oibAxis / 2)
   picHidden.Circle (Xstart, Ystart), a1, 0, Pi, 2 * Pi, a2 / a1

   a1 = Int(nFactor * SaturnRingData.iiaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.iibAxis / 2)
   picHidden.Circle (Xstart, Ystart), a1, RGB(255, 0, 0), Pi, 2 * Pi, a2 / a1
   
   a1 = Int(nFactor * SaturnRingData.idaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.idbAxis / 2)
   picHidden.Circle (Xstart, Ystart), a1, RGB(100, 0, 0), Pi, 2 * Pi, a2 / a1
End If
r1 = Int(nFactor * Semidiameter)
r2 = Int(nFactor * PolarSemiDiameter)
picHidden.FillColor = RGB(200, 200, 200)
picHidden.FillStyle = 0
picHidden.Circle (Xstart, Ystart), r1, 0, , , r2 / r1
picHidden.Circle (Xstart, Ystart), r1, 0, , , r2 / r1
picHidden.FillStyle = 0

If Sin(SaturnRingData.B) > 0 Then
   a1 = Int(nFactor * SaturnRingData.aAxis / 2)
   a2 = Int(nFactor * SaturnRingData.bAxis / 2)
   picHidden.Circle (Xstart, Ystart), a1, RGB(255, 0, 0), 0, Pi, a2 / a1

   a1 = Int(nFactor * SaturnRingData.ioaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.iobAxis / 2)
   picHidden.Circle (Xstart, Ystart), a1, 0, 0, Pi, a2 / a1

   a1 = Int(nFactor * SaturnRingData.oiaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.oibAxis / 2)
   picHidden.Circle (Xstart, Ystart), a1, 0, 0, Pi, a2 / a1

   a1 = Int(nFactor * SaturnRingData.iiaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.iibAxis / 2)
   picHidden.Circle (Xstart, Ystart), a1, RGB(255, 0, 0), 0, Pi, a2 / a1
   
   a1 = Int(nFactor * SaturnRingData.idaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.idbAxis / 2)
   picHidden.Circle (Xstart, Ystart), a1, RGB(100, 0, 0), 0, Pi, a2 / a1
Else
   a1 = Int(nFactor * SaturnRingData.aAxis / 2)
   a2 = Int(nFactor * SaturnRingData.bAxis / 2)
   picHidden.Circle (Xstart, Ystart), a1, RGB(255, 0, 0), Pi, 2 * Pi, a2 / a1

   a1 = Int(nFactor * SaturnRingData.ioaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.iobAxis / 2)
   picHidden.Circle (Xstart, Ystart), a1, 0, Pi, 2 * Pi, a2 / a1

   a1 = Int(nFactor * SaturnRingData.oiaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.oibAxis / 2)
   picHidden.Circle (Xstart, Ystart), a1, 0, Pi, 2 * Pi, a2 / a1

   a1 = Int(nFactor * SaturnRingData.iiaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.iibAxis / 2)
   picHidden.Circle (Xstart, Ystart), a1, RGB(255, 0, 0), Pi, 2 * Pi, a2 / a1
   
   a1 = Int(nFactor * SaturnRingData.idaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.idbAxis / 2)
   picHidden.Circle (Xstart, Ystart), a1, RGB(100, 0, 0), Pi, 2 * Pi, a2 / a1
End If
    
' en nu de maantjes die voor de planeet langsgaan
    For I = 1 To 6
     If (I = 1) Or (I = 2) Or (I = 3) Or (I = 4) Or (I = 5) Or (I = 6) Then
        If Abs(u(I)) <= 90 Then
            If Schaal > 1.5 Then
                picHidden.Circle (Xstart - Int(nFactor * x(I)), Ystart + Int(nFactor * (y(I)))), 2
            ElseIf Schaal > 0.7 Then
                picHidden.Circle (Xstart - Int(nFactor * x(I)), Ystart + Int(nFactor * (y(I)))), 1
            Else
                picHidden.PSet (Xstart - Int(nFactor * x(I)), Ystart + Int(nFactor * (y(I)))), 0
            End If
        End If
     End If
    Next
    'DoEvents
End Sub

Private Sub Command2_Click()
Dim dat As tDatum
dat.jj = frmPlanets.Year
dat.mm = frmPlanets.MonthSelect.ListIndex + 1
dat.DD = frmPlanets.DaySelect
tt = (frmPlanets.Hrs + frmPlanets.Min / 60 + frmPlanets.Sec / 3600) / 24
dat.DD = dat.DD + tt
JD = KalenderNaarJD(dat)
Schaal = 1
nStap = 1
Timer1.Enabled = True
End Sub

Private Sub Form_Activate()
Command2_Click
End Sub

Private Sub Form_Load()
#If FRANS Then
    Me.Caption = "Satellites de Saturne"
    Me.Command2.Caption = "Départ"
    Me.Command3.Caption = "Pas"
#End If
End Sub

Private Sub Timer1_Timer()
Dim dDatum As tDatum
Dim JD_ZT As Double
Dim JD_WT As Double
Dim JD_Calc As Double
   
    DoEvents
    JD = JD + nStap / 86400 * (Timer1.Interval / 1000)
    dDatum = JDNaarKalender(JD)
    Call Zomertijd_Wintertijd(dDatum.jj, JD_ZT, JD_WT)
    dDatum = JDNaarKalender(JD)
    frmSaturnus.Text1.Text = Format(Int(dDatum.DD), "00") + "-" + Format(dDatum.mm, "00") + "-" + Format(dDatum.jj, "00") + " " + Format(dDatum.DD - Int(dDatum.DD), "hh:mm:ss")
    Call Zomertijd_Wintertijd(dDatum.jj, JD_ZT, JD_WT)
    JD_Calc = JD + TijdCorrectie(JD, JD_ZT, JD_WT) + ApproxDeltaT(JDToT(JD)) / 86400
    Call teken(JD_Calc)
    BitBlt picCanvas.hDC, _
    0, 0, picCanvas.ScaleWidth, picCanvas.ScaleHeight, _
    picHidden.hDC, 0, 0, SRCCOPY
    picCanvas.Refresh
End Sub

Private Sub Command3_Click()
#If FRANS Then
    nStap = Val(InputBox("Donner par in secondes: ", "Donnée pas"))
#Else
    nStap = Val(InputBox("Give step in seconds: ", "Input step"))
#End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
lDoorgaan = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Static nTmpStap As Long
If KeyAscii = 62 Then '>
    nStap = nStap + 1
ElseIf KeyAscii = 60 Then '<
    nStap = nStap - 1
ElseIf KeyAscii = 43 Then '+
    picHidden.Cls
    If Schaal < 10 Then Schaal = Schaal + 0.1
ElseIf KeyAscii = 45 Then '-
    picHidden.Cls
    If Schaal > 0.6 Then Schaal = Schaal - 0.1
ElseIf KeyAscii = 46 Then '.
    nTmpStap = nStap
    nStap = 0
ElseIf KeyAscii = 44 Then ',
    nStap = nTmpStap
End If
End Sub

