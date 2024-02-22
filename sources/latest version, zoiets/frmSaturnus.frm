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
Private schaal As Double
Private JD As Double
Private Const MERGEPAINT = &HBB0226
Private Const SRCAND = &H8800C6
Private Const SRCCOPY = &HCC0020

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private lDoorgaan As Boolean
Private nStap As Long

Public Sub tekenSaturnusRingsMoons(picHiddenSat As PictureBox, ByVal JD As Double, schaal As Double, Optional withMoons As Boolean = True)
Dim a1 As Double
Dim a2 As Double
Dim k As Double
Dim nFactor As Double
Dim u(6) As Double
Dim r(6) As Double
Dim x(6) As Double
Dim Y(6) As Double
Dim SaturnB As Double
Dim Xstart As Double
Dim Ystart As Double
Xstart = Int(picHiddenSat.ScaleWidth / 2)
Ystart = Int(picHiddenSat.ScaleHeight / 2)

picHiddenSat.Cls
picHiddenSat.Line (0, Ystart)-(Xstart * 2, Ystart)
picHiddenSat.Line (Xstart, 0)-(Xstart, Ystart * 2)
  k = 0
  nFactor = Int(picHiddenSat.ScaleWidth / 350) * schaal
  Call SaturnusGeg(JD + k, u(), r(), x(), Y(), SaturnB)
  If withMoons Then
    For i = 1 To 6
        x(i) = -x(i)
        'y(i) = -y(i)
        u(i) = u(i) - 360 * Int(u(i) / 360) - 180 't.b.v. controle voor of achter Saturnus
    Next
    ' eerste loop tekenen saturnusmaantjes: eerst de maantjes die achter de planeet staan
    For i = 1 To 6
     If (i = 1) Or (i = 2) Or (i = 3) Or (i = 4) Or (i = 5) Or (i = 6) Then
        If Abs(u(i)) > 90 Then
            If schaal > 1.5 Then
                picHiddenSat.Circle (Xstart - Int(nFactor * x(i)), Ystart + Int(nFactor * (Y(i)))), 2
            ElseIf schaal > 0.7 Then
                picHiddenSat.Circle (Xstart - Int(nFactor * x(i)), Ystart + Int(nFactor * (Y(i)))), 1
            Else
                picHiddenSat.PSet (Xstart - Int(nFactor * x(i)), Ystart + Int(nFactor * (Y(i)))), 0
            End If
        End If
     End If
    Next
    'DoEvents
End If
' Het tekenen is in 2 delen gesplitst as  de onderkant en de bovenkant
'  Dit om het mogelijk te maken ringen voor en achter de planeet te tekenen}
If Sin(SaturnB) <= 0 Then
   a1 = Int(nFactor * SaturnRingData.aAxis / 2)
   a2 = Int(nFactor * SaturnRingData.bAxis / 2)
   picHiddenSat.Circle (Xstart, Ystart), a1, RGB(255, 0, 0), 0, Pi, a2 / a1
   a1 = Int(nFactor * SaturnRingData.ioaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.iobAxis / 2)
   picHiddenSat.Circle (Xstart, Ystart), a1, 0, 0, Pi, a2 / a1
   a1 = Int(nFactor * SaturnRingData.oiaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.oibAxis / 2)
   picHiddenSat.Circle (Xstart, Ystart), a1, 0, 0, Pi, a2 / a1

   a1 = Int(nFactor * SaturnRingData.iiaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.iibAxis / 2)
   picHiddenSat.Circle (Xstart, Ystart), a1, RGB(255, 0, 0), 0, Pi, a2 / a1
   
   a1 = Int(nFactor * SaturnRingData.idaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.idbAxis / 2)
   picHiddenSat.Circle (Xstart, Ystart), a1, RGB(100, 0, 0), 0, Pi, a2 / a1
Else
   a1 = Int(nFactor * SaturnRingData.aAxis / 2)
   a2 = Int(nFactor * SaturnRingData.bAxis / 2)
   picHiddenSat.Circle (Xstart, Ystart), a1, RGB(255, 0, 0), Pi, 2 * Pi, a2 / a1

   a1 = Int(nFactor * SaturnRingData.ioaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.iobAxis / 2)
   picHiddenSat.Circle (Xstart, Ystart), a1, 0, Pi, 2 * Pi, a2 / a1

   a1 = Int(nFactor * SaturnRingData.oiaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.oibAxis / 2)
   picHiddenSat.Circle (Xstart, Ystart), a1, 0, Pi, 2 * Pi, a2 / a1

   a1 = Int(nFactor * SaturnRingData.iiaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.iibAxis / 2)
   picHiddenSat.Circle (Xstart, Ystart), a1, RGB(255, 0, 0), Pi, 2 * Pi, a2 / a1
   
   a1 = Int(nFactor * SaturnRingData.idaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.idbAxis / 2)
   picHiddenSat.Circle (Xstart, Ystart), a1, RGB(100, 0, 0), Pi, 2 * Pi, a2 / a1
End If
r1 = Int(nFactor * Semidiameter)
r2 = Int(nFactor * PolarSemiDiameter)
picHiddenSat.FillColor = RGB(200, 200, 200)
picHiddenSat.FillStyle = 0
picHiddenSat.Circle (Xstart, Ystart), r1, 0, , , r2 / r1
picHiddenSat.Circle (Xstart, Ystart), r1, 0, , , r2 / r1
picHiddenSat.FillStyle = 0

If Sin(SaturnRingData.B) > 0 Then
   a1 = Int(nFactor * SaturnRingData.aAxis / 2)
   a2 = Int(nFactor * SaturnRingData.bAxis / 2)
   picHiddenSat.Circle (Xstart, Ystart), a1, RGB(255, 0, 0), 0, Pi, a2 / a1

   a1 = Int(nFactor * SaturnRingData.ioaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.iobAxis / 2)
   picHiddenSat.Circle (Xstart, Ystart), a1, 0, 0, Pi, a2 / a1

   a1 = Int(nFactor * SaturnRingData.oiaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.oibAxis / 2)
   picHiddenSat.Circle (Xstart, Ystart), a1, 0, 0, Pi, a2 / a1

   a1 = Int(nFactor * SaturnRingData.iiaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.iibAxis / 2)
   picHiddenSat.Circle (Xstart, Ystart), a1, RGB(255, 0, 0), 0, Pi, a2 / a1
   
   a1 = Int(nFactor * SaturnRingData.idaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.idbAxis / 2)
   picHiddenSat.Circle (Xstart, Ystart), a1, RGB(100, 0, 0), 0, Pi, a2 / a1
Else
   a1 = Int(nFactor * SaturnRingData.aAxis / 2)
   a2 = Int(nFactor * SaturnRingData.bAxis / 2)
   picHiddenSat.Circle (Xstart, Ystart), a1, RGB(255, 0, 0), Pi, 2 * Pi, a2 / a1

   a1 = Int(nFactor * SaturnRingData.ioaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.iobAxis / 2)
   picHiddenSat.Circle (Xstart, Ystart), a1, 0, Pi, 2 * Pi, a2 / a1

   a1 = Int(nFactor * SaturnRingData.oiaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.oibAxis / 2)
   picHiddenSat.Circle (Xstart, Ystart), a1, 0, Pi, 2 * Pi, a2 / a1

   a1 = Int(nFactor * SaturnRingData.iiaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.iibAxis / 2)
   picHiddenSat.Circle (Xstart, Ystart), a1, RGB(255, 0, 0), Pi, 2 * Pi, a2 / a1
   
   a1 = Int(nFactor * SaturnRingData.idaAxis / 2)
   a2 = Int(nFactor * SaturnRingData.idbAxis / 2)
   picHiddenSat.Circle (Xstart, Ystart), a1, RGB(100, 0, 0), Pi, 2 * Pi, a2 / a1
End If
    
  If withMoons Then
' en nu de maantjes die voor de planeet langsgaan
    For i = 1 To 6
     If (i = 1) Or (i = 2) Or (i = 3) Or (i = 4) Or (i = 5) Or (i = 6) Then
        If Abs(u(i)) <= 90 Then
            If schaal > 1.5 Then
                picHiddenSat.Circle (Xstart - Int(nFactor * x(i)), Ystart + Int(nFactor * (Y(i)))), 2
            ElseIf schaal > 0.7 Then
                picHiddenSat.Circle (Xstart - Int(nFactor * x(i)), Ystart + Int(nFactor * (Y(i)))), 1
            Else
                picHiddenSat.PSet (Xstart - Int(nFactor * x(i)), Ystart + Int(nFactor * (Y(i)))), 0
            End If
        End If
     End If
    Next
  End If
    'DoEvents
End Sub

Private Sub Command2_Click()
Dim dat As tDatum
dat.jj = frmPlanets.Year
dat.MM = frmPlanets.MonthSelect.ListIndex + 1
dat.DD = frmPlanets.DaySelect
tt = (frmPlanets.Hrs + frmPlanets.Min / 60 + frmPlanets.Sec / 3600) / 24
dat.DD = dat.DD + tt
JD = KalenderNaarJD(dat)
schaal = 1
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
    frmSaturnus.Text1.Text = Format(Int(dDatum.DD), "00") + "-" + Format(dDatum.MM, "00") + "-" + Format(dDatum.jj, "00") + " " + Format(dDatum.DD - Int(dDatum.DD), "hh:mm:ss")
    Call Zomertijd_Wintertijd(dDatum.jj, JD_ZT, JD_WT)
    JD_Calc = JD + TijdCorrectie(JD, JD_ZT, JD_WT) + ApproxDeltaT(JDToT(JD)) / 86400
    Call tekenSaturnusRingsMoons(Me.picHidden, JD_Calc, schaal)
    BitBlt picCanvas.hdc, _
    0, 0, picCanvas.ScaleWidth, picCanvas.ScaleHeight, _
    picHidden.hdc, 0, 0, SRCCOPY
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
    If schaal < 10 Then schaal = schaal + 0.1
ElseIf KeyAscii = 45 Then '-
    picHidden.Cls
    If schaal > 0.6 Then schaal = schaal - 0.1
ElseIf KeyAscii = 46 Then '.
    nTmpStap = nStap
    nStap = 0
ElseIf KeyAscii = 44 Then ',
    nStap = nTmpStap
End If
End Sub

