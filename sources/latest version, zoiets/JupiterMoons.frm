VERSION 5.00
Begin VB.Form frmJupiterMoons 
   BackColor       =   &H8000000A&
   Caption         =   "Moons of Jupiter"
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15375
   Icon            =   "JupiterMoons.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   9435
   ScaleWidth      =   15375
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picHidden 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      Height          =   8475
      Left            =   0
      ScaleHeight     =   8475
      ScaleWidth      =   15165
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   15165
      Begin VB.Image imgJup 
         Height          =   5010
         Left            =   2160
         Picture         =   "JupiterMoons.frx":030A
         Top             =   600
         Width           =   5220
      End
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Height          =   8505
      Left            =   0
      ScaleHeight     =   8445
      ScaleWidth      =   15165
      TabIndex        =   4
      Top             =   840
      Width           =   15225
   End
   Begin VB.CommandButton Command3 
      Caption         =   "S&tep"
      Height          =   495
      Left            =   10680
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   360
      Top             =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&End"
      Height          =   495
      Left            =   8760
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start"
      Height          =   495
      Left            =   6840
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image imgMaan 
      Height          =   3015
      Left            =   3600
      Picture         =   "JupiterMoons.frx":CD2B
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Image imgSchaduw 
      Height          =   3015
      Left            =   3600
      Picture         =   "JupiterMoons.frx":CF50
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Image imgBedekt 
      Height          =   3015
      Left            =   3600
      Picture         =   "JupiterMoons.frx":E86F
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Image imgVerduisterd 
      Height          =   3015
      Left            =   3600
      Picture         =   "JupiterMoons.frx":1086B
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
End
Attribute VB_Name = "frmJupiterMoons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blnGo As Boolean
Private nStap As Double
Private schaal As Double
Private T As Double
Private tt As Double
Private ttx As Double
Private dat As tDatum
Private Obl As Double
Private deltaT As Double
Private NutLon As Double, NutObl As Double
Private vsMaan(4) As TVECTOR
Private vTemp(4) As TVECTOR
Private v As TVECTOR
Private vMaan As TVECTOR
Private vs As TSVECTOR
Private vDummy As TVECTOR
Private W As TVECTOR
Private SHelio As TSVECTOR, SGeo As TSVECTOR
    'Q1 = SHelio, Q2 = SGeo
Private sAarde As TSVECTOR
Private Const MERGEPAINT = &HBB0226
Private Const SRCAND = &H8800C6
Private Const SRCCOPY = &HCC0020
Private JD As Double
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Sub command1_click()
Dim JD_ZT As Double
Dim JD_WT As Double
dat.jj = frmPlanets.Year
dat.MM = frmPlanets.MonthSelect.ListIndex + 1
dat.DD = frmPlanets.DaySelect
tt = (frmPlanets.Hrs + frmPlanets.Min / 60 + frmPlanets.Sec / 3600) / 24
dat.DD = dat.DD + tt
JD = KalenderNaarJD(dat)
T = JDToT(JD)
deltaT = ApproxDeltaT(T)
T = T + deltaT * secToT
Call Zomertijd_Wintertijd(dat.jj, JD_ZT, JD_WT)
T = T + TijdCorrectie(TToJD(T), JD_ZT, JD_WT) / 36525
Call NutationConst(T, NutLon, NutObl)
Obl = Obliquity(T)

schaal = 400
nStap = 1
picHidden.Cls
ttx = -9999
Timer1.Enabled = True
End Sub

Private Sub Command3_Click()
#If FRANS Then
    nStap = Val(InputBox("Donner par in secondes: ", "Donnée pas"))
#Else
    nStap = Val(InputBox("Give step in seconds: ", "Input step"))
#End If
End Sub

Private Sub Form_Activate()
    command1_click
End Sub

Private Sub Form_Load()
#If FRANS Then
    Me.Caption = "Satellites de Jupiter"
    Me.Command1.Caption = "Départ"
    Me.Command2.Caption = "Fin"
    Me.Command3.Caption = "Pas"
#End If
End Sub

Private Sub Timer1_Timer()

Dim i As Long
Dim j As Long
Dim dDatum As tDatum
Dim JD_ZT As Double
Dim JD_WT As Double
Dim Semidiameter As Double, PolarSemiDiameter As Double
    
    DoEvents
    dDatum = JDNaarKalender(TToJD(T))
    Call Zomertijd_Wintertijd(dDatum.jj, JD_ZT, JD_WT)
    dDatum = JDNaarKalender(TToJD(T - TijdCorrectie(TToJD(T), JD_ZT, JD_WT) / 36525 - ApproxDeltaT(T) * secToT))
    frmJupiterMoons.Text1.Text = Format(Int(dDatum.DD)) + "-" + Format(dDatum.MM, "00") + "-" + Format(dDatum.jj, "00") + " " + Format(dDatum.DD - Int(dDatum.DD), "hh:mm:ss")
    'Me.Refresh
    'Form2.Cls

If Abs(ttx - T) > 1 / 36525 Then
    ttx = T
    SHelio.l = 0: SHelio.B = 0: SHelio.r = 0
    Call PlanetPosHi(0, T, sAarde, True)
    Call PlanetPosHi(5, T, SHelio, True)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call PlanetPosHi(5, T - SGeo.r * LightTimeConst, SHelio, True)
    Call HelioToGeo(SHelio, sAarde, SGeo)

End If
Semidiameter = PlanetSemiDiameter(5, SGeo.r, PolarSemiDiameter)
' teken jupiter
    picHidden.FillStyle = 1
    'picHidden.FillColor = RGB(255, 255, 255)
    'picHidden.Circle (picHidden.ScaleWidth / 2, picHidden.ScaleHeight / 2), Schaal, RGB(0, 0, 0), , , 1 / 1.071374
    picHidden.FillColor = RGB(0, 0, 0)
    picHidden.Line (0, 0)-(picHidden.ScaleWidth, picHidden.ScaleHeight), , BF
    picHidden.PaintPicture imgJup, picHidden.ScaleWidth / 2 - schaal, picHidden.ScaleHeight / 2 - schaal, schaal * 2, schaal * 2
    
For i = 1 To 4
    'If (vTemp(i).Z < 0) And (Abs((vSMaan(i).x * vSMaan(i).x + vSMaan(i).y * vSMaan(i).y)) < 1) Then
    '    Call TekenCirkelWit(vSMaan(i), Schaal, Schaal, RGB(255, 255, 255))
    'End If
     'Call TekenCirkelWit(vTemp(i), Schaal, Schaal, RGB(255, 255, 255))
    ' de maantjes zelf
    'Debug.Print
    Call JSatEclipticPosition(DUMMY_SATELLITE, T - SGeo.r * LightTimeConst, vDummy)
    Call JSatViewFrom(DUMMY_SATELLITE, vDummy, SGeo, vDummy, False, vDummy, False)
    Call JSatEclipticPosition(i, T - SGeo.r * LightTimeConst, vMaan)
    Call JSatViewFrom(i, vMaan, SGeo, vDummy, False, vMaan, True)
    'Debug.Print "Moon " & i, vMaan.x, vMaan.y, vMaan.z, Sqr(vMaan.x * vMaan.x + vMaan.y * vMaan.y + vMaan.z * vMaan.z)
' teken maantje
    vTemp(i) = vMaan
    vTemp(i).Y = vTemp(i).Y * 1.071374
    Call JSatEclipticPosition(DUMMY_SATELLITE, T - SGeo.r * LightTimeConst, vDummy)
    Call JSatViewFrom(DUMMY_SATELLITE, vDummy, SHelio, vDummy, False, vDummy, False)
    Call JSatEclipticPosition(i, T - SGeo.r * LightTimeConst, vsMaan(i))
    Call JSatViewFrom(i, vsMaan(i), SHelio, vDummy, False, vsMaan(i), True)
    'Debug.Print "Sh.  " & i, vSMaan.x, vMaan.y, , , Sqr(vSMaan.x * vSMaan.x + vSMaan.y * vSMaan.y)
    vsMaan(i).Y = vsMaan(i).Y * 1.071374
    
    If (vTemp(i).Z > 0) And (Abs((vTemp(i).x * vTemp(i).x + vTemp(i).Y * vTemp(i).Y)) < 1) Then
    '  {bedekt}
        'picHidden.PaintPicture imgBedekt, frmJupiterMoons.picHidden.ScaleWidth / 2 - vTemp(i).x * Schaal - 25, frmJupiterMoons.picHidden.ScaleHeight / 2 + vTemp(i).y * Schaal - 25, 50, 50, , , , , vbMergeCopy
        Call TekenCirkel(picHidden, vTemp(i), schaal, schaal, RGB(150, 150, 150))
    Else   '{niet bedekt}
        If (vTemp(i).Z > 0) And (Abs((vsMaan(i).x * vsMaan(i).x + vsMaan(i).Y * vsMaan(i).Y)) < 1) Then
        '    {verduisterd (schaduw jupiter op maantje)}
            'picHidden.PaintPicture imgVerduisterd, frmJupiterMoons.picHidden.ScaleWidth / 2 - vTemp(i).x * Schaal - 25, frmJupiterMoons.picHidden.ScaleHeight / 2 + vTemp(i).y * Schaal - 25, 50, 50, , , , , vbMergeCopy
            Call TekenCirkel(picHidden, vTemp(i), schaal, schaal, RGB(0, 0, 255))
        Else  '{niet bedekt en niet verduisterd}
            '{zichtbaar}
            Call TekenCirkel(picHidden, vTemp(i), schaal, schaal, RGB(255, 255, 255))
            'picHidden.PaintPicture imgMaan, frmJupiterMoons.picHidden.ScaleWidth / 2 - vTemp(i).x * Schaal - 25, frmJupiterMoons.picHidden.ScaleHeight / 2 + vTemp(i).y * Schaal - 25, 50, 50, , , , , vbMergeCopy
        End If
    End If
    If (vTemp(i).Z < 0) And (Abs((vsMaan(i).x * vsMaan(i).x + vsMaan(i).Y * vsMaan(i).Y)) < 1) Then
           '{schaduw op jupiter}
           'picHidden.PaintPicture imgSchaduw, frmJupiterMoons.picHidden.ScaleWidth / 2 - vSMaan(i).x * Schaal - 25, frmJupiterMoons.picHidden.ScaleHeight / 2 + vSMaan(i).y * Schaal - 25, 50, 50, , , , , vbMergeCopy
           Call TekenCirkel(picHidden, vsMaan(i), schaal, schaal, RGB(0, 0, 0))
    End If
Next
BitBlt picCanvas.hdc, _
0, 0, picCanvas.ScaleWidth, picCanvas.ScaleHeight, _
picHidden.hdc, 0, 0, SRCCOPY
picCanvas.Refresh
T = T + nStap / 1440 / 36525 / 60 * (Timer1.Interval / 1000)
End Sub

Private Sub Picture1_Click()
Call CircleDemo
End Sub

Sub CircleDemo()
Dim i
   Dim radius
    picHidden.FillStyle = 0
For i = 1 To 10000
    picHidden.FillColor = 255
   picHidden.Circle (XPos, YPos), radius, RGB(255, 255, 255)
' Set Red to a random value.
   r = 255 * Rnd
' Set Green to a random value.
   g = 255 * Rnd
' Set Blue to a random value.
   B = 255 * Rnd
' Set x-coordinate in middle of form.
   XPos = picHidden.ScaleWidth / 2
' Set y-coordinate in middle of form.
   YPos = picHidden.ScaleHeight / 2
   ' Set radius between 0 & 50% of form height.
   radius = ((YPos * 0.9) + 1) * Rnd
   ' Draw the circle using a random color.
    picHidden.FillColor = 123
   picHidden.Circle (XPos, YPos), radius, RGB(r, g, B)
Next
    picHidden.FillColor = 0
   picHidden.Circle (XPos, YPos), radius, RGB(255, 255, 255)
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
blnGo = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Static nTmpStap As Long
If KeyAscii = 62 Then '>
    nStap = nStap + 1
ElseIf KeyAscii = 60 Then '<
    nStap = nStap - 1
ElseIf KeyAscii = 43 Then '+
    picHidden.Cls
    If schaal < 1800 Then schaal = schaal + 10
ElseIf KeyAscii = 45 Then '-
    picHidden.Cls
    If schaal > 150 Then schaal = schaal - 10
ElseIf KeyAscii = 46 Then '.
    nTmpStap = nStap
    nStap = 0
ElseIf KeyAscii = 44 Then ',
    nStap = nTmpStap
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
blnGo = False
End Sub

Private Sub Form_Terminate()
blnGo = False
End Sub

