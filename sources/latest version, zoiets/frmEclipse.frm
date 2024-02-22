VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEclipse 
   Caption         =   "Zonsverduistering"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14475
   Icon            =   "frmEclipse.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   14475
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmVoortgang 
      Caption         =   "Voortgang"
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   10080
      Width           =   13935
      Begin VB.Label lblVoortgang 
         BackColor       =   &H00FF0000&
         Height          =   235
         Left            =   50
         TabIndex        =   10
         Top             =   200
         Width           =   13800
      End
   End
   Begin RichTextLib.RichTextBox rtfKoptekst 
      Height          =   735
      Left            =   600
      TabIndex        =   8
      Top             =   2205
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   1296
      _Version        =   393217
      TextRTF         =   $"frmEclipse.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfBerekening 
      Height          =   6620
      Left            =   600
      TabIndex        =   7
      Top             =   3000
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   11668
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmEclipse.frx":0392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStripEclipse 
      Height          =   8055
      Left            =   360
      TabIndex        =   6
      Top             =   1800
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   14208
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   8
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Local"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Central line"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Limits"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contacts"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Outline curves"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Rise and Set"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Greatest eclipse"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General information"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdBereken 
      Caption         =   "&Calculate"
      Default         =   -1  'True
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.Frame frmVoorspelling 
      Caption         =   "Solareclipse"
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   10575
      Begin VB.CommandButton cmdVolgende 
         Caption         =   "&Next"
         Height          =   375
         Left            =   5880
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdVorig 
         Caption         =   "&Last"
         Height          =   375
         Left            =   4680
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtDatumVerduistering 
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Text            =   "txtDatumVerduistering"
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblEclipse 
         Caption         =   "Date solareclipse"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmEclipse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private dat As tDatum
Private SolarEclipse As SOLARECLIPSEDATA
Private LunarEclipse As LUNARECLIPSEDATA
Private SolarMoonJD As Double
Private nInitHeight As Long
Private blndoorgaan As Boolean
Private nWidthVoortgang As Long
Private Sub cmdBereken_Click()
blndoorgaan = True
Select Case TabStripEclipse.SelectedItem.Caption
Case Is = "Local"
    Bereken_plaatselijk
Case Is = "Central line"
    Centrale_lijn
Case Is = "Limits"
    Limieten
Case Is = "Contacts"
    Contacts
Case Is = "Outline curves"
    OutlineCurves
Case Is = "Rise and Set"
    RiseSetCurves
Case Is = "Greatest eclipse"
    GrootsteEclips
Case Is = "General information"
    Maanberekening
End Select
End Sub
Private Sub Maanberekening()
Dim i As Long
Dim JD As Double
'LunarEclipse geeft een eerste en goede benadering voor de tijdstippen en de magnitude
'Deze kunnen echter nog nauwkeuriger bepaald worden

rtfBerekening.Text = ""
With LunarEclipse
    Select Case .EclipseType
        Case PENUMBRAL
            i = 3
        Case PARTIAL
            i = 2
        Case TOTAL
            i = 1
    End Select
    
    sMax = Trim(s) + " (Maximum)"
    If i <= 3 Then
        JD = NauwkeurigerTijdstipMaansverduistering(.JD - .SpartPenumbra, "PB")
        rtfBerekening.Text = rtfBerekening.Text + MaakDatumstringT(JDToT(JD)) & "    First contact with penumbra" & vbCrLf
        If i <= 2 Then
            JD = NauwkeurigerTijdstipMaansverduistering(.JD - .SpartUmbra, "UB")
            rtfBerekening.Text = rtfBerekening.Text + MaakDatumstringT(JDToT(JD)) & "        First contact withe umbra" & vbCrLf
            If i = 1 Then
                JD = NauwkeurigerTijdstipMaansverduistering(.JD - .StotUmbra, "TB")
                rtfBerekening.Text = rtfBerekening.Text + MaakDatumstringT(JDToT(JD)) & "            Begin total eclipse" & vbCrLf
                JD = NauwkeurigerTijdstipMaansverduistering(.JD, "T")
                rtfBerekening.Text = rtfBerekening.Text + MaakDatumstringT(JDToT(JD)) & "                Maximum eclipse, magn. " & Format(MoonEclipseMagnitude(JD, "TT"), "0.000") & vbCrLf
                JD = NauwkeurigerTijdstipMaansverduistering(.JD + .StotUmbra, "TE")
                rtfBerekening.Text = rtfBerekening.Text + MaakDatumstringT(JDToT(JD)) & "            End total eclipse" & vbCrLf
            End If
            If i = 2 Then
                JD = NauwkeurigerTijdstipMaansverduistering(.JD, "T")
                rtfBerekening.Text = rtfBerekening.Text + MaakDatumstringT(JDToT(JD)) & "            Maximum eclipse, magn. " & Format(MoonEclipseMagnitude(JD, "TU"), "0.000") & vbCrLf
            End If
            JD = NauwkeurigerTijdstipMaansverduistering(.JD + .SpartUmbra, "UE")
            rtfBerekening.Text = rtfBerekening.Text + MaakDatumstringT(JDToT(JD)) & "        Last contact with umbra" & vbCrLf
        End If
        If i = 3 Then
            JD = NauwkeurigerTijdstipMaansverduistering(.JD, "T")
            rtfBerekening.Text = rtfBerekening.Text + MaakDatumstringT(JDToT(JD)) & "        Maximum eclipse, magn. " & Format(MoonEclipseMagnitude(JD, "TP"), "0.000") & vbCrLf
        End If
        JD = NauwkeurigerTijdstipMaansverduistering(.JD + .SpartPenumbra, "PE")
        rtfBerekening.Text = rtfBerekening.Text + MaakDatumstringT(JDToT(JD)) & "    Last contact with penumbra" & vbCrLf
    End If
End With
End Sub
Private Function NauwkeurigerTijdstipMaansverduistering(ByVal JD As Double, ByVal stype As String) As Double
Dim T As Double, T0 As Double
Dim BessElmt As tBessElmt, AuxElmt As tAuxElmt, dBess As tDiffBess, PredData As tPredData, _
    Extr As tExtremes, limits As tLimits, OutCurve As tOutCurve, MaxEclCurve As tMaxEclCurve, _
    RiseSet As tRiseSetCurve, RSMax As tRSMaxCurve, GreatestEclipse As tGreatestEclipse, _
    localeclipse As tLocalEclipse
Dim AppTime        As Double
Dim RkM As Double, DecM As Double, ParM As Double, RkZ As Double, DecZ As Double, ParZ As Double, RZ     As Double
Dim BessElmt1 As tBessElmt, BessElmt2 As tBessElmt, sParZ0 As Double
Dim j
Dim eps As Double
Dim dRkmM As Double, dDecM As Double, Lambda As Double, b As Double
Dim Obl As Double

T = JDToT(JD)
Obl = Obliquity(T + 1 / 876600)
Call PositieZonMaan(T + 1 / 876600, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime)
Call EquToEcl(RkM, DecM, Obl, Lambda, b)
dRkM = -Sin(Obl) * Cos(Lambda) / Cos(DecM) / Cos(DecM) * -0.6 / 3600 * Pi / 180
dDecM = (Cos(Obl) * Cos(Lambda) * Cos(RkM) + Sin(Lambda) * Sin(RkM)) * -0.6 / 3600 * Pi / 180
RkM = RkM + dRkM
DecM = DecM + dDecM
BessElmt2.x = modpi((RkM - (RkZ + Pi))) * Cos(DecM)
eps = 0.25 * modpi((RkM - (RkZ + Pi))) * Sin(2 * -DecZ) * Sin((RkM - (RkZ + Pi)))
BessElmt2.y = modpi(DecM + DecZ + eps)
'Call Bess_elmts(RkM, DecM, ParM, RkZ + Pi, -DecZ, ParZ, RZ, AppTime, BessElmt2)
BessElmt2.x = BessElmt2.x * 180 / Pi * 3600: BessElmt2.y = BessElmt2.y * 180 / Pi * 3600

Obl = Obliquity(T)
Call PositieZonMaan(T, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime)
Call EquToEcl(RkM, DecM, Obl, Lambda, b)
dRkM = -Sin(Obl) * Cos(Lambda) / Cos(DecM) / Cos(DecM) * -0.6 / 3600 * Pi / 180
dDecM = (Cos(Obl) * Cos(Lambda) * Cos(RkM) + Sin(Lambda) * Sin(RkM)) * -0.6 / 3600 * Pi / 180
RkM = RkM + dRkM
DecM = DecM + dDecM
sParZ0 = asin(0.272274 * Sin(ParM))
BessElmt1.x = modpi((RkM - (RkZ + Pi))) * Cos(DecM)
eps = 0.25 * modpi((RkM - (RkZ + Pi))) * Sin(2 * -DecZ) * Sin((RkM - (RkZ + Pi)))
BessElmt1.y = modpi(DecM + DecZ + eps)
' Call Bess_elmts(RkM, DecM, ParM, RkZ + Pi, -DecZ, ParZ, RZ, AppTime, BessElmt1)
BessElmt1.x = BessElmt1.x * 180 / Pi * 3600: BessElmt1.y = BessElmt1.y * 180 / Pi * 3600
Call DiffBess(BessElmt1, BessElmt2, dBess)

If stype = "PB" Or stype = "PE" Then
    lf = 1.02 * (0.99834 * ParM + 959.63 / 3600 * Pi / 180 / RZ + SolarParallax / RZ)
    lL = lf + sParZ0
ElseIf stype = "UB" Or stype = "UE" Then
    lf = 1.02 * (0.99834 * ParM - 959.63 / 3600 * Pi / 180 / RZ + SolarParallax / RZ)
    lL = lf + sParZ0
ElseIf stype = "TB" Or stype = "TE" Then
    lf = 1.02 * (0.99834 * ParM - 959.63 / 3600 * Pi / 180 / RZ + SolarParallax / RZ)
    lL = (lf - sParZ0)
End If
lL = lL * 180 / Pi * 3600
n2 = dBess.x1 * dBess.x1 + dBess.y1 * dBess.y1
n = Sqr(n2)
xx1pyy1 = BessElmt1.x * dBess.x1 + BessElmt1.y * dBess.y1
xy1_x1y = BessElmt1.x * dBess.y1 - dBess.x1 * BessElmt1.y
lDelta = Abs(1 / n * xy1_x1y)
lt = -1 / n2 * (xx1pyy1)
If stype = "PB" Or stype = "UB" Or stype = "TB" Then
    lt = lt - Sqr(lL * lL - lDelta * lDelta) / n
ElseIf stype = "PE" Or stype = "UE" Or stype = "TE" Then
    lt = lt + Sqr(lL * lL - lDelta * lDelta) / n
ElseIf stype = "T" Then
    'lt = lt
End If
NauwkeurigerTijdstipMaansverduistering = JD + lt / 24 - ApproxDeltaT(T) / 86400

'sType = PB (penumbral begin, PE (einde), UB (umbral begin), UE (einde), M (maximum)
End Function
Private Function MoonEclipseMagnitude(ByVal JD As Double, stype As String) As Double
Dim T As Double, T0 As Double
Dim BessElmt As tBessElmt, AuxElmt As tAuxElmt, dBess As tDiffBess, PredData As tPredData, _
    Extr As tExtremes, limits As tLimits, OutCurve As tOutCurve, MaxEclCurve As tMaxEclCurve, _
    RiseSet As tRiseSetCurve, RSMax As tRSMaxCurve, GreatestEclipse As tGreatestEclipse, _
    localeclipse As tLocalEclipse
Dim AppTime        As Double
Dim RkM As Double, DecM As Double, ParM As Double, RkZ As Double, DecZ As Double, ParZ As Double, RZ     As Double
Dim BessElmt1 As tBessElmt, BessElmt2 As tBessElmt, sParZ0 As Double

JD = JD + ApproxDeltaT(T) / 86400
Call PositieZonMaan(JDToT(JD), RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime)
sParZ0 = asin(0.272274 * Sin(ParM))
Call Bess_elmts(RkM, DecM, ParM, RkZ + Pi, -DecZ, ParZ, RZ, AppTime, BessElmt1)
BessElmt1.x = BessElmt1.x * ParM * 180 / Pi * 3600: BessElmt1.y = BessElmt1.y * ParM * 180 / Pi * 3600
M = Sqr(BessElmt1.x * BessElmt1.x + BessElmt1.y * BessElmt1.y)
If stype = "TP" Then
    lf = 1.02 * (0.99834 * ParM + 959.63 / 3600 * Pi / 180 / RZ)
    lL = lf + sParZ0
ElseIf stype = "TU" Or stype = "TT" Then
    lf = 1.02 * (0.99834 * ParM - 959.63 / 3600 * Pi / 180 / RZ)
    lL = lf + sParZ0
End If
lL = lL * 180 / Pi * 3600
MoonEclipseMagnitude = 1 / (2 * sParZ0 * 180 / Pi * 3600) * (lL - M)
End Function
Private Sub GrootsteEclips()
Dim T As Double, T0 As Double
Dim BessElmt As tBessElmt, AuxElmt As tAuxElmt, dBess As tDiffBess, PredData As tPredData, _
    Extr As tExtremes, limits As tLimits, OutCurve As tOutCurve, MaxEclCurve As tMaxEclCurve, _
    RiseSet As tRiseSetCurve, RSMax As tRSMaxCurve, GreatestEclipse As tGreatestEclipse, _
    localeclipse As tLocalEclipse

Dim AppTime        As Double
Dim RkM As Double, DecM As Double, ParM As Double, RkZ As Double, DecZ As Double, ParZ As Double, RZ     As Double
Dim BessElmt1 As tBessElmt, BessElmt2 As tBessElmt

Dim lx0 As Double, ly0 As Double, lm2 As Double, ly10 As Double, lm12 As Double, lrho2 As Double, lrho As Double, lx1 As Double, ly1     As Double
Dim ln1 As Double, ln12 As Double, ll1pRho As Double, ln1ll1pRho As Double, lx0x1py0y1 As Double, lx0y1mx1y0 As Double, lsX1 As Double, lcX1     As Double
Dim lt As Double, lXi As Double, lNu As Double, lcPhisd As Double, lcPhicd As Double, ltDel As Double, lDel As Double, lLambda As Double, lsPhi As Double, lPhi     As Double
Dim lsPhi1 As Double, lPhi1     As Double
Dim i As Long, pPsi As Double
Dim nRes As Boolean
Dim nRes2 As Long
Dim JD As Double

dat = JDNaarKalender(SolarEclipse.JD)
dat.DD = Int(dat.DD * 24 * 60) / 24 / 60
JD = KalenderNaarJD(dat)
T = JDToT(JD)
deltaT = ApproxDeltaT(T) / 86400 / 36525
T0 = T
rtfBerekening = ""

'Debug.Print "Time, position and magnitude of greatest eclipse"

'Greatest eclipse berekent gegevens voor gedeeltelijke zonsverduistering
'voor totale verduisteringen is dit niet geschikt
If Greatest_Eclipse(T0, GreatestEclipse) Then
    rtfBerekening.Text = rtfBerekening.Text & MaakDatumstringT(GreatestEclipse.T - deltaT)
    Call PositieZonMaan(GreatestEclipse.T, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime, BessElmt1)
    Call PositieZonMaan(GreatestEclipse.T + 1 / 876600, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime, BessElmt2)
    Call DiffBess(BessElmt1, BessElmt2, dBess)
    Call Aux_elmts(BessElmt1, AuxElmt, dBess)
    
    nRes = PredDataSolarEcl(BessElmt1, AuxElmt, dBess, PredData)
    'onderzoeken of totale verduistering
    If Local_Eclipse(GreatestEclipse.T, PredData.Phi * 180 / Pi, PredData.Lambda * 180 / Pi, 0, localeclipse, "T") Then
        rtfBerekening.Text = rtfBerekening.Text & "  " & StrHMS_DMS(PredData.Lambda * 180 / Pi, 7, 1, False, True, "g", 4) _
            & "  " & StrHMS_DMS(PredData.Phi * 180 / Pi, 7, 1, False, False, "g", 3)
        rtfBerekening.Text = rtfBerekening.Text & "  " & Format(localeclipse.MagTotaal, "0.00000")
    Else
        rtfBerekening.Text = rtfBerekening.Text & "  " & StrHMS_DMS(GreatestEclipse.pos.lng * 180 / Pi, 7, 1, False, True, "g", 4) _
            & "  " & StrHMS_DMS(GreatestEclipse.pos.nb * 180 / Pi, 7, 1, False, False, "g", 3)
        rtfBerekening.Text = rtfBerekening.Text & "  " & Format(GreatestEclipse.magn, "0.00000")
    End If
End If
End Sub
Private Sub RiseSetCurves()
Dim T As Double, T0 As Double
Dim BessElmt As tBessElmt, AuxElmt As tAuxElmt, dBess As tDiffBess, PredData As tPredData, _
    Extr As tExtremes, limits As tLimits, OutCurve As tOutCurve, MaxEclCurve As tMaxEclCurve, _
    RiseSet As tRiseSetCurve, RSMax As tRSMaxCurve, GreatestEclipse As tGreatestEclipse, _
    localeclipse As tLocalEclipse

Dim AppTime        As Double
Dim RkM As Double, DecM As Double, ParM As Double, RkZ As Double, DecZ As Double, ParZ As Double, RZ     As Double
Dim BessElmt1 As tBessElmt, BessElmt2 As tBessElmt
Dim lx0 As Double, ly0 As Double, lm2 As Double, ly10 As Double, lm12 As Double, lrho2 As Double, lrho As Double, lx1 As Double, ly1     As Double
Dim ln1 As Double, ln12 As Double, ll1pRho As Double, ln1ll1pRho As Double, lx0x1py0y1 As Double, lx0y1mx1y0 As Double, lsX1 As Double, lcX1     As Double
Dim lt As Double, lXi As Double, lNu As Double, lcPhisd As Double, lcPhicd As Double, ltDel As Double, lDel As Double, lLambda As Double, lsPhi As Double, lPhi     As Double
Dim lsPhi1 As Double, lPhi1     As Double
Dim i As Long, pPsi As Double
Dim nRes As Boolean
Dim nRes2 As Long
Dim JD As Double

dat = JDNaarKalender(SolarEclipse.JD)
dat.DD = Int(dat.DD * 24 * 60) / 24 / 60
JD = KalenderNaarJD(dat)
T = JDToT(JD)
deltaT = ApproxDeltaT(T) / 86400 / 36525
T0 = T - 6 / 876600 '
rtfBerekening = ""
While T0 < T + 6 / 876600 And blndoorgaan
    lblVoortgang.Width = nWidthVoortgang * (T0 - T + 6 / 876600) / (12 / 876600)
'Debug.Print "Points on the curve of maximum eclipse at sunrise and sunset"
    Call PositieZonMaan(T0 + deltaT, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime, BessElmt1)
    Call PositieZonMaan(T0 + 1 / 876600 + deltaT, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime, BessElmt2)
    Call DiffBess(BessElmt1, BessElmt2, dBess)
    Call Aux_elmts(BessElmt1, AuxElmt, dBess)
    nRes2 = RSMaxCurve(BessElmt1, AuxElmt, dBess, PredData, RSMax)
    If nRes2 <> 0 Then
        rtfBerekening.Text = rtfBerekening.Text & MaakDatumstringT(T0)
        If nRes2 And 1 Then
            rtfBerekening.Text = rtfBerekening.Text & "  " & StrHMS_DMS(RSMax.pos1.lng * 180 / Pi, 1, 1, False, False, "g", 5) _
            & "  " & StrHMS_DMS(RSMax.pos1.nb * 180 / Pi, 1, 1, False, True, "g", 5)
        Else
            rtfBerekening.Text = rtfBerekening.Text & "     - " & "  " & "    -  "
        End If
        If nRes2 And 2 Then
            rtfBerekening.Text = rtfBerekening.Text & "  " & StrHMS_DMS(RSMax.pos2.lng * 180 / Pi, 1, 1, False, False, "g", 5) _
            & "  " & StrHMS_DMS(RSMax.pos2.nb * 180 / Pi, 1, 1, False, True, "g", 5)
        Else
            rtfBerekening.Text = rtfBerekening.Text & "     - " & "  " & "    -  "
        End If
        rtfBerekening.Text = rtfBerekening.Text & vbCrLf
    End If
    DoEvents
    T0 = T0 + 0.05 / 876600 / 3
Wend
lblVoortgang.Width = 0
End Sub

Private Sub OutlineCurves()
Dim T As Double, T0 As Double
Dim BessElmt As tBessElmt, AuxElmt As tAuxElmt, dBess As tDiffBess, PredData As tPredData, _
    Extr As tExtremes, limits As tLimits, OutCurve As tOutCurve, MaxEclCurve As tMaxEclCurve, _
    RiseSet As tRiseSetCurve, RSMax As tRSMaxCurve, GreatestEclipse As tGreatestEclipse, _
    localeclipse As tLocalEclipse

Dim AppTime        As Double
Dim RkM As Double, DecM As Double, ParM As Double, RkZ As Double, DecZ As Double, ParZ As Double, RZ     As Double
Dim BessElmt1 As tBessElmt, BessElmt2 As tBessElmt
Dim lx0 As Double, ly0 As Double, lm2 As Double, ly10 As Double, lm12 As Double, lrho2 As Double, lrho As Double, lx1 As Double, ly1     As Double
Dim ln1 As Double, ln12 As Double, ll1pRho As Double, ln1ll1pRho As Double, lx0x1py0y1 As Double, lx0y1mx1y0 As Double, lsX1 As Double, lcX1     As Double
Dim lt As Double, lXi As Double, lNu As Double, lcPhisd As Double, lcPhicd As Double, ltDel As Double, lDel As Double, lLambda As Double, lsPhi As Double, lPhi     As Double
Dim lsPhi1 As Double, lPhi1     As Double
Dim i As Long, pPsi As Double
Dim nRes As Boolean
Dim nRes2 As Long
Dim JD As Double

dat = JDNaarKalender(SolarEclipse.JD)
dat.DD = Int(dat.DD * 24 * 60) / 24 / 60
JD = KalenderNaarJD(dat)
T = JDToT(JD)
deltaT = ApproxDeltaT(T) / 86400 / 36525
T0 = T - 6 / 876600 '
rtfBerekening = ""
While T0 < T + 6 / 876600 And blndoorgaan
    lblVoortgang.Width = nWidthVoortgang * (T0 - T + 6 / 876600) / (12 / 876600)
'Debug.Print "Outline curves of an eclipse"
    Call PositieZonMaan(T0, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime, BessElmt1)
    Call PositieZonMaan(T0 + 1 / 876600, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime, BessElmt2)
    Call DiffBess(BessElmt1, BessElmt2, dBess)
    Call Aux_elmts(BessElmt1, AuxElmt, dBess)
    nRes = Outline1Curve(BessElmt1, AuxElmt, dBess, PredData, OutCurve)
    If nRes Then
        rtfBerekening.Text = rtfBerekening.Text & MaakDatumstringT(T0)
        rtfBerekening.Text = rtfBerekening.Text & "  " & StrHMS_DMS(OutCurve.bQ * 180 / Pi, 1, 1, False, False, "g", 5) _
        & "  " & StrHMS_DMS(OutCurve.eQ * 180 / Pi, 1, 1, False, False, "g", 5) & vbCrLf
        i = 1
        While OutCurve.bQ < OutCurve.eQ
            nRes = Outline2Curve(BessElmt1, AuxElmt, dBess, PredData, OutCurve.bQ, OutCurve)
            If nRes Then
                rtfBerekening.Text = rtfBerekening.Text & "  (" & StrHMS_DMS(OutCurve.pos.lng * 180 / Pi, 1, 1, False, False, "g", 5) _
                & " " & StrHMS_DMS(OutCurve.pos.nb * 180 / Pi, 1, 1, True, False, "g", 5) & ")"
                If i = 0 Then
                    rtfBerekening.Text = rtfBerekening.Text & vbCrLf
                End If
                i = (i + 1) Mod 7
'            Else
'                rtfBerekening.Text = rtfBerekening.Text & vbCrLf
            End If
            OutCurve.bQ = OutCurve.bQ + Pi / 180 'per graad punt berekenen
            DoEvents
        Wend
        rtfBerekening.Text = rtfBerekening.Text & vbCrLf
    End If
    DoEvents
    T0 = T0 + 0.5 / 876600 / 3
Wend
lblVoortgang.Width = 0
End Sub
Private Sub Limieten()
Dim T As Double, T0 As Double
Dim BessElmt As tBessElmt, AuxElmt As tAuxElmt, dBess As tDiffBess, PredData As tPredData, _
    Extr As tExtremes, limits As tLimits, OutCurve As tOutCurve, MaxEclCurve As tMaxEclCurve, _
    RiseSet As tRiseSetCurve, RSMax As tRSMaxCurve, GreatestEclipse As tGreatestEclipse, _
    localeclipse As tLocalEclipse

Dim AppTime        As Double
Dim RkM As Double, DecM As Double, ParM As Double, RkZ As Double, DecZ As Double, ParZ As Double, RZ     As Double
Dim BessElmt1 As tBessElmt, BessElmt2 As tBessElmt
Dim lx0 As Double, ly0 As Double, lm2 As Double, ly10 As Double, lm12 As Double, lrho2 As Double, lrho As Double, lx1 As Double, ly1     As Double
Dim ln1 As Double, ln12 As Double, ll1pRho As Double, ln1ll1pRho As Double, lx0x1py0y1 As Double, lx0y1mx1y0 As Double, lsX1 As Double, lcX1     As Double
Dim lt As Double, lXi As Double, lNu As Double, lcPhisd As Double, lcPhicd As Double, ltDel As Double, lDel As Double, lLambda As Double, lsPhi As Double, lPhi     As Double
Dim lsPhi1 As Double, lPhi1     As Double
Dim i As Long, pPsi As Double
Dim nRes As Boolean
Dim nRes2 As Long
Dim JD As Double

dat = JDNaarKalender(SolarEclipse.JD)
dat.DD = Int(dat.DD * 24 * 60) / 24 / 60
JD = KalenderNaarJD(dat)
T = JDToT(JD)
deltaT = ApproxDeltaT(T) / 86400 / 36525
T0 = T - 6 / 876600 '
rtfBerekening = ""
While T0 < T + 6 / 876600 And blndoorgaan
    lblVoortgang.Width = nWidthVoortgang * (T0 - T + 6 / 876600) / (12 / 876600)
    Call PositieZonMaan(T0 + deltaT, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime, BessElmt1)
    Call PositieZonMaan(T0 + 1 / 876600 + deltaT, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime, BessElmt2)
    Call DiffBess(BessElmt1, BessElmt2, dBess)
    Call Aux_elmts(BessElmt1, AuxElmt, dBess)
    nRes = PredDataSolarEcl(BessElmt1, AuxElmt, dBess, PredData)
    nRes2 = LimitsUmbraPenumbra(BessElmt1, AuxElmt, dBess, PredData, limits)
    If nRes2 <> 0 Then
'    Debug.Print MaakDatumstringT(T0)
        rtfBerekening.Text = rtfBerekening.Text & MaakDatumstringT(T0)
        If nRes2 And 1 Then
            rtfBerekening.Text = rtfBerekening.Text & "  " & StrHMS_DMS(limits.ULimN.lng * 180 / Pi, 3, 1, False, True, "g", 4) _
            & "  " & StrHMS_DMS(limits.ULimN.nb * 180 / Pi, 3, 1, True, False, "g", 3)
        Else
            rtfBerekening.Text = rtfBerekening.Text & "      -     " _
            & "     -     "
        End If
        If nRes2 And 2 Then
            rtfBerekening.Text = rtfBerekening.Text & "  " & StrHMS_DMS(limits.ULimZ.lng * 180 / Pi, 3, 1, False, True, "g", 4) _
            & "  " & StrHMS_DMS(limits.ULimZ.nb * 180 / Pi, 3, 1, True, False, "g", 3)
        Else
            rtfBerekening.Text = rtfBerekening.Text & "      -     " _
            & "     -     "
        End If
        If nRes2 And 4 Then
            rtfBerekening.Text = rtfBerekening.Text & "  " & StrHMS_DMS(limits.PLimN.lng * 180 / Pi, 3, 1, False, True, "g", 4) _
            & "  " & StrHMS_DMS(limits.PLimN.nb * 180 / Pi, 3, 1, True, False, "g", 3)
        Else
            rtfBerekening.Text = rtfBerekening.Text & "      -     " _
            & "     -     "
        End If
        If nRes2 And 8 Then
            rtfBerekening.Text = rtfBerekening.Text & "  " & StrHMS_DMS(limits.PLimZ.lng * 180 / Pi, 3, 1, False, True, "g", 4) _
            & "  " & StrHMS_DMS(limits.PLimZ.nb * 180 / Pi, 3, 1, True, False, "g", 3)
        Else
            rtfBerekening.Text = rtfBerekening.Text & "      -     " _
            & "     -    "
        End If
        rtfBerekening.Text = rtfBerekening.Text & vbCrLf
    End If
    DoEvents
    T0 = T0 + 0.05 / 876600 / 3
Wend
lblVoortgang.Width = 0
End Sub
Private Sub Centrale_lijn()
Dim T As Double, T0 As Double
Dim BessElmt As tBessElmt, AuxElmt As tAuxElmt, dBess As tDiffBess, PredData As tPredData, _
    Extr As tExtremes, limits As tLimits, OutCurve As tOutCurve, MaxEclCurve As tMaxEclCurve, _
    RiseSet As tRiseSetCurve, RSMax As tRSMaxCurve, GreatestEclipse As tGreatestEclipse, _
    localeclipse As tLocalEclipse

Dim AppTime        As Double
Dim RkM As Double, DecM As Double, ParM As Double, RkZ As Double, DecZ As Double, ParZ As Double, RZ     As Double
Dim BessElmt1 As tBessElmt, BessElmt2 As tBessElmt
Dim lx0 As Double, ly0 As Double, lm2 As Double, ly10 As Double, lm12 As Double, lrho2 As Double, lrho As Double, lx1 As Double, ly1     As Double
Dim ln1 As Double, ln12 As Double, ll1pRho As Double, ln1ll1pRho As Double, lx0x1py0y1 As Double, lx0y1mx1y0 As Double, lsX1 As Double, lcX1     As Double
Dim lt As Double, lXi As Double, lNu As Double, lcPhisd As Double, lcPhicd As Double, ltDel As Double, lDel As Double, lLambda As Double, lsPhi As Double, lPhi     As Double
Dim lsPhi1 As Double, lPhi1     As Double
Dim i As Long, pPsi As Double
Dim nRes As Boolean
Dim nRes2 As Long
Dim JD As Double
'rtfKoptekst.Text = "Tijdstip                  Westerlengte Noorderbreedte    duur" & vbCrLf
'rtfKoptekst.Text = rtfKoptekst.Text & "===============================================================" & vbCrLf
dat = JDNaarKalender(SolarEclipse.JD)
dat.DD = Int(dat.DD * 24 * 60) / 24 / 60
JD = KalenderNaarJD(dat)
T = JDToT(JD) - 6 / 876600
deltaT = ApproxDeltaT(T) / 86400 / 36525
T0 = T
rtfBerekening = ""
For i = 0 To 1200
    lblVoortgang.Width = nWidthVoortgang * i / 1200
'While T0 < T + 10 / 876600
    Call PositieZonMaan(T0 + 1 / 876600 + deltaT, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime, BessElmt2)
    Call PositieZonMaan(T0 + deltaT, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime, BessElmt1)
    Call DiffBess(BessElmt1, BessElmt2, dBess)
    Call Aux_elmts(BessElmt1, AuxElmt, dBess)
    nRes = PredDataSolarEcl(BessElmt1, AuxElmt, dBess, PredData)
    If nRes = True Then
'        Debug.Print MaakDatumstringT(T0)
        rtfBerekening.Text = rtfBerekening.Text & MaakDatumstringT(T0) & "  " & StrHMS_DMS(PredData.Lambda * 180 / Pi, 3, 1, False, True, "g", 4) _
        & "     " & StrHMS_DMS(PredData.Phi * 180 / Pi, 3, 1, True, False, "g", 3) _
        & "   " & StrHMS_DMS(2 * PredData.s * 15, 6, 1, True, False, "h", 2)
        If Local_Eclipse(T0 + deltaT, PredData.Phi * 180 / Pi, PredData.Lambda * 180 / Pi, 0, localeclipse, "T") Then
            rtfBerekening.Text = rtfBerekening.Text & "  " & Format(localeclipse.MagTotaal, "0.00000")
        End If
        rtfBerekening.Text = rtfBerekening.Text & vbCrLf

    End If
    T0 = T + 1 / 60 * i / 876600
    If Not blndoorgaan Then Exit For
    DoEvents
Next
lblVoortgang.Width = 0
End Sub

Private Sub Bereken_plaatselijk()
Dim localeclipse As tLocalEclipse
Dim T As Double
Dim delta As Double
Dim M(6)
Dim sLongitude As String, sLatitude As String, sAltitude As String

Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Latitude", _
        sLatitude)
Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Longitude", _
        sLongitude)
Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Altitude", _
        sAltitude)
deltaT = ApproxDeltaT(T) / 86400 / 36525
M(1) = "Begin partial solareclipse : "
M(2) = "End   partial solareclipse : "
M(3) = "Maximum       solareclipse : "
M(4) = "Begin total   solareclipse : "
M(5) = "End   total   solareclipse : "
M(6) = "Max. magn. total eclipse   : "
T = JDToT(SolarEclipse.JD)
rtfBerekening = ""
rtfKopteks = ""
If Local_Eclipse(T, Val(sLatitude), Val(sLongitude), Val(sAltitude), localeclipse, "B") Then Me.rtfBerekening.Text = Me.rtfBerekening.Text & M(1) & MaakDatumstringT(localeclipse.Tb) & vbCrLf
If Local_Eclipse(T, Val(sLatitude), Val(sLongitude), Val(sAltitude), localeclipse, "T") Then
    Me.rtfBerekening.Text = Me.rtfBerekening.Text & M(4) & MaakDatumstringT(localeclipse.Ttotaalb) & vbCrLf
    Me.rtfBerekening.Text = Me.rtfBerekening.Text & M(3) & MaakDatumstringT(localeclipse.Tm) & ", " & Format(localeclipse.MagTotaal, "0.000") & ", diepte " & Format(2 * (localeclipse.MagTotaal - 1), "0.000") & vbCrLf
    Me.rtfBerekening.Text = Me.rtfBerekening.Text & M(5) & MaakDatumstringT(localeclipse.Ttotaale) & vbCrLf
    'If Local_Eclipse(T, LocalEclipse, "M") Then Me.rtfBerekening.Text = Me.rtfBerekening.Text & M(3) & MaakDatumstringT(LocalEclipse.Tm - deltaT) & ", " & Format(LocalEclipse.Mag, "0.000") & vbCrLf
ElseIf Local_Eclipse(T, Val(sLatitude), Val(sLongitude), Val(sAltitude), localeclipse, "M") Then
    Me.rtfBerekening.Text = Me.rtfBerekening.Text & M(3) & MaakDatumstringT(localeclipse.Tm) & ", " & Format(localeclipse.mag, "0.000") & vbCrLf
End If
If Local_Eclipse(T, Val(sLatitude), Val(sLongitude), Val(sAltitude), localeclipse, "E") Then Me.rtfBerekening.Text = Me.rtfBerekening.Text & M(2) & MaakDatumstringT(localeclipse.Te) & vbCrLf
End Sub
Private Sub Contacts()
Dim AppTime        As Double
Dim RkM As Double, DecM As Double, ParM As Double, RkZ As Double, DecZ As Double, ParZ As Double, RZ     As Double
Dim BessElmt1 As tBessElmt, BessElmt2 As tBessElmt, dBess As tDiffBess
Dim lx0 As Double, ly0 As Double, lm2 As Double, ly10 As Double, lm12 As Double, lrho2 As Double, lrho As Double, lx1 As Double, ly1     As Double
Dim ln1 As Double, ln12 As Double, ll1pRho As Double, ln1ll1pRho As Double, lx0x1py0y1 As Double, lx0y1mx1y0 As Double, lsX1 As Double, lcX1     As Double
Dim lt As Double, lXi As Double, lNu As Double, lcPhisd As Double, lcPhicd As Double, ltDel As Double, lDel As Double, lLambda As Double, lsPhi As Double, lPhi     As Double
Dim lsPhi1 As Double, lPhi1     As Double
Dim i As Long, pPsi As Double
Dim nRes As Boolean
Dim nRes2 As Long
Dim JD As Double
Dim localeclipse As tLocalEclipse
Dim T As Double
Dim delta As Double
Dim M(6)
Dim dat As tDatum, PredData As tPredData
Dim BessElmt As tBessElmt, AuxElmt As tAuxElmt, _
    Extr As tExtremes, limits As tLimits, OutCurve As tOutCurve, MaxEclCurve As tMaxEclCurve, _
    RiseSet As tRiseSetCurve, RSMax As tRSMaxCurve, GreatestEclipse As tGreatestEclipse

M(1) = "Begin partial solareclipse : "
M(2) = "End   partial solareclipse : "
M(3) = "Maximum       solareclipse : "
M(4) = "Begin total   solareclipse : "
M(5) = "End   total   solareclipse : "
M(6) = "Max. magn. total eclipse   : "
'rtfKoptekst.Text = "Tijdstip                 duur        P1             U1         Maximaal   magn        P4             U4" & vbCrLf
'rtfKoptekst.Text = rtfKoptekst.Text & "=============================================================================================================" & vbCrLf

dat = JDNaarKalender(SolarEclipse.JD)
dat.DD = Int(dat.DD * 24 * 60) / 24 / 60
JD = KalenderNaarJD(dat)
T = JDToT(JD) - 6 / 876600
deltaT = ApproxDeltaT(T) / 86400 / 36525
T0 = T
rtfBerekening = ""
For i = 0 To 1200
    lblVoortgang.Width = nWidthVoortgang * i / 1200
'While T0 < T + 10 / 876600
    Call PositieZonMaan(T0 + 1 / 876600 + deltaT, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime, BessElmt2)
    Call PositieZonMaan(T0 + deltaT, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime, BessElmt1)
    Call DiffBess(BessElmt1, BessElmt2, dBess)
    Call Aux_elmts(BessElmt1, AuxElmt, dBess)
    If PredDataSolarEcl(BessElmt1, AuxElmt, dBess, PredData) Then
'        Debug.Print MaakDatumstringT(T0)
        Me.rtfBerekening.Text = Me.rtfBerekening.Text & MaakDatumstringT(T0)
        Me.rtfBerekening.Text = Me.rtfBerekening.Text & " " & StrHMS_DMS(2 * PredData.s * 15, 6, 1, True, False, "h", 2) & " "
        If Local_Eclipse(T, PredData.Phi * 180 / Pi, PredData.Lambda * 180 / Pi, 0, localeclipse, "B") Then
            Me.rtfBerekening.Text = Me.rtfBerekening.Text & "  " & MaakDatumstringT(localeclipse.Tb, True) & "  "
        Else
            rtfBerekening.Text = rtfBerekening.Text & "       -     "
        End If
        If Local_Eclipse(T, PredData.Phi * 180 / Pi, PredData.Lambda * 180 / Pi, 0, localeclipse, "T") Then
            Me.rtfBerekening.Text = Me.rtfBerekening.Text & "  " & MaakDatumstringT(localeclipse.Ttotaalb, True)
            Me.rtfBerekening.Text = Me.rtfBerekening.Text & "  " & MaakDatumstringT(localeclipse.Tm, True) & " " & Format(localeclipse.MagTotaal, "0.000") & "  "
            Me.rtfBerekening.Text = Me.rtfBerekening.Text & "  " & MaakDatumstringT(localeclipse.Ttotaale, True) & "  "
            'If Local_Eclipse(T, LocalEclipse, "M") Then Me.rtfBerekening.Text = Me.rtfBerekening.Text & M(3) & MaakDatumstringT(LocalEclipse.Tm - deltaT) & ", " & Format(LocalEclipse.Mag, "0.000") & vbCrLf
        ElseIf Local_Eclipse(T, PredData.Phi * 180 / Pi, PredData.Lambda * 180 / Pi, 0, localeclipse, "M") Then
            rtfBerekening.Text = rtfBerekening.Text & "       -     "
            Me.rtfBerekening.Text = Me.rtfBerekening.Text & "  " & MaakDatumstringT(localeclipse.Tm, True) & " " & Format(localeclipse.mag, "0.000") & "  "
            rtfBerekening.Text = rtfBerekening.Text & "       -       "
        Else
            rtfBerekening.Text = rtfBerekening.Text & "       -     "
            rtfBerekening.Text = rtfBerekening.Text & "       -             "
            rtfBerekening.Text = rtfBerekening.Text & "       -       "
        End If
        If Local_Eclipse(T, PredData.Phi * 180 / Pi, PredData.Lambda * 180 / Pi, 0, localeclipse, "E") Then
            Me.rtfBerekening.Text = Me.rtfBerekening.Text & "  " & MaakDatumstringT(localeclipse.Te, True)
        Else
            rtfBerekening.Text = rtfBerekening.Text & "       -     "
        End If
        Me.rtfBerekening.Text = Me.rtfBerekening.Text & vbCrLf
    End If
    If Not blndoorgaan Then Exit For
    T0 = T + 1 / 60 * i / 876600
    DoEvents
Next
lblVoortgang.Width = 0
End Sub
Private Sub cmdVolgende_Click()
Call NextLunarEclipse(SolarMoonJD, 0, LunarEclipse)
Call NextSolarEclipse(SolarMoonJD, SolarEclipse)
If LunarEclipse.JD < SolarEclipse.JD Then
    Me.Caption = "Mooneclipse"
    SolarMoonJD = LunarEclipse.JD
    TabStripEclipse.Tabs(TabStripEclipse.Tabs.Count).Selected = True
Else
    Me.Caption = "Solareclipse"
    SolarMoonJD = SolarEclipse.JD
    If TabStripEclipse.Tabs(TabStripEclipse.Tabs.Count).Selected Then
        TabStripEclipse.Tabs(1).Selected = True
    End If
End If
frmVoorspelling.Caption = Me.Caption
lblEclipse.Caption = "Date " & LCase(Me.Caption)

Call ZetDatum(SolarMoonJD)
End Sub

Private Sub cmdVorig_Click()
Call LastSolarEclipse(SolarMoonJD, SolarEclipse)
Call LastLunarEclipse(SolarMoonJD, 0, LunarEclipse)
If LunarEclipse.JD > SolarEclipse.JD Then
    Me.Caption = "Mooneclipse"
    SolarMoonJD = LunarEclipse.JD
    TabStripEclipse.Tabs(TabStripEclipse.Tabs.Count).Selected = True
Else
    Me.Caption = "Solareclipse"
    SolarMoonJD = SolarEclipse.JD
    If TabStripEclipse.Tabs(TabStripEclipse.Tabs.Count).Selected Then
        TabStripEclipse.Tabs(1).Selected = True
    End If
End If
frmVoorspelling.Caption = Me.Caption
lblEclipse.Caption = "Date " & LCase(Me.Caption)

Call ZetDatum(SolarMoonJD)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If (KeyCode = 17 And Shift = 0) Or (KeyCode = 67 And Shift = 2) Then
    Clipboard.Clear
    Clipboard.SetText (Me.rtfKoptekst.Text & vbCrLf & String(InStr(Me.rtfBerekening.Text, vbCrLf), "=") & vbCrLf & Me.rtfBerekening.Text)
End If
End Sub

Private Sub Form_Load()
Dim tt As Double
Dim JD As Double
dat.jj = frmPlanets.Year
dat.mm = frmPlanets.MonthSelect.ListIndex + 1
dat.DD = frmPlanets.DaySelect
tt = (Hrs + Min / 60 + Sec / 3600) / 24
dat.DD = dat.DD + tt
SolarMoonJD = KalenderNaarJD(dat)
Call cmdVolgende_Click
'Call NextSolarEclipse(JD, SolarEclipse)
'SolarMoonJD = SolarEclipse.JD
'Call ZetDatum(JD)
rtfBerekening = ""
rtfKoptekst = ""
nInitHeight = rtfBerekening.Height
nWidthVoortgang = Me.lblVoortgang.Width
lblVoortgang.Width = 0
End Sub

'Zetdatum zet datum op scherm maar rekening houdend met WT/ZT
Sub ZetDatum(JD As Double)
Dim dat As tDatum
Dim jd1 As Double, JD_ZT As Double, JD_WT As Double

jd1 = JD - ApproxDeltaT(JDToT(JD)) / 86400
dat = JDNaarKalender(jd1)
Call Zomertijd_Wintertijd(dat.jj, JD_ZT, JD_WT)
dat = JDNaarKalender(jd1 - TijdCorrectie(jd1, JD_ZT, JD_WT))
txtDatumVerduistering = MaakDatumstring(dat, False)
End Sub

Private Function MaakDatumstring(dat As tDatum, alleentijd As Boolean)
If alleentijd Then
    MaakDatumstring = StrHMS_DMS(Frac(dat.DD) * 360, 7, 1, False, False, "h", 2)
Else
    MaakDatumstring = Format(Int(dat.DD), "00") & "-" & Format(dat.mm, "00") & "-" & Format(dat.jj) & ":" & StrHMS_DMS(Frac(dat.DD) * 360, 7, 1, False, False, "h", 2)
End If
End Function

Private Function MaakDatumstringT(ByVal T As Double, Optional alleentijd As Boolean = False)
Dim JD As Double
Dim dat As tDatum
Dim JD_ZT As Double, JD_WT As Double

JD = TToJD(T)
dat = JDNaarKalender(JD)
Call Zomertijd_Wintertijd(dat.jj, JD_ZT, JD_WT)
dat = JDNaarKalender(JD - TijdCorrectie(JD, JD_ZT, JD_WT))
MaakDatumstringT = MaakDatumstring(dat, alleentijd)
End Function

Private Sub Form_Unload(Cancel As Integer)
blndoorgaan = False
End Sub

Private Sub TabStripEclipse_Click()
If Me.Caption = "Mooneclipse" And Not TabStripEclipse.Tabs(TabStripEclipse.Tabs.Count).Selected Then
   TabStripEclipse.Tabs(TabStripEclipse.Tabs.Count).Selected = True
ElseIf Me.Caption = "Solareclipse" And TabStripEclipse.Tabs(TabStripEclipse.Tabs.Count).Selected Then
        TabStripEclipse.Tabs(1).Selected = True
End If
rtfKoptekst.Text = ""
Me.rtfBerekening.Text = ""
Select Case TabStripEclipse.SelectedItem.Caption
Case Is = "Local"
Case Is = "Central line"
    rtfKoptekst.Text = vbCrLf & vbCrLf
    rtfKoptekst.Text = rtfKoptekst.Text & "Time                   Longitude W. Latitude       Duration   Magn"
Case Is = "Limits"
    rtfKoptekst.Text = "Time                  |             Limits total eclipse            |             Limits partial eclipse" & vbCrLf
    rtfKoptekst.Text = rtfKoptekst.Text & "                      |         North                 South         |         North                 South" & vbCrLf
    rtfKoptekst.Text = rtfKoptekst.Text & "                      |  Long. W.    Latitude   Long. W.    Latitude|  Long. W.    Latitude   Long. W.    Latitude."
Case Is = "Contacts"
    rtfKoptekst.Text = vbCrLf & vbCrLf
    rtfKoptekst.Text = rtfKoptekst.Text & "Time                   Duration      P1             U1         Maximum    Magn        U4             P4"
Case Is = "Outline curves"
    rtfKoptekst.Text = vbCrLf & vbCrLf
    rtfKoptekst.Text = rtfKoptekst.Text & "Time     with (long. w., lat.) as outline of penumbra"
Case Is = "Rise and Set"
    rtfKoptekst.Text = "                                 Solareclipse" & vbCrLf
    rtfKoptekst.Text = rtfKoptekst.Text & "                           Rising           Set" & vbCrLf
    rtfKoptekst.Text = rtfKoptekst.Text & "Time                    WL        Lat   WL        Lat"
Case Is = "Greatest eclipse"
    rtfKoptekst.Text = vbCrLf & vbCrLf
    rtfKoptekst.Text = rtfKoptekst.Text & "Time                    Longitude W.    Latitude      Magn."
End Select
End Sub
