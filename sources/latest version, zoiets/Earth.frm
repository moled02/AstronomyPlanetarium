VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlanets 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Astronomy"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Earth.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkGrootstePrecisie 
      Caption         =   "Greatest Precision"
      Height          =   255
      Left            =   3720
      TabIndex        =   26
      Top             =   840
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.Timer tmrInterval 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8760
      Top             =   7560
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6360
      ItemData        =   "Earth.frx":030A
      Left            =   120
      List            =   "Earth.frx":030C
      MultiSelect     =   2  'Extended
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1200
      Width           =   9015
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Working objects"
      ForeColor       =   &H00000000&
      Height          =   3615
      Left            =   9720
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   3375
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Frame1"
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   60
         TabIndex        =   22
         Top             =   720
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   60
         TabIndex        =   21
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Check1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   60
         TabIndex        =   20
         Top             =   1680
         Width           =   1875
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "Earth.frx":030E
         Left            =   120
         List            =   "Earth.frx":0310
         TabIndex        =   19
         Text            =   "Combo1"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Command1"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label1"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   180
         TabIndex        =   24
         Top             =   2760
         Width           =   1815
      End
   End
   Begin VB.CommandButton ComputeButton 
      BackColor       =   &H00C0C0C0&
      Caption         =   "COMPUTE"
      Default         =   -1  'True
      Height          =   315
      Left            =   7320
      TabIndex        =   9
      Top             =   780
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Set Date && Time To"
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   7380
      TabIndex        =   16
      Top             =   60
      Width           =   1695
      Begin VB.CommandButton SetNowButton 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Now"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Time"
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   3120
      TabIndex        =   12
      Top             =   60
      Width           =   4215
      Begin VB.TextBox Hrs 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Min 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   600
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Sec 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Text            =   "00"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Set00HrButton 
         BackColor       =   &H00C0C0C0&
         Caption         =   "00 Hr"
         Height          =   315
         Left            =   2700
         TabIndex        =   6
         Top             =   240
         Width           =   675
      End
      Begin VB.CommandButton SetNoonButton 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Noon"
         Height          =   315
         Left            =   3420
         TabIndex        =   7
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Set to"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.Frame DateTime1Frame 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Date"
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   2895
      Begin VB.TextBox Year 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   2
         Text            =   "2001"
         Top             =   240
         Width           =   555
      End
      Begin VB.ComboBox MonthSelect 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "Earth.frx":0312
         Left            =   720
         List            =   "Earth.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   795
      End
      Begin VB.ComboBox DaySelect 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "Earth.frx":0316
         Left            =   120
         List            =   "Earth.frx":0318
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   25
      Top             =   7935
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11738
            MinWidth        =   176
            Text            =   "Local Star Time"
            TextSave        =   "Local Star Time"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1614
            MinWidth        =   1605
            TextSave        =   "21-7-2008"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   291
            MinWidth        =   282
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuAlmanac 
         Caption         =   "&Almanac"
      End
      Begin VB.Menu mnuSterrenkaart 
         Caption         =   "&StarChart"
      End
      Begin VB.Menu mnuJupiter 
         Caption         =   "&Jupiter"
         Begin VB.Menu mnuJupiterMoonsPhenomena 
            Caption         =   "&Phenomena"
         End
         Begin VB.Menu mnuJupiterMoons 
            Caption         =   "&Movements"
         End
         Begin VB.Menu mnuJupiterDiagram 
            Caption         =   "&Diagram"
         End
         Begin VB.Menu mnuEmpty1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuJupiterMeridiaan 
            Caption         =   "&Meridians"
         End
      End
      Begin VB.Menu mnuSaturnus 
         Caption         =   "&Saturn"
         Begin VB.Menu mnuSaturnusMoonsPhenomena 
            Caption         =   "&Phenomena"
         End
         Begin VB.Menu mnuSaturnusMoons 
            Caption         =   "&Movements"
         End
         Begin VB.Menu mnuSaturnusDiagram 
            Caption         =   "&Diagram"
         End
      End
      Begin VB.Menu mnuEmpty 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTabelen 
         Caption         =   "&Tabels"
      End
      Begin VB.Menu mnuSurvey 
         Caption         =   "&Survey "
      End
      Begin VB.Menu mnuEphemerides 
         Caption         =   "&Ephemerides"
      End
      Begin VB.Menu mnuEclipse 
         Caption         =   "E&clipse"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "&Info"
      End
   End
End
Attribute VB_Name = "frmPlanets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub ComputeButton_Click()

Dim SHelio As TSVECTOR, SGeo As TSVECTOR, SSun As TSVECTOR
    'Q1 = SHelio, Q2 = SGeo
Dim sAarde As TSVECTOR
Dim RA As Double
Dim Decl As Double
Dim RA1 As Double
Dim Decl1 As Double
Dim RA2 As Double
Dim Decl2 As Double
Dim dat As tDatum
Dim tt As Double
Dim T As Double
Dim T0 As Double 'tijdstip op 0h
Dim DtofUT As Double
Dim Obl As Double
Dim phase As Double
Dim PhaseAngle As Double
Dim ParAngle As Double
Dim Elongation As Double
Dim Magnitude As Double
Dim Semidiameter As Double
Dim PolarSemiDiameter As Double
Dim NutLon As Double, NutObl As Double
Dim Parallax As Double, MoonHeight As Double
Dim sSterbeeld As String
Dim JupiterPhysData As TJUPITERPHYSDATA
Dim MarsPhysData As TMARSPHYSDATA
Dim SunPhysData As TSUNPHYSDATA
Dim SaturnRingData As TSATURNRINGDATA
Dim AltSaturnRingData As TALTSATURNRINGDATA
Dim MoonPhysData As TMOONPHYSDATA
Dim C As Long
Dim JDOfCarr As Double
Dim deltaT As Double
Dim Az As Double, Alt As Double
Dim RTS As tRiseSetTran, RTS1 As tRiseSetTran, RTS2 As tRiseSetTran
Dim sLatitude As String, sLongitude As String
Dim JD_ZT As Double, JD_WT As Double

Call Zomertijd_Wintertijd(Year, JD_ZT, JD_WT)
dat.jj = Year
dat.mm = MonthSelect.ListIndex + 1
dat.DD = DaySelect
tt = (Hrs + Min / 60 + Sec / 3600) / 24
dat.DD = dat.DD + tt
T = JDToT(KalenderNaarJD(dat))
deltaT = ApproxDeltaT(T)
T0 = (floor(T * 36525 + 0.50001) - 0.5) / 36525 + TijdCorrectie(KalenderNaarJD(dat) + 0.2, JD_ZT, JD_WT) / 36525#
DtofUT = T0 + secToT * deltaT
T = T + TijdCorrectie(KalenderNaarJD(dat) + 0.2, JD_ZT, JD_WT) / 36525#
Call NutationConst(T, NutLon, NutObl)
Obl = Obliquity(T)

Dim LAST As Double, ObsLat As Double, ObsLon As Double, Height As Double
Dim RhoCosPhi As Double, RhoSinPhi As Double
Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Latitude", _
            sLatitude)
ObsLat = Val(sLatitude) * PI / 180
Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Longitude", _
            sLongitude)
ObsLon = Val(sLongitude) * PI / 180
LAST = SiderealTime(T) + NutLon * Cos(Obl) - ObsLon
Call ObserverCoord(ObsLat, Height, RhoCosPhi, RhoSinPhi)

'ObsLat = 42.3333 * pi / 180
'ObsLon = 71.0833 * pi / 180
Call ObserverCoord(ObsLat, Height, RhoCosPhi, RhoSinPhi)
List1.Clear
'======================== MAAN =======================

Dim Dist As Double, dkm As Double, illum As Double, diam As Double
Dim dist1 As Double, dkm1 As Double
Dim dist2 As Double, dkm2 As Double
Dim l As Double, B As Double
Dim sMoon As TSVECTOR
If chkGrootstePrecisie = 0 Then
    Call modMoonPos.MoonPos(T, sMoon)
    Call EclToEqu(sMoon.l, sMoon.B, Obl, RA, Decl)
Else
    Call Lune(TToJD(T), RA, Decl, Dist, dkm, diam, phase, illum)
    Call Lune(TToJD(T - Dist * LightTimeConst), RA, Decl, Dist, dkm, diam, phase, illum)
    RA = RA * PI / 12
    Decl = Decl * PI / 180
    'coordinaten zijn voor J2000. Omzetten naar huidige dag, en daarna appearent berekenen
    Call PrecessFK5(0, T, RA, Decl)
End If

Call Nutation(NutLon, NutObl, Obl, RA, Decl)
List1.AddItem "Maan          : " + StrHMS_DMS(RA * 180 / PI, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(Decl * 180 / PI, 7, 2, True, False, "g", 3)

Call EquToEcl(RA, Decl, Obl, SGeo.l, SGeo.B)
If chkGrootstePrecisie = 0 Then
    SGeo.R = sMoon.R
Else
    SGeo.R = dkm
End If
SHelio.l = 0: SHelio.B = 0: SHelio.R = 0
Call PlanetPosHi(0, T, sAarde, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SSun)
Call PlanetPosHi(0, T - SSun.R * LightTimeConst, sAarde, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SSun)
Call MoonPhysEphemeris(T, SGeo, SSun, Obl, NutLon, NutObl, MoonPhysData)
List1.AddItem "Position Angle: " + StrHMS_DMS(MoonPhysData.P * 180 / PI, 3, 0, True, False, "g", 3)
ParAngle = ParallacticAngle(RA, Decl, ObsLat, LAST)
List1.AddItem "Parall. Angle : " + StrHMS_DMS(ParAngle * 180 / PI, 3, 0, True, False, "g", 3)
List1.AddItem "Fase          : " + Format(MoonPhysData.k, "0.00000")
List1.AddItem "Bright Limb   : " + StrHMS_DMS(MoonPhysData.X * 180 / PI, 3, 0, True, False, "g", 3)
List1.AddItem "Terminator    : " + StrHMS_DMS(MoonPhysData.T * 180 / PI, 1, 1, True, False, "g", 5)
List1.AddItem "Libration in l: " + StrHMS_DMS(-MoonPhysData.l * 180 / PI, 1, 1, True, False, "g", 4)
List1.AddItem "Libration in b: " + StrHMS_DMS(MoonPhysData.B * 180 / PI, 1, 1, True, False, "g", 4)
List1.AddItem "Distance      : " + Format(SGeo.R, "000000.00 km")
Call SterBld(RA, Decl, 0#, sSterbeeld)
List1.AddItem "Sterrenbeeld  : " + sSterbeeld
Parallax = asin(EarthRadius / SGeo.R)
MoonHeight = MoonSetHeight(Parallax)

If chkGrootstePrecisie = 0 Then
    Call modMoonPos.MoonPos(T0, sMoon)
    Call EclToEqu(sMoon.l, sMoon.B, Obl, RA, Decl)
Else
    Call Lune(TToJD(T0), RA, Decl, Dist, dkm, diam, phase, illum)
    Call Lune(TToJD(T0 - Dist * LightTimeConst), RA, Decl, Dist, dkm, diam, phase, illum)
    RA = RA * PI / 12
    Decl = Decl * PI / 180
    'coordinaten zijn voor J2000. Omzetten naar huidige dag, en daarna appearent berekenen
    Call PrecessFK5(0, T0, RA, Decl)
End If

Call Nutation(NutLon, NutObl, Obl, RA, Decl)

If chkGrootstePrecisie = 0 Then
    Call modMoonPos.MoonPos(T0 - 1 / 36525, sMoon)
    Call EclToEqu(sMoon.l, sMoon.B, Obl, RA1, Decl1)
Else
    Call Lune(TToJD(T0 - 1 / 36525), RA1, Decl1, dist1, dkm1, diam, phase, illum)
    Call Lune(TToJD(T0 - dist1 * LightTimeConst - 1 / 36525), RA1, Decl1, dist1, dkm1, diam, phase, illum)
    RA1 = RA1 * PI / 12
    Decl1 = Decl1 * PI / 180
    'coordinaten zijn voor J2000. Omzetten naar huidige dag, en daarna appearent berekenen
    Call PrecessFK5(0, T0, RA1, Decl1)
End If
Call Nutation(NutLon, NutObl, Obl, RA1, Decl1)

If chkGrootstePrecisie = 0 Then
    Call modMoonPos.MoonPos(T0 + 1 / 36525, sMoon)
    Call EclToEqu(sMoon.l, sMoon.B, Obl, RA2, Decl2)
Else
    Call Lune(TToJD(T0 + 1 / 36525), RA2, Decl2, dist2, dkm2, diam, phase, illum)
    Call Lune(TToJD(T0 - dist1 * LightTimeConst + 1 / 36525), RA2, Decl2, dist2, dkm2, diam, phase, illum)
    RA2 = RA2 * PI / 12
    Decl2 = Decl2 * PI / 180
    'coordinaten zijn voor J2000. Omzetten naar huidige dag, en daarna appearent berekenen
    Call PrecessFK5(0, T0, RA2, Decl2)
End If

Call Nutation(NutLon, NutObl, Obl, RA2, Decl2)

Call RiseSet(T0, deltaT, RA1, Decl1, RA, Decl, RA2, Decl2, MoonHeight, ObsLon, ObsLat, RTS)
'DMO, 03-07-2008 indien flags aangaf dat
If RTS.flags > 0 Then
   RTS.Rise = -1
   RTS.Setting = -1
End If
If RTS.Rise < 0 Then
    List1.AddItem "Opkomst       : ------"
Else
    List1.AddItem "Opkomst       : " + StrHMS_DMS(RTS.Rise * 180 / PI, 3, 0, False, False, "h", 2)
End If
If RTS.Transit < 0 Then
    List1.AddItem "Doorgang      : ------"
Else
    List1.AddItem "Doorgang      : " + StrHMS_DMS(RTS.Transit * 180 / PI, 3, 0, False, False, "h", 2)
End If
If RTS.Setting < 0 Then
    List1.AddItem "Ondergang     : ------"
Else
    List1.AddItem "Ondergang     : " + StrHMS_DMS(RTS.Setting * 180 / PI, 3, 0, False, False, "h", 2)
End If

'======================== ZON ========================
List1.AddItem "------------------------------------------------------------------"
SHelio.l = 0: SHelio.B = 0: SHelio.R = 0
Call PlanetPosHi(0, T, sAarde, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(0, T - SGeo.R * LightTimeConst, sAarde, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
List1.AddItem "Zon           : " + StrHMS_DMS(180 / PI * RA, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / PI * Decl, 7, 2, True, False, "g", 3)
Call SunPhysEphemeris(T, SGeo.l, Obl, NutLon, SunPhysData)
C = CarringtonRotation(KalenderNaarJD(dat))
JDOfCarr = JDOfCarringtonRotation(C)
List1.AddItem "p             : " + StrHMS_DMS(SunPhysData.P * 180 / PI, 1, 1, True, False, "g", 4)
List1.AddItem "b0            : " + StrHMS_DMS(SunPhysData.b0 * 180 / PI, 1, 1, True, False, "g", 4)
List1.AddItem "l0            : " + StrHMS_DMS(SunPhysData.l0 * 180 / PI, 1, 1, False, False, "g", 4)
ParAngle = ParallacticAngle(RA, Decl, ObsLat, LAST)
List1.AddItem "Parall. Angle : " + StrHMS_DMS(ParAngle * 180 / PI, 3, 0, True, False, "g", 3)
Call SterBld(RA, Decl, 0#, sSterbeeld)
List1.AddItem "Sterrenbeeld  : " + sSterbeeld

Call ConvertVSOP_FK5(T, SGeo.l, SGeo.B)
Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA, Decl)
List1.AddItem "Appearent     : " + StrHMS_DMS(180 / PI * RA, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / PI * Decl, 7, 2, True, False, "g", 3)
Call EquToHor(RA, Decl, LAST, ObsLat, Az, Alt)
List1.AddItem "Azimuth/hoogte: " + StrHMS_DMS(180 / PI * Az, 3, 0, False, False, "g", 3) + vbTab + StrHMS_DMS(180 / PI * Alt, 3, 0, False, False, "g", 3)
        
' bepalen opkomst e.d.
Call PlanetPosHi(0, T0 - 1 / 36525, sAarde, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA1, Decl1)

Call PlanetPosHi(0, T0, sAarde, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)

Call PlanetPosHi(0, T0 + 1 / 36525, sAarde, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA2, Decl2)

Call RiseSet(T0, deltaT, RA1, Decl1, RA, Decl, RA2, Decl2, h0Sun, ObsLon, ObsLat, RTS)
'DMO, 03-07-2008 indien flags aangaf dat
If RTS.flags > 0 Then
   RTS.Rise = -1
   RTS.Setting = -1
End If
Call RiseSet(T0, deltaT, RA1, Decl1, RA, Decl, RA2, Decl2, -6 * DToR, ObsLon, ObsLat, RTS1)
'DMO, 03-07-2008 indien flags aangaf dat
If RTS1.flags > 0 Then
   RTS1.Rise = -1
   RTS1.Setting = -1
End If
Call RiseSet(T0, deltaT, RA1, Decl1, RA, Decl, RA2, Decl2, -18 * DToR, ObsLon, ObsLat, RTS2)
'DMO, 03-07-2008 indien flags aangaf dat
If RTS2.flags > 0 Then
   RTS2.Rise = -1
   RTS2.Setting = -1
End If
If RTS2.Rise < 0 Then
    List1.AddItem "Opkomst   -18 : ------"
Else
    List1.AddItem "Opkomst   -18 : " + StrHMS_DMS(RTS2.Rise * 180 / PI, 3, 0, False, False, "h", 2)
End If
If RTS1.Rise < 0 Then
    List1.AddItem "Opkomst   - 6 : ------"
Else
    List1.AddItem "Opkomst   - 6 : " + StrHMS_DMS(RTS1.Rise * 180 / PI, 3, 0, False, False, "h", 2)
End If
If RTS.Rise < 0 Then
    List1.AddItem "Opkomst       : ------"
Else
    List1.AddItem "Opkomst       : " + StrHMS_DMS(RTS.Rise * 180 / PI, 3, 0, False, False, "h", 2)
End If

If RTS.Transit < 0 Then
    List1.AddItem "Doorgang      : ------"
Else
    List1.AddItem "Doorgang      : " + StrHMS_DMS(RTS.Transit * 180 / PI, 3, 0, False, False, "h", 2)
End If

If RTS.Setting < 0 Then
    List1.AddItem "Ondergang     : ------"
Else
    List1.AddItem "Ondergang     : " + StrHMS_DMS(RTS.Setting * 180 / PI, 3, 0, False, False, "h", 2)
End If
If RTS1.Setting < 0 Then
    List1.AddItem "Ondergang + 6 : ------"
Else
    List1.AddItem "Ondergang + 6 : " + StrHMS_DMS(RTS1.Setting * 180 / PI, 3, 0, False, False, "h", 2)
End If
If RTS2.Setting < 0 Then
    List1.AddItem "Ondergang +18 : ------"
Else
    List1.AddItem "Ondergang +18 : " + StrHMS_DMS(RTS2.Setting * 180 / PI, 3, 0, False, False, "h", 2)
End If

'(INTERFACE_DATE, INTERFACE_TIME, Q1)
'List1.AddItem Q
List1.AddItem "------------------------------------------------------------------"
Call PlanetPosHi(0, T, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(1, T, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(1, T - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
List1.AddItem "Mercurius     : " + StrHMS_DMS(180 / PI * RA, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / PI * Decl, 7, 2, True, False, "g", 3)
Call ConvertVSOP_FK5(T, SGeo.l, SGeo.B)
Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA, Decl)
Call Aberration(T, Obl, FK5System, RA, Decl)
List1.AddItem "Appearent     : " + StrHMS_DMS(180 / PI * RA, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / PI * Decl, 7, 2, True, False, "g", 3)
Call EquToHor(RA, Decl, LAST, ObsLat, Az, Alt)
List1.AddItem "Azimuth/hoogte: " + StrHMS_DMS(180 / PI * Az, 3, 0, False, False, "g", 3) + vbTab + StrHMS_DMS(180 / PI * Alt, 3, 0, False, False, "g", 3)

phase = CalcPhase(SHelio.R, sAarde.R, SGeo.R)
PhaseAngle = acos(2 * phase - 1)
Elongation = CalcElongation(SHelio.R, sAarde.R, SGeo.R)
If modpi(SHelio.l - SGeo.l) > 0 Then Elongation = -Elongation
Magnitude = PlanetMagnitude(1, SHelio.R, SGeo.R, PhaseAngle, SaturnRingData.DeltaU, SaturnRingData.B)
Semidiameter = PlanetSemiDiameter(1, SGeo.R, PolarSemiDiameter)
List1.AddItem "Diameter      : " + StrHMS_DMS(2 * Semidiameter * SToR * RToD, 4, 1, True, False, "g", 4)
List1.AddItem "Magnitude     : " + Format(Magnitude, "0.0")
List1.AddItem "Phase         : " + Format(phase, "0.000")
List1.AddItem "PhaseAngle    : " + StrHMS_DMS(PhaseAngle * 180 / PI, 1, 1, True, False, "g", 4)
List1.AddItem "Elongation    : " + StrHMS_DMS(Elongation * 180 / PI, 1, 1, True, False, "g", 4)
ParAngle = ParallacticAngle(RA, Decl, ObsLat, LAST)
List1.AddItem "Parall. Angle : " + StrHMS_DMS(ParAngle * 180 / PI, 3, 0, True, False, "g", 3)
Call SterBld(RA, Decl, 0#, sSterbeeld)
List1.AddItem "Sterrenbeeld  : " + sSterbeeld

' bepalen opkomst e.d.
Call PlanetPosHi(0, T0 - 1 / 36525, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(1, T0 - 1 / 36525, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(1, T0 - 1 / 36525 - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA1, Decl1)

Call PlanetPosHi(0, T0, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(1, T0, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(1, T0 - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)

Call PlanetPosHi(0, T0 + 1 / 36525, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(1, T0 + 1 / 36525, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(1, T0 + 1 / 36525 - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA2, Decl2)

Call RiseSet(T0, deltaT, RA1, Decl1, RA, Decl, RA2, Decl2, h0Sun, ObsLon, ObsLat, RTS)
'DMO, 03-07-2008 indien flags aangaf dat
If RTS.flags > 0 Then
   RTS.Rise = -1
   RTS.Setting = -1
End If
List1.AddItem "Opkomst       : " + StrHMS_DMS(RTS.Rise * 180 / PI, 3, 0, False, False, "h", 2)
List1.AddItem "Doorgang      : " + StrHMS_DMS(RTS.Transit * 180 / PI, 3, 0, False, False, "h", 2)
List1.AddItem "Ondergang     : " + StrHMS_DMS(RTS.Setting * 180 / PI, 3, 0, False, False, "h", 2)

List1.AddItem "------------------------------------------------------------------"
Call PlanetPosHi(0, T, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(2, T, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(0, T - SGeo.R * LightTimeConst, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(2, T - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
List1.AddItem "Venus         : " + StrHMS_DMS(180 / PI * RA, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / PI * Decl, 7, 2, True, False, "g", 3)
Call ConvertVSOP_FK5(T, SGeo.l, SGeo.B)
Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA, Decl)
List1.AddItem "Appearent     : " + StrHMS_DMS(180 / PI * RA, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / PI * Decl, 7, 2, True, False, "g", 3)
Call EquToHor(RA, Decl, LAST, ObsLat, Az, Alt)
List1.AddItem "Azimuth/hoogte: " + StrHMS_DMS(180 / PI * Az, 3, 0, False, False, "g", 3) + vbTab + StrHMS_DMS(180 / PI * Alt, 3, 0, False, False, "g", 3)

'Call Aberration(T, Obl, FK5System, RA, Decl)
phase = CalcPhase(SHelio.R, sAarde.R, SGeo.R)
PhaseAngle = acos(2 * phase - 1)
Elongation = CalcElongation(SHelio.R, sAarde.R, SGeo.R)
If modpi(SHelio.l - SGeo.l) > 0 Then Elongation = -Elongation
Magnitude = PlanetMagnitude(2, SHelio.R, SGeo.R, PhaseAngle, SaturnRingData.DeltaU, SaturnRingData.B)
Semidiameter = PlanetSemiDiameter(2, SGeo.R, PolarSemiDiameter)
List1.AddItem "Diameter      : " + StrHMS_DMS(2 * Semidiameter * SToR * RToD, 4, 1, True, False, "g", 4)
List1.AddItem "Magnitude     : " + Format(Magnitude, "0.0")
List1.AddItem "Phase         : " + Format(phase, "0.000")
List1.AddItem "PhaseAngle    : " + StrHMS_DMS(PhaseAngle * 180 / PI, 1, 1, True, False, "g", 4)
List1.AddItem "Elongation    : " + StrHMS_DMS(Elongation * 180 / PI, 1, 1, True, False, "g", 4)
ParAngle = ParallacticAngle(RA, Decl, ObsLat, LAST)
List1.AddItem "Parall. Angle : " + StrHMS_DMS(ParAngle * 180 / PI, 3, 0, True, False, "g", 3)
Call SterBld(RA, Decl, 0#, sSterbeeld)
List1.AddItem "Sterrenbeeld  : " + sSterbeeld

' bepalen opkomst e.d.
Call PlanetPosHi(0, T0 - 1 / 36525, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(2, T0 - 1 / 36525, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(2, T0 - 1 / 36525 - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA1, Decl1)
Call Aberration(T0, Obl, FK5System, RA1, Decl1)

Call PlanetPosHi(0, T0, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(2, T0, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(2, T0 - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obliquity(T0), RA, Decl)
Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA, Decl)
Call Aberration(T0, Obliquity(T0 - 1 / 36525), FK5System, RA, Decl)

Call PlanetPosHi(0, T0 + 1 / 36525, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(2, T0 + 1 / 36525, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(2, T0 + 1 / 36525 - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA2, Decl2)

Call RiseSet(T0, deltaT, RA1, Decl1, RA, Decl, RA2, Decl2, h0Planet, ObsLon, ObsLat, RTS)
'DMO, 03-07-2008 indien flags aangaf dat
If RTS.flags > 0 Then
   RTS.Rise = -1
   RTS.Setting = -1
End If
List1.AddItem "Opkomst       : " + StrHMS_DMS(RTS.Rise * 180 / PI, 3, 0, False, False, "h", 2)
List1.AddItem "Doorgang      : " + StrHMS_DMS(RTS.Transit * 180 / PI, 3, 0, False, False, "h", 2)
List1.AddItem "Ondergang     : " + StrHMS_DMS(RTS.Setting * 180 / PI, 3, 0, False, False, "h", 2)




List1.AddItem "------------------------------------------------------------------"
Call PlanetPosHi(0, T, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(4, T, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(4, T - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
List1.AddItem "Mars          : " + StrHMS_DMS(180 / PI * RA, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / PI * Decl, 7, 2, True, False, "g", 3)
phase = CalcPhase(SHelio.R, sAarde.R, SGeo.R)
PhaseAngle = acos(2 * phase - 1)
Elongation = CalcElongation(SHelio.R, sAarde.R, SGeo.R)
If modpi(SHelio.l - SGeo.l) > 0 Then Elongation = -Elongation
Magnitude = PlanetMagnitude(4, SHelio.R, SGeo.R, PhaseAngle, SaturnRingData.DeltaU, SaturnRingData.B)
Semidiameter = PlanetSemiDiameter(4, SGeo.R, PolarSemiDiameter)
List1.AddItem "Diameter      : " + StrHMS_DMS(2 * Semidiameter * SToR * RToD, 4, 1, True, False, "g", 4)
List1.AddItem "Magnitude     : " + Format(Magnitude, "0.0")
List1.AddItem "Phase         : " + Format(phase, "0.000")
List1.AddItem "PhaseAngle    : " + StrHMS_DMS(PhaseAngle * 180 / PI, 1, 1, False, False, "g", 4)
List1.AddItem "Elongation    : " + StrHMS_DMS(Elongation * 180 / PI, 1, 1, False, False, "g", 4)
ParAngle = ParallacticAngle(RA, Decl, ObsLat, LAST)
Call ConvertVSOP_FK5(T, SGeo.l, SGeo.B)
Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA, Decl)
List1.AddItem "Appearent     : " + StrHMS_DMS(180 / PI * RA, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / PI * Decl, 7, 2, True, False, "g", 3)
Call EquToHor(RA, Decl, LAST, ObsLat, Az, Alt)
List1.AddItem "Azimuth/hoogte: " + StrHMS_DMS(180 / PI * Az, 3, 0, False, False, "g", 3) + vbTab + StrHMS_DMS(180 / PI * Alt, 3, 0, False, False, "g", 3)
      
List1.AddItem "Parall. Angle : " + StrHMS_DMS(ParAngle * 180 / PI, 3, 0, True, False, "g", 3)
Call SterBld(RA, Decl, 0#, sSterbeeld)
List1.AddItem "Sterrenbeeld  : " + sSterbeeld
Call MarsPhysEphemeris(T, SHelio, sAarde, SGeo, _
                            Obl, NutLon, NutObl, _
                            MarsPhysData)
List1.AddItem "DS            : " + StrHMS_DMS(MarsPhysData.DS * 180 / PI, 1, 2, True, False, "g", 3)
List1.AddItem "DE            : " + StrHMS_DMS(MarsPhysData.DE * 180 / PI, 1, 2, True, False, "g", 3)
List1.AddItem "P             : " + StrHMS_DMS(MarsPhysData.P * 180 / PI, 1, 2, True, False, "g", 3)
List1.AddItem "q             : " + StrHMS_DMS(MarsPhysData.qq / 3600, 4, 2, True, False, "g", 4)
List1.AddItem "Q             : " + StrHMS_DMS(MarsPhysData.Q * 180 / PI, 1, 2, True, False, "g", 3)
List1.AddItem "Om            : " + StrHMS_DMS(MarsPhysData.Om * 180 / PI, 1, 2, False, False, "g", 3)

' bepalen opkomst e.d.
Call PlanetPosHi(0, T0 - 1 / 36525, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(4, T0 - 1 / 36525, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(4, T0 - 1 / 36525 - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA1, Decl1)
Call Aberration(T0, Obl, FK5System, RA1, Decl1)

Call PlanetPosHi(0, T0, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(4, T0, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(4, T0 - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obliquity(T0), RA, Decl)
Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA, Decl)
Call Aberration(T0, Obliquity(T0 - 1 / 36525), FK5System, RA, Decl)

Call PlanetPosHi(0, T0 + 1 / 36525, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(4, T0 + 1 / 36525, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(4, T0 + 1 / 36525 - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA2, Decl2)

Call RiseSet(T0, deltaT, RA1, Decl1, RA, Decl, RA2, Decl2, h0Planet, ObsLon, ObsLat, RTS)
'DMO, 03-07-2008 indien flags aangaf dat
If RTS.flags > 0 Then
   RTS.Rise = -1
   RTS.Setting = -1
End If
List1.AddItem "Opkomst       : " + StrHMS_DMS(RTS.Rise * 180 / PI, 3, 0, False, False, "h", 2)
List1.AddItem "Doorgang      : " + StrHMS_DMS(RTS.Transit * 180 / PI, 3, 0, False, False, "h", 2)
List1.AddItem "Ondergang     : " + StrHMS_DMS(RTS.Setting * 180 / PI, 3, 0, False, False, "h", 2)


List1.AddItem "------------------------------------------------------------------"
Call PlanetPosHi(0, T, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(5, T, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(5, T - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
List1.AddItem "Jupiter       : " + StrHMS_DMS(180 / PI * RA, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / PI * Decl, 7, 2, True, False, "g", 3)
phase = CalcPhase(SHelio.R, sAarde.R, SGeo.R)
PhaseAngle = acos(2 * phase - 1)
Elongation = CalcElongation(SHelio.R, sAarde.R, SGeo.R)
If modpi(SHelio.l - SGeo.l) > 0 Then Elongation = -Elongation
Magnitude = PlanetMagnitude(5, SHelio.R, SGeo.R, PhaseAngle, SaturnRingData.DeltaU, SaturnRingData.B)
Semidiameter = PlanetSemiDiameter(5, SGeo.R, PolarSemiDiameter)
List1.AddItem "Diameter      : " + StrHMS_DMS(2 * Semidiameter * SToR * RToD, 4, 1, True, False, "g", 4)
List1.AddItem "Polardiameter : " + StrHMS_DMS(2 * PolarSemiDiameter * SToR * RToD, 4, 1, True, False, "g", 4)
List1.AddItem "Magnitude     : " + Format(Magnitude, "0.0")
List1.AddItem "Phase         : " + Format(phase, "0.000")
List1.AddItem "PhaseAngle    : " + StrHMS_DMS(PhaseAngle * 180 / PI, 1, 1, True, False, "g", 4)
List1.AddItem "Elongation    : " + StrHMS_DMS(Elongation * 180 / PI, 1, 1, True, False, "g", 4)
ParAngle = ParallacticAngle(RA, Decl, ObsLat, LAST)
List1.AddItem "Parall. Angle : " + StrHMS_DMS(ParAngle * 180 / PI, 3, 0, True, False, "g", 3)
Call ConvertVSOP_FK5(T, SGeo.l, SGeo.B)
Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA, Decl)
List1.AddItem "Appearent     : " + StrHMS_DMS(180 / PI * RA, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / PI * Decl, 7, 2, True, False, "g", 3)
Call EquToHor(RA, Decl, LAST, ObsLat, Az, Alt)
List1.AddItem "Azimuth/hoogte: " + StrHMS_DMS(180 / PI * Az, 3, 0, False, False, "g", 3) + vbTab + StrHMS_DMS(180 / PI * Alt, 3, 0, False, False, "g", 3)
      
Call SterBld(RA, Decl, 0#, sSterbeeld)
List1.AddItem "Sterrenbeeld  : " + sSterbeeld

Call JupiterPhysEphemeris(T + deltaT / 36525 / 86400, SHelio, sAarde, SGeo, _
                               Obl, NutLon, NutObl, _
                              JupiterPhysData)
List1.AddItem "DS            : " + StrHMS_DMS(JupiterPhysData.DS * 180 / PI, 1, 2, True, False, "g", 3)
List1.AddItem "DE            : " + StrHMS_DMS(JupiterPhysData.DE * 180 / PI, 1, 2, True, False, "g", 3)
List1.AddItem "Om1           : " + StrHMS_DMS(JupiterPhysData.Om1 * 180 / PI, 1, 2, False, False, "g", 3)
List1.AddItem "Om2           : " + StrHMS_DMS(JupiterPhysData.Om2 * 180 / PI, 1, 2, False, False, "g", 3)
List1.AddItem "P             : " + StrHMS_DMS(JupiterPhysData.P * 180 / PI, 1, 2, True, False, "g", 3)
          
Call PlanetPosHi(0, T0, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(5, T0, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(5, T0 - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obliquity(T0), RA, Decl)
Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA, Decl)
Call Aberration(T0, Obliquity(T0 - 1 / 36525), FK5System, RA, Decl)

Call PlanetPosHi(0, T0 + 1 / 36525, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(5, T0 + 1 / 36525, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(5, T0 + 1 / 36525 - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA2, Decl2)

Call RiseSet(T0, deltaT, RA1, Decl1, RA, Decl, RA2, Decl2, h0Planet, ObsLon, ObsLat, RTS)
'DMO, 03-07-2008 indien flags aangaf dat
If RTS.flags > 0 Then
   RTS.Rise = -1
   RTS.Setting = -1
End If
List1.AddItem "Opkomst       : " + StrHMS_DMS(RTS.Rise * 180 / PI, 3, 0, False, False, "h", 2)
List1.AddItem "Doorgang      : " + StrHMS_DMS(RTS.Transit * 180 / PI, 3, 0, False, False, "h", 2)
List1.AddItem "Ondergang     : " + StrHMS_DMS(RTS.Setting * 180 / PI, 3, 0, False, False, "h", 2)

          
List1.AddItem "------------------------------------------------------------------"
Call PlanetPosHi(0, T, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(6, T, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(6, T - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
List1.AddItem "Saturnus      : " + StrHMS_DMS(180 / PI * RA, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / PI * Decl, 7, 2, True, False, "g", 3)
Call SaturnRing(T, SHelio, SGeo, Obl, NutLon, NutObl, SaturnRingData)
Call AltSaturnRing(T, SHelio, SGeo, Obl, NutLon, NutObl, AltSaturnRingData)
phase = CalcPhase(SHelio.R, sAarde.R, SGeo.R)
PhaseAngle = acos(2 * phase - 1)
Elongation = CalcElongation(SHelio.R, sAarde.R, SGeo.R)
If modpi(SHelio.l - SGeo.l) > 0 Then Elongation = -Elongation
Magnitude = PlanetMagnitude(6, SHelio.R, SGeo.R, PhaseAngle, SaturnRingData.DeltaU, SaturnRingData.B)
Semidiameter = PlanetSemiDiameter(6, SGeo.R, PolarSemiDiameter)
List1.AddItem "Diameter      : " + StrHMS_DMS(2 * Semidiameter * SToR * RToD, 4, 1, True, False, "g", 4)
List1.AddItem "PolarDiameter : " + StrHMS_DMS(2 * PolarSemiDiameter * SToR * RToD, 4, 1, True, False, "g", 4)
Call CorrectSaturnSemiDiameter(SaturnRingData.B, PolarSemiDiameter)
List1.AddItem "PolarDiameter : " + StrHMS_DMS(2 * PolarSemiDiameter * SToR * RToD, 4, 1, True, False, "g", 4) & " gecorrigeerd voor Earth Lat. B"
List1.AddItem "Magnitude     : " + Format(Magnitude, "0.0")
List1.AddItem "Phase         : " + Format(phase, "0.000")
List1.AddItem "PhaseAngle    : " + StrHMS_DMS(PhaseAngle * 180 / PI, 1, 1, True, False, "g", 4)
List1.AddItem "Elongation    : " + StrHMS_DMS(Elongation * 180 / PI, 1, 1, True, False, "g", 4)
ParAngle = ParallacticAngle(RA, Decl, ObsLat, LAST)
List1.AddItem "Parall. Angle : " + StrHMS_DMS(ParAngle * 180 / PI, 3, 0, True, False, "g", 3)
Call SterBld(RA, Decl, 0#, sSterbeeld)
List1.AddItem "Sterrenbeeld  : " + sSterbeeld
Call ConvertVSOP_FK5(T, SGeo.l, SGeo.B)
Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA, Decl)
List1.AddItem "Appearent     : " + StrHMS_DMS(180 / PI * RA, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / PI * Decl, 7, 2, True, False, "g", 3)
Call EquToHor(RA, Decl, LAST, ObsLat, Az, Alt)
List1.AddItem "Azimuth/hoogte: " + StrHMS_DMS(180 / PI * Az, 3, 0, False, False, "g", 3) + vbTab + StrHMS_DMS(180 / PI * Alt, 3, 0, False, False, "g", 3)
      
List1.AddItem "B             : " + StrHMS_DMS(SaturnRingData.B * 180 / PI, 1, 3, True, False, "g", 4)
List1.AddItem "Bd            : " + StrHMS_DMS(SaturnRingData.Bd * 180 / PI, 1, 3, True, False, "g", 4)
List1.AddItem "dU            : " + StrHMS_DMS(SaturnRingData.DeltaU * 180 / PI, 1, 3, True, False, "g", 4)
List1.AddItem "P             : " + StrHMS_DMS(SaturnRingData.P * 180 / PI, 1, 3, True, False, "g", 4)
List1.AddItem "B             : " + StrHMS_DMS(AltSaturnRingData.B * 180 / PI, 1, 2, True, False, "g", 3)
List1.AddItem "B1            : " + StrHMS_DMS(AltSaturnRingData.b1 * 180 / PI, 1, 2, True, False, "g", 3)
List1.AddItem "U             : " + StrHMS_DMS(AltSaturnRingData.u * 180 / PI, 1, 2, True, False, "g", 3)
List1.AddItem "U1            : " + StrHMS_DMS(AltSaturnRingData.U1 * 180 / PI, 1, 2, True, False, "g", 3)
List1.AddItem "P             : " + StrHMS_DMS(AltSaturnRingData.P * 180 / PI, 1, 2, True, False, "g", 3)
List1.AddItem "P1            : " + StrHMS_DMS(AltSaturnRingData.P1 * 180 / PI, 1, 2, True, False, "g", 3)

List1.AddItem "A: " & vbTab & "Axis" & vbTab & "ioAxis" & vbTab & "oiAxis" & vbTab & "iiAxis" & vbTab & "idAxis"
List1.AddItem vbTab & StrHMS_DMS(SaturnRingData.aAxis / 3600, 4, 1, True, False, "g", 3) & vbTab & _
     StrHMS_DMS(SaturnRingData.ioaAxis / 3600, 4, 1, True, False, "g", 3) & vbTab & _
     StrHMS_DMS(SaturnRingData.oiaAxis / 3600, 4, 1, True, False, "g", 3) & vbTab & _
     StrHMS_DMS(SaturnRingData.iiaAxis / 3600, 4, 1, True, False, "g", 3) & vbTab & _
     StrHMS_DMS(SaturnRingData.idaAxis / 3600, 4, 1, True, False, "g", 3)
List1.AddItem "B: " & vbTab & "Axis" & vbTab & "ioAxis" & vbTab & "oiAxis" & vbTab & "iiAxis" & vbTab & "idAxis"
List1.AddItem vbTab & StrHMS_DMS(SaturnRingData.bAxis / 3600, 4, 1, True, False, "g", 3) & vbTab & _
     StrHMS_DMS(SaturnRingData.iobAxis / 3600, 4, 1, True, False, "g", 3) & vbTab & _
     StrHMS_DMS(SaturnRingData.oibAxis / 3600, 4, 1, True, False, "g", 3) & vbTab & _
     StrHMS_DMS(SaturnRingData.iibAxis / 3600, 4, 1, True, False, "g", 3) & vbTab & _
     StrHMS_DMS(SaturnRingData.idbAxis / 3600, 4, 1, True, False, "g", 3)
'List1.AddItem "A-Axis        : " + StrHMS_DMS(SaturnRingData.aAxis / 3600, 4, 2, True, "g", 3)
'List1.AddItem "B-Axis        : " + StrHMS_DMS(SaturnRingData.bAxis / 3600, 4, 2, True, "g", 3)
'List1.AddItem "idaAxis       : " + StrHMS_DMS(SaturnRingData.idaAxis / 3600, 4, 2, True, "g", 3)
'List1.AddItem "idbAxis       : " + StrHMS_DMS(SaturnRingData.idbAxis / 3600, 4, 2, True, "g", 3)
'List1.AddItem "iiaAxis       : " + StrHMS_DMS(SaturnRingData.iiaAxis / 3600, 4, 2, True, "g", 3)
'List1.AddItem "iibAxis       : " + StrHMS_DMS(SaturnRingData.iibAxis / 3600, 4, 2, True, "g", 3)
'List1.AddItem "ioaAxis       : " + StrHMS_DMS(SaturnRingData.ioaAxis / 3600, 4, 2, True, "g", 3)
'List1.AddItem "iobAxis       : " + StrHMS_DMS(SaturnRingData.iobAxis / 3600, 4, 2, True, "g", 3)
'List1.AddItem "oiaAxis       : " + StrHMS_DMS(SaturnRingData.oiaAxis / 3600, 4, 2, True, "g", 3)
'List1.AddItem "oibAxis       : " + StrHMS_DMS(SaturnRingData.oibAxis / 3600, 4, 2, True, "g", 3)

' bepalen opkomst e.d.
Call PlanetPosHi(0, T0 - 1 / 36525, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(6, T0 - 1 / 36525, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(6, T0 - 1 / 36525 - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA1, Decl1)
Call Aberration(T0, Obl, FK5System, RA1, Decl1)

Call PlanetPosHi(0, T0, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(6, T0, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(6, T0 - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obliquity(T0), RA, Decl)
Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA, Decl)
Call Aberration(T0, Obliquity(T0 - 1 / 36525), FK5System, RA, Decl)

Call PlanetPosHi(0, T0 + 1 / 36525, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(6, T0 + 1 / 36525, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(6, T0 + 1 / 36525 - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA2, Decl2)

Call RiseSet(T0, deltaT, RA1, Decl1, RA, Decl, RA2, Decl2, h0Planet, ObsLon, ObsLat, RTS)
'DMO, 03-07-2008 indien flags aangaf dat
If RTS.flags > 0 Then
   RTS.Rise = -1
   RTS.Setting = -1
End If
List1.AddItem "Opkomst       : " + StrHMS_DMS(RTS.Rise * 180 / PI, 3, 0, False, False, "h", 2)
List1.AddItem "Doorgang      : " + StrHMS_DMS(RTS.Transit * 180 / PI, 3, 0, False, False, "h", 2)
List1.AddItem "Ondergang     : " + StrHMS_DMS(RTS.Setting * 180 / PI, 3, 0, False, False, "h", 2)

List1.AddItem "------------------------------------------------------------------"
Call PlanetPosHi(0, T, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(7, T, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(7, T - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
List1.AddItem "Uranus        : " + StrHMS_DMS(180 / PI * RA, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / PI * Decl, 7, 2, True, False, "g", 3)
phase = CalcPhase(SHelio.R, sAarde.R, SGeo.R)
PhaseAngle = acos(2 * phase - 1)
Elongation = CalcElongation(SHelio.R, sAarde.R, SGeo.R)
If modpi(SHelio.l - SGeo.l) > 0 Then Elongation = -Elongation
Magnitude = PlanetMagnitude(7, SHelio.R, SGeo.R, PhaseAngle, SaturnRingData.DeltaU, SaturnRingData.B)
Semidiameter = PlanetSemiDiameter(7, SGeo.R, PolarSemiDiameter)
List1.AddItem "Diameter      : " + StrHMS_DMS(2 * Semidiameter * SToR * RToD, 4, 1, True, False, "g", 4)
List1.AddItem "Magnitude     : " + Format(Magnitude, "0.0")
List1.AddItem "Phase         : " + Format(phase, "0.000")
List1.AddItem "PhaseAngle    : " + StrHMS_DMS(PhaseAngle * 180 / PI, 1, 1, True, False, "g", 4)
List1.AddItem "Elongation    : " + StrHMS_DMS(Elongation * 180 / PI, 1, 1, True, False, "g", 4)
ParAngle = ParallacticAngle(RA, Decl, ObsLat, LAST)
List1.AddItem "Parall. Angle : " + StrHMS_DMS(ParAngle * 180 / PI, 3, 0, True, False, "g", 3)
Call SterBld(RA, Decl, 0#, sSterbeeld)
List1.AddItem "Sterrenbeeld  : " + sSterbeeld
Call ConvertVSOP_FK5(T, SGeo.l, SGeo.B)
Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA, Decl)
List1.AddItem "Appearent     : " + StrHMS_DMS(180 / PI * RA, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / PI * Decl, 7, 2, True, False, "g", 3)
Call EquToHor(RA, Decl, LAST, ObsLat, Az, Alt)
List1.AddItem "Azimuth/hoogte: " + StrHMS_DMS(180 / PI * Az, 3, 0, False, False, "g", 3) + vbTab + StrHMS_DMS(180 / PI * Alt, 3, 0, False, False, "g", 3)
      
' bepalen opkomst e.d.
Call PlanetPosHi(0, T0 - 1 / 36525, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(7, T0 - 1 / 36525, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(7, T0 - 1 / 36525 - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA1, Decl1)
Call Aberration(T0, Obl, FK5System, RA1, Decl1)

Call PlanetPosHi(0, T0, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(7, T0, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(7, T0 - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obliquity(T0), RA, Decl)
Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA, Decl)
Call Aberration(T0, Obliquity(T0 - 1 / 36525), FK5System, RA, Decl)

Call PlanetPosHi(0, T0 + 1 / 36525, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(7, T0 + 1 / 36525, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(7, T0 + 1 / 36525 - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA2, Decl2)

Call RiseSet(T0, deltaT, RA1, Decl1, RA, Decl, RA2, Decl2, h0Planet, ObsLon, ObsLat, RTS)
'DMO, 03-07-2008 indien flags aangaf dat
If RTS.flags > 0 Then
   RTS.Rise = -1
   RTS.Setting = -1
End If
List1.AddItem "Opkomst       : " + StrHMS_DMS(RTS.Rise * 180 / PI, 3, 0, False, False, "h", 2)
List1.AddItem "Doorgang      : " + StrHMS_DMS(RTS.Transit * 180 / PI, 3, 0, False, False, "h", 2)
List1.AddItem "Ondergang     : " + StrHMS_DMS(RTS.Setting * 180 / PI, 3, 0, False, False, "h", 2)

List1.AddItem "------------------------------------------------------------------"
Call PlanetPosHi(0, T, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(8, T, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(8, T - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
List1.AddItem "Neptunus      : " + StrHMS_DMS(180 / PI * RA, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / PI * Decl, 7, 2, True, False, "g", 3)
phase = CalcPhase(SHelio.R, sAarde.R, SGeo.R)
PhaseAngle = acos(2 * phase - 1)
Elongation = CalcElongation(SHelio.R, sAarde.R, SGeo.R)
If modpi(SHelio.l - SGeo.l) > 0 Then Elongation = -Elongation
Magnitude = PlanetMagnitude(8, SHelio.R, SGeo.R, PhaseAngle, SaturnRingData.DeltaU, SaturnRingData.B)
Semidiameter = PlanetSemiDiameter(8, SGeo.R, PolarSemiDiameter)
List1.AddItem "Diameter      : " + StrHMS_DMS(2 * Semidiameter * SToR * RToD, 4, 1, True, False, "g", 4)
List1.AddItem "Magnitude     : " + Format(Magnitude, "0.0")
List1.AddItem "Phase         : " + Format(phase, "0.000")
List1.AddItem "PhaseAngle    : " + StrHMS_DMS(PhaseAngle * 180 / PI, 1, 1, True, False, "g", 4)
List1.AddItem "Elongation    : " + StrHMS_DMS(Elongation * 180 / PI, 1, 1, True, False, "g", 4)
ParAngle = ParallacticAngle(RA, Decl, ObsLat, LAST)
List1.AddItem "Parall. Angle : " + StrHMS_DMS(ParAngle * 180 / PI, 3, 0, True, False, "g", 3)
Call SterBld(RA, Decl, 0#, sSterbeeld)
List1.AddItem "Sterrenbeeld  : " + sSterbeeld
Call ConvertVSOP_FK5(T, SGeo.l, SGeo.B)
Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA, Decl)
List1.AddItem "Appearent     : " + StrHMS_DMS(180 / PI * RA, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / PI * Decl, 7, 2, True, False, "g", 3)
Call EquToHor(RA, Decl, LAST, ObsLat, Az, Alt)
List1.AddItem "Azimuth/hoogte: " + StrHMS_DMS(180 / PI * Az, 3, 0, False, False, "g", 3) + vbTab + StrHMS_DMS(180 / PI * Alt, 3, 0, False, False, "g", 3)
      
' bepalen opkomst e.d.
Call PlanetPosHi(0, T0 - 1 / 36525, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(8, T0 - 1 / 36525, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(8, T0 - 1 / 36525 - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA1, Decl1)
Call Aberration(T0, Obl, FK5System, RA1, Decl1)

Call PlanetPosHi(0, T0, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(8, T0, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(8, T0 - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obliquity(T0), RA, Decl)
Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA, Decl)
Call Aberration(T0, Obliquity(T0 - 1 / 36525), FK5System, RA, Decl)

Call PlanetPosHi(0, T0 + 1 / 36525, sAarde, chkGrootstePrecisie.Value = 0)
Call PlanetPosHi(8, T0 + 1 / 36525, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call PlanetPosHi(8, T0 + 1 / 36525 - SGeo.R * LightTimeConst, SHelio, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA2, Decl2)

Call RiseSet(T0, deltaT, RA1, Decl1, RA, Decl, RA2, Decl2, h0Planet, ObsLon, ObsLat, RTS)
'DMO, 03-07-2008 indien flags aangaf dat
If RTS.flags > 0 Then
   RTS.Rise = -1
   RTS.Setting = -1
End If
List1.AddItem "Opkomst       : " + StrHMS_DMS(RTS.Rise * 180 / PI, 3, 0, False, False, "h", 2)
List1.AddItem "Doorgang      : " + StrHMS_DMS(RTS.Transit * 180 / PI, 3, 0, False, False, "h", 2)
List1.AddItem "Ondergang     : " + StrHMS_DMS(RTS.Setting * 180 / PI, 3, 0, False, False, "h", 2)

List1.AddItem "------------------------------------------------------------------"
'================== PLUTO ==============
'Dit is een speciaal geval. Coordinaten zijn voor 2000. Deze moeten voor de Zon worden berekend.
'Dat is: bereken positie voor vandaag. De coordinaten omzetten naar J2000
'dit is niet voldoende. Er moet nog een correctie plaatsvinden van Ecl. VSOP -> equ FK5-2000
'vervolgens de positie van pluto berekenen
Dim TAarde As TVECTOR
Dim sZon As TSVECTOR
Dim TPluto As TVECTOR
SHelio.l = 0: SHelio.B = 0: SHelio.R = 0
Call PlanetPosHi(0, T, sAarde, chkGrootstePrecisie.Value = 0)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call SphToRect(SGeo, TAarde)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
' Call Reduction2000(0, RA, Decl)
'coordinaten omzetten naar J2000
Call PrecessFK5(T, 0, RA, Decl)
Call EquToEcl(RA, Decl, Obliquity(0), SGeo.l, SGeo.B)
Call SphToRect(SGeo, TAarde)
Call EclVSOP2000_equFK52000(TAarde.X, TAarde.Y, TAarde.Z)
Call RectToSph(TAarde, sZon)
sAarde = SGeo

Call PlanetPosHi(0, T, sAarde, chkGrootstePrecisie.Value = 0)
Call PlutoPos(T, SHelio)
Call EclToRect(SHelio, Obliquity(0), TPluto)
Dist = Sqr((TAarde.X + TPluto.X) * (TAarde.X + TPluto.X) + (TAarde.Y + TPluto.Y) * (TAarde.Y + TPluto.Y) + (TAarde.Z + TPluto.Z) * (TAarde.Z + TPluto.Z))
Call PlutoPos(T - Dist * LightTimeConst, SHelio)
Call EclToRect(SHelio, Obliquity(0), TPluto)
Dist = Sqr((TAarde.X + TPluto.X) * (TAarde.X + TPluto.X) + (TAarde.Y + TPluto.Y) * (TAarde.Y + TPluto.Y) + (TAarde.Z + TPluto.Z) * (TAarde.Z + TPluto.Z))
RA = atan2(TPluto.Y + TAarde.Y, TPluto.X + TAarde.X)
If RA < 0 Then
    RA = RA + Pi2
End If
Decl = asin((TPluto.Z + TAarde.Z) / Dist)
List1.AddItem "Pluto         : " + StrHMS_DMS(180 / PI * RA, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / PI * Decl, 7, 2, True, False, "g", 3)
phase = CalcPhase(SHelio.R, sAarde.R, Dist)
PhaseAngle = acos(2 * phase - 1)
Magnitude = PlanetMagnitude(9, SHelio.R, Dist, PhaseAngle, SaturnRingData.DeltaU, SaturnRingData.B)
Semidiameter = PlanetSemiDiameter(9, Dist, PolarSemiDiameter)
List1.AddItem "Diameter      : " + StrHMS_DMS(2 * Semidiameter * SToR * RToD, 4, 2, True, False, "g", 4)
List1.AddItem "Magnitude     : " + Format(Magnitude, "0.0")
ParAngle = ParallacticAngle(RA, Decl, ObsLat, LAST)
List1.AddItem "Parall. Angle : " + StrHMS_DMS(ParAngle * 180 / PI, 3, 0, True, False, "g", 3)
Call SterBld(RA, Decl, 0#, sSterbeeld)
List1.AddItem "Sterrenbeeld  : " + sSterbeeld
Call EquToEcl(RA, Decl, Obl, SGeo.l, SGeo.B)
Call ConvertVSOP_FK5(T, SGeo.l, SGeo.B)
Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA, Decl)
List1.AddItem "Appearent     : " + StrHMS_DMS(180 / PI * RA, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / PI * Decl, 7, 2, True, False, "g", 3)
Call EquToHor(RA, Decl, LAST, ObsLat, Az, Alt)
List1.AddItem "Azimuth/hoogte: " + StrHMS_DMS(180 / PI * Az, 3, 0, False, False, "g", 3) + vbTab + StrHMS_DMS(180 / PI * Alt, 3, 0, False, False, "g", 3)

Call RiseSet(T0, deltaT, RA, Decl, RA, Decl, RA, Decl, h0Planet, ObsLon, ObsLat, RTS)
'DMO, 03-07-2008 indien flags aangaf dat
If RTS.flags > 0 Then
   RTS.Rise = -1
   RTS.Setting = -1
End If
List1.AddItem "Opkomst       : " + StrHMS_DMS(RTS.Rise * 180 / PI, 3, 0, False, False, "h", 2)
List1.AddItem "Doorgang      : " + StrHMS_DMS(RTS.Transit * 180 / PI, 3, 0, False, False, "h", 2)
List1.AddItem "Ondergang     : " + StrHMS_DMS(RTS.Setting * 180 / PI, 3, 0, False, False, "h", 2)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Long
Dim stext As String
If Me.ActiveControl.Name = "List1" Then
    If KeyCode = 67 And Shift = 2 Then
       ' Me.List1.
       For I = 1 To List1.ListCount
           stext = stext & vbCrLf & List1.List(I - 1)
       Next
       Clipboard.Clear
       Clipboard.SetText (stext)
    End If
End If
End Sub

Private Sub Form_Load()
' What to do when this program starts up.

#If FRANS Then
    Me.Caption = "Astronomie"
    Me.mnuFile.Caption = "&Fichier"
    Me.mnuSettings.Caption = "&Mise"
    Me.mnuOptions.Caption = "&Options"
    Me.mnuAlmanac.Caption = "&Almanach"
    Me.mnuSterrenkaart.Caption = "&Carte du Ciel"
    Me.mnuJupiter.Caption = "&Jupiter"
    Me.mnuSaturnus.Caption = "&Saturne"
    Me.mnuTabelen.Caption = "&Tableaux"
    Me.mnuHelp.Caption = "A&ide"
    Me.mnuJupiterDiagram.Caption = "Graphique Jupiter"
    Me.mnuJupiterMeridiaan.Caption = "Méridian de Jupiter"
    Me.mnuJupiterMoons.Caption = "Satellites de Jupiter"
    Me.mnuJupiterMoonsPhenomena.Caption = "Phénomène de satellites de Jupiter"
    Me.mnuSaturnusDiagram.Caption = "Graphique Saturne"
    Me.mnuSaturnusMoons.Caption = "Satellites de Saturne"
    Me.mnuSaturnusMoonsPhenomena.Caption = "Phénomène de satellites de Saturne"
    Me.mnuAbout.Caption = "Info"
    Me.mnuInfo.Caption = "Aide"
    Me.chkGrootstePrecisie.Caption = "Grande Précision"
    Me.DateTime1Frame.Caption = "Date"
    Me.Frame2.Caption = "Temps"
    Me.SetNoonButton.Caption = "Midi"
    Me.Frame3.Caption = "Changer temps"
    Me.SetNowButton.Caption = "à présent"
    Me.ComputeButton.Caption = "Calculer"
    Me.stBar.Panels(1).Text = "Local sidérale temps"
#End If
Dim Q, D, M, Date0

'PI = 4 * Atn(1)

For D = 1 To 31
    DaySelect.AddItem Right("  " & Trim(D) & " ", 4)
Next D
    DaySelect.ListIndex = 0

    Q = "JanFebMarAprMayJunJulAugSepOctNovDec"
For M = 1 To 12
    MonthSelect.AddItem " " & Mid(Q, 3 * (M - 1) + 1, 3) & " "
Next M
    MonthSelect.ListIndex = 0

   Q = Trim(DaySelect.Text) & " " & Trim(MonthSelect.Text) & " " & Year.Text
'If BCOption.Value = True Then Q = Q & " BC" Else Q = Q & " AD"

' Set initial default startup date
  SET_INTERFACE_DATE_AND_TIME_TO_NOW
 
  startTimer
End Sub

Private Sub Form_Terminate()
' What to do upon shutting down this program.
  Unload Me
End Sub

Private Sub Hrs_LostFocus()
Hrs.Text = Format(Val(Hrs.Text), "0#")
End Sub

Private Sub Min_LostFocus()
Min.Text = Format(Val(Min.Text), "0#")
End Sub

Private Sub mnuAbout_Click()
 MsgBox _
        "Astronomisch programma ontwikkeld door Dominique Molenkamp." & vbNewLine & _
        "Version " & App.Major & "." & App.Minor & "." & App.Revision & "  Copyright © 2006", _
        vbInformation, "About " & App.Title
End Sub

Private Sub mnuAlmanac_Click()
frmAlmanac.Show vbModal, Me
End Sub

Private Sub mnuEclipse_Click()
frmEclipse.Show vbModal, Me
End Sub

Private Sub mnuEphemerides_Click()
frmEphem.Show vbModal, Me
End Sub

Private Sub mnuInfo_Click()
frmInfo.Show vbModal
End Sub

Private Sub mnuJupiterDiagram_Click()
frmJupiterDiagram.Show vbModal
End Sub

Private Sub mnuJupiterMeridiaan_Click()
frmJupiterMeridiaan.Show vbModal, Me
End Sub

Private Sub mnuJupiterMoons_Click()
frmJupiterMoons.Show vbModal
End Sub

Private Sub mnuJupiterMoonsPhenomena_Click()
frmJupiterPhenomena.Show vbModal, Me
End Sub

Private Sub mnuSaturnusDiagram_Click()
frmSaturnusDiagram.Show vbModal, Me
End Sub



Private Sub mnuSaturnusMoonsPhenomena_Click()
frmSaturnusPhenomena.Show vbModal, Me
End Sub

Private Sub mnuSaturnusMoons_Click()
frmSaturnus.Show vbModal, Me
End Sub

Private Sub mnuSettings_Click()
    frmSettings.Show
End Sub

Private Sub mnuSterrenkaart_Click()
frmSterren.Show vbModal, Me
End Sub

Private Sub mnuSurvey_Click()
frmOverzicht.Show vbModal, Me
End Sub

Private Sub mnuTabelen_Click()
frmTabellen.Show vbModal, Me
End Sub

Private Sub Sec_LostFocus()
Sec.Text = Format(Sec.Text, "#0.00")
End Sub

Private Sub Year_Change()
ADJUST_MONTH_LENGTH
End Sub

Private Sub ADOption_Click()
' Set date to AD mode.
'  ADOption.FontBold = True
'  BCOption.FontBold = False
  ADJUST_MONTH_LENGTH
End Sub

Private Sub BCOption_Click()
' Set date to BC mode.
  'BCOption.FontBold = True
  'ADOption.FontBold = False
  ADJUST_MONTH_LENGTH
End Sub

Private Sub MonthSelect_Click()
ADJUST_MONTH_LENGTH
End Sub

Private Sub Set00HrButton_Click()
' Set interface clock time setting to 00:00
  Hrs.Text = "00": Min.Text = "00": Sec.Text = "00"
End Sub

Private Sub SetNoonButton_Click()
' Set interface clock time setting to 12:00
  Hrs.Text = "12": Min.Text = "00": Sec.Text = "00"
End Sub

Private Sub ADJUST_MONTH_LENGTH()
       
End Sub

Private Sub SetNowButton_Click()
' Set interface date and time to Now according to system clock.

  SET_INTERFACE_DATE_AND_TIME_TO_NOW
  
End Sub

Private Sub SET_INTERFACE_DATE_AND_TIME_TO_NOW()
' Set interface date and time to Now according to system clock.

' DEPENDENCIES: ADJUST_MONTH_LENGTH

Dim Q, QD, QT, MMM, DD, YYYY, HH, mm, ss

' Read current time and date settings from system clock.
   Q = Now
  QT = Format(Q, "hh:mm:ss")
  QD = Format(Q, "dd mm yyyy")

' Set interface time setting to match system clock time.
  HH = Left(QT, 2)
  mm = Mid(QT, InStr(1, QT, ":") + 1, 2)
  ss = Mid(QT, InStr(InStr(1, QT, ":") + 1, QT, ":") + 1, 2)
  Hrs.Text = HH: Min.Text = mm: Sec.Text = ss

' Set interface date setting to match system clock date.
  MMM = Mid(QD, InStr(1, QD, " ") + 1, 3)
  DD = Mid(QD, 1, InStr(1, QD, " "))
  YYYY = Val(Mid(QD, InStr(InStr(1, QD, " ") + 1, QD, " ") + 1, 4))
  MonthSelect.ListIndex = MMM - 1
  ADJUST_MONTH_LENGTH
  DaySelect.ListIndex = DD - 1
  Year.Text = Trim(YYYY)
'  ADOption.Value = True
  
End Sub

Private Function INTERFACE_DATE() As String
' Return the current interface date setting as a date string
' in the standard format such as "20 MAY 1977 BC|AD"

  Dim Q, M, D, Y
  
   D = Right(" " & Trim(DaySelect.Text) & " ", 3)
   M = Trim(MonthSelect.Text) & " "
   Y = Year.Text
'If BCOption.Value = True Then y = y & " BC" Else y = y & " AD"
   Y = Right("      " & Y, 7)
  
   INTERFACE_DATE = D & M & Y
  
End Function

Private Function INTERFACE_TIME() As String
' Return the current interface time setting as a time string
' in the standard format such as "01:23:45"

  INTERFACE_TIME = Hrs.Text & ":" & Min.Text & ":" & Sec.Text
  
End Function

Private Sub startTimer()
    With tmrInterval
        .Enabled = False
        .Interval = 100 'Event every minute
        .Enabled = True
    End With
    updateTimerStatus
End Sub

Private Sub stopTimer()
    tmrInterval.Enabled = False
    stBar.Panels(1).Enabled = False
End Sub

Private Sub tmrInterval_Timer()
    updateTimerStatus
End Sub
    
Private Sub updateTimerStatus()
Dim sdat As String
Dim dat As tDatum
    Dim iNext As Integer
    sdat = Format(Now(), "dd-mm-yyyy hh:mm:ss")
    dat.jj = Val(Mid(sdat, 7, 4))
    dat.mm = Val(Mid(sdat, 4, 2))
    dat.DD = Val(Mid(sdat, 1, 2)) + Val(Mid(sdat, 12, 2)) / 24# + Val(Mid(sdat, 15, 2)) / 1440# + Val(Mid(sdat, 18, 2)) / 86400#
    stBar.Panels(2).Text = Format(PlaatselijkeSterrentijd(dat), "hh:mm:ss")
    stBar.Panels(4).Text = Format(sdat, "hh:mm:ss")
End Sub

