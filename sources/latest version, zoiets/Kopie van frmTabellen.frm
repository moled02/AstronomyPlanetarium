VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTabellen 
   Caption         =   "Tabelgenerator"
   ClientHeight    =   8535
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   16845
   Icon            =   "frmTabellen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   16845
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdEditSettings 
      Caption         =   "Edit Settings"
      Height          =   495
      Left            =   2040
      TabIndex        =   11
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periode"
      Height          =   975
      Left            =   8640
      TabIndex        =   8
      Top             =   480
      Width           =   3735
      Begin VB.TextBox txtEindPeriode 
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtBeginperiode 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.Line Line1 
         X1              =   1800
         X2              =   1920
         Y1              =   547
         Y2              =   547
      End
   End
   Begin VB.CommandButton cmdPlaneet 
      Caption         =   "per &Planet"
      Height          =   495
      Left            =   6240
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CheckBox chkGrootstePrecisie 
      Caption         =   "Greatest Precision"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   840
      Width           =   2295
   End
   Begin MSComctlLib.ProgressBar pgbBereken 
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   7500
      Width           =   4300
      _ExtentX        =   7594
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin RichTextLib.RichTextBox txtTabellen 
      Height          =   5175
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   9128
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmTabellen.frx":030A
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   8160
      Width           =   16845
      _ExtentX        =   29713
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Progress"
            TextSave        =   "Progress"
            Key             =   "progress"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Width           =   7832
            MinWidth        =   7832
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtInstellingen 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   360
      Width           =   5895
   End
   Begin VB.CommandButton cmdBereken 
      Caption         =   "&Calculate"
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblInstellingen 
      Caption         =   "File with settings:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmTabellen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objInstellingen As New clsInstellingen
Private blnDoorgaan As Boolean
Private Sub cmdBereken_Click()
Dim I As Long
Dim plnaam As Variant
Dim j As Long
Dim JD0 As Double
Dim sUitvoer As String
Dim BeginWeekNr As Long, EindWeekNr As Long
#If FRANS Then
plnaam = Array("Soleil   ", "Mercure  ", "Vénus    ", "Terre    ", "Mars     ", _
               "Jupiter  ", "Saturne  ", "Uranus   ", "Neptune  ", "Pluto    ", "Lune     ")
#Else
plnaam = Array("Zon      ", "Mercurius", "Venus    ", "Aarde    ", "Mars     ", _
               "Jupiter  ", "Saturnus ", "Uranus   ", "Neptunus ", "Pluto    ", "Maan     ")
#End If
Call objInstellingen.leesopmaakbestand(txtInstellingen.Text)
Dim ddate As tDatum
Dim ObsLon As Double, ObsLat As Double, TimeZone As Double, Height As Double
Dim JD_ZT As Double, JD_WT As Double
Dim weeknr As Long

ddate.jj = frmPlanets.Year
ddate.mm = 1
ddate.DD = 1
sUitvoer = ""
BeginWeekNr = Int(ddate.jj * 100#) + 1
Call WeekDate(BeginWeekNr, ddate)
JD0 = KalenderNaarJD(ddate)
Call Zomertijd_Wintertijd(frmPlanets.Year, JD_ZT, JD_WT)

EindWeekNr = Int((frmPlanets.Year + 1) * 100#) + 1
Call WeekDate(EindWeekNr, ddate)
jde = KalenderNaarJD(ddate)
blnDoorgaan = True
Me.txtTabellen.Text = ""
j = 0
Do While JD0 <= jde And blnDoorgaan
    DoEvents
    sUitvoer = ""
    j = j + 1
    TimeZone = TijdCorrectie(JD0 + 0.2, JD_ZT, JD_WT)
    ddate = JDNaarKalender(JD0)
    #If ENGELS Then
    sUitvoer = sUitvoer + "Sunrising and sets" + vbCrLf
    #ElseIf FRANS Then
    sUitvoer = sUitvoer + "Soleil lever et déclin" + vbCrLf
    #Else
    sUitvoer = sUitvoer + "Zonsopkomsten en ondergangen" + vbCrLf
    #End If
    Call Calculate(0, ddate, ObsLon, ObsLat, TimeZone, Height, 1, 7, sUitvoer)
    sUitvoer = sUitvoer + vbCrLf
    #If ENGELS Then
    sUitvoer = sUitvoer + "Moonrising and sets" & vbCrLf
    #ElseIf FRANS Then
    sUitvoer = sUitvoer + "Lune lever et déclin" & vbCrLf
    #Else
    sUitvoer = sUitvoer + "Maansopkomsten en ondergangen" & vbCrLf
    #End If
    Call Calculate(10, ddate, ObsLon, ObsLat, TimeZone, Height, 1, 7, sUitvoer)
    sUitvoer = sUitvoer + vbCrLf
    #If ENGELS Then
    sUitvoer = sUitvoer + "Rising and settings of the planets. (Calculated for the 1st day of the week)" + vbCrLf
    #ElseIf FRANS Then
    sUitvoer = sUitvoer + "Lever et déclin de planète. (Valable pour 1 jour de semain)" + vbCrLf
    #Else
    sUitvoer = sUitvoer + "Opkomsten en ondergangen van de planeten. (Gelden voor de 1e dag van de week)" + vbCrLf
    #End If
        
    For I = 1 To 9
        If I <> 3 Then
             sUitvoer = sUitvoer + plnaam(I) + " "
             Call Calculate(I, ddate, ObsLon, ObsLat, TimeZone, Height, 1, 1, sUitvoer)
        End If
    Next
    Me.txtTabellen.Text = Me.txtTabellen.Text + sUitvoer + vbCrLf
    DoEvents
    pgbBereken.Value = Min(j / 53# * 100#, 100#)
    'Debug.Print j
    JD0 = JD0 + 7
Loop
'Me.txtTabellen.Text = sUitvoer

pgbBereken.Value = 0
End Sub

Function Min(x, Y)
If x < Y Then
    Min = x
Else
    Min = Y
End If
End Function


Private Sub Calculate(Planet As Long, ddate As tDatum, _
                   ObsLon As Double, ObsLat As Double, TimeZone As Double, Height As Double, _
                   Interval_dagen As Double, Aant_berek As Long, ByRef sUitvoer As String)

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
Dim Elongation As Double
Dim Magnitude As Double
Dim Semidiameter As Double
Dim PolarSemiDiameter As Double
Dim NutLon As Double, NutObl As Double
Dim Parallax As Double, MoonHeight As Double
Dim JupiterPhysData As TJUPITERPHYSDATA
Dim MarsPhysData As TMARSPHYSDATA
Dim SunPhysData As TSUNPHYSDATA
Dim SaturnRingData As TSATURNRINGDATA
Dim AltSaturnRingData As TALTSATURNRINGDATA
Dim MoonPhysData As TMOONPHYSDATA
Dim C As Long
Dim JDOfCarr As Double
Dim deltaT As Double
Dim RTS As tRiseSetTran, RTS1 As tRiseSetTran, RTS2 As tRiseSetTran
Dim sLatitude As String, sLongitude As String
Dim LAST As Double
Dim RhoCosPhi As Double, RhoSinPhi As Double
Dim JD0 As Double
Dim sSterbeeld As String
Dim Az As Double, hg As Double, Alt As Double, dAlt As Double, maxhoogte As Double
Dim sScheidingsteken As String
dat.jj = ddate.jj
dat.mm = ddate.mm
dat.DD = ddate.DD
'tt = (Hrs + Min / 60 + Sec / 3600) / 24
'dat.DD = dat.DD + tt

Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Latitude", _
            sLatitude)
ObsLat = Val(sLatitude) * Pi / 180
Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Longitude", _
            sLongitude)
ObsLon = Val(sLongitude) * Pi / 180

JD0 = KalenderNaarJD(dat)
I = 0
Do While I < Aant_berek
    T = JDToT(JD0 + TimeZone) '+ i * Interval_dagen)
    deltaT = ApproxDeltaT(T)
    T0 = (floor(T * 36525 + 0.50001 - TimeZone) - 0.5 + TimeZone) / 36525#
    DtofUT = T0 + secToT * deltaT
    't = DtofUT '+ i * Interval_dagen / 36525#
    't = t + secToT * deltaT
    Call NutationConst(T, NutLon, NutObl)
    Obl = Obliquity(T)
    Call objInstellingen.OpslaanGegeven(Planet, "date", JD0, "")
    
    LAST = SiderealTime(T) + NutLon * Cos(Obl) - ObsLon
    Call ObserverCoord(ObsLat, Height, RhoCosPhi, RhoSinPhi)
    Call objInstellingen.OpslaanGegeven(Planet, "time", I * Pi / 12 * 60, "")
    
    '======================== MAAN =======================
    
    Dim Dist As Double, dkm As Double, illum As Double, diam As Double
    Dim dist1 As Double, dkm1 As Double
    Dim dist2 As Double, dkm2 As Double
    Dim l As Double, B As Double
    Dim sMoon As TSVECTOR
    Dim RAx As Double, DeclX As Double
    
    If Planet = 10 Then
        If chkGrootstePrecisie = 0 Then
            Call modMoonPos.MoonPos(T, sMoon)
            Call EclToEqu(sMoon.l, sMoon.B, Obl, RA, Decl)
        Else
            Call Lune(TToJD(T), RA, Decl, Dist, dkm, diam, phase, illum)
            Call Lune(TToJD(T - Dist * LightTimeConst), RA, Decl, Dist, dkm, diam, phase, illum)
            RA = RA * Pi / 12
            Decl = Decl * Pi / 180
            'coordinaten zijn voor J2000. Omzetten naar huidige dag, en daarna appearent berekenen
            Call PrecessFK5(0, T, RA, Decl)
        End If
        maxhoogte = 90 - RToD * (ObsLat - Decl)
        Call objInstellingen.OpslaanGegeven(Planet, "maxhoogte", maxhoogte, "")
        
        Call objInstellingen.OpslaanGegeven(10, "a_equm", RA * RToD, "")
        Call objInstellingen.OpslaanGegeven(10, "d_equm", Decl * RToD, "")

        RAx = RA: DeclX = Decl
        Call PrecessFK5(T, 0, RAx, DeclX)
        Call objInstellingen.OpslaanGegeven(10, "a_equ2000", RAx * RToD, "")
        Call objInstellingen.OpslaanGegeven(10, "d_equ2000", DeclX * RToD, "")
        
        Call SterBld(RAx, DeclX, 0#, sSterbeeld)
        Call objInstellingen.OpslaanGegeven(10, "sterrenbeeld", 0, sSterbeeld)

        Call Nutation(NutLon, NutObl, Obl, RA, Decl)
        'List1.AddItem "Maan          : " + StrHMS_DMS(RA * 180 / Pi, 7, 3, False, "h", 2) + vbTab + StrHMS_DMS(Decl * 180 / Pi, 7, 2, True, "g", 3)
        Call objInstellingen.OpslaanGegeven(10, "a_equt", RA * RToD, "")
        Call objInstellingen.OpslaanGegeven(10, "d_equt", Decl * RToD, "")
        
        Call EquToHor(RA, Decl, LAST, ObsLat, Az, Alt)
        Call objInstellingen.OpslaanGegeven(Planet, "az", Az * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "hoogte", Alt * RToD, "")
        
        If (Alt > 0) Then
            dAlt = AtmRefraction(Alt)
            Call objInstellingen.OpslaanGegeven(Planet, "hoogteatm", (Alt + dAlt) * RToD, "")
        Else
            Call objInstellingen.OpslaanGegeven(Planet, "hoogteatm", Alt * RToD, "")
        End If


        Call EquToEcl(RA, Decl, Obl, SGeo.l, SGeo.B)

        SHelio.l = 0: SHelio.B = 0: SHelio.r = 0
        Call PlanetPosHi(0, T, sAarde, chkGrootstePrecisie = False)
        Call HelioToGeo(SHelio, sAarde, SSun)
        Call PlanetPosHi(0, T - SSun.r * LightTimeConst, sAarde, chkGrootstePrecisie = 0)
        Call HelioToGeo(SHelio, sAarde, SSun)
        Call MoonPhysEphemeris(T, SGeo, SSun, Obl, NutLon, NutObl, MoonPhysData)
        
        Call objInstellingen.OpslaanGegeven(10, "pos_h", MoonPhysData.I * RToD, "")
        Call objInstellingen.OpslaanGegeven(10, "fase", MoonPhysData.k, "")
        Call objInstellingen.OpslaanGegeven(10, "posh_limb", MoonPhysData.x * RToD, "")
        Call objInstellingen.OpslaanGegeven(10, "l_olib", -MoonPhysData.ld * RToD, "")
        Call objInstellingen.OpslaanGegeven(10, "b_olib", MoonPhysData.Bd * RToD, "")
        Call objInstellingen.OpslaanGegeven(10, "l_plib", MoonPhysData.ldd * RToD, "")
        Call objInstellingen.OpslaanGegeven(10, "b_plib", MoonPhysData.bdd * RToD, "")
        Call objInstellingen.OpslaanGegeven(10, "l_tlib", MoonPhysData.l * RToD, "")
        Call objInstellingen.OpslaanGegeven(10, "b_tlib", MoonPhysData.B * RToD, "")
        Call objInstellingen.OpslaanGegeven(10, "term", MoonPhysData.T * RToD, "")

        SGeo.r = dkm
        Call objInstellingen.OpslaanGegeven(10, "afstand", SGeo.r, "")
        If chkGrootstePrecisie = 1 Then
            Parallax = asin(EarthRadius / dkm)
        Else
            Parallax = asin(EarthRadius / sMoon.r)
        End If
        MoonHeight = MoonSetHeight(Parallax)
        Call objInstellingen.OpslaanGegeven(10, "par", Parallax, "")
        If chkGrootstePrecisie = 1 Then
            Call objInstellingen.OpslaanGegeven(10, "geoc_diam", 2 * MoonSemiDiameter(SGeo.r) * SToR * RToD, "")
        Else
            Call objInstellingen.OpslaanGegeven(10, "geoc_diam", 2 * MoonSemiDiameter(sMoon.r) * SToR * RToD, "")
        End If
        
        If chkGrootstePrecisie = 0 Then
            Call modMoonPos.MoonPos(T - 1 / 36525, sMoon)
            Call EclToEqu(sMoon.l, sMoon.B, Obl, RA1, Decl1)
        Else
            Call Lune(TToJD(T - 1 / 36525), RA1, Decl1, dist1, dkm1, diam, phase, illum)
            Call Lune(TToJD(T - dist1 * LightTimeConst - 1 / 36525), RA1, Decl1, dist1, dkm1, diam, phase, illum)
            RA1 = RA1 * Pi / 12
            Decl1 = Decl1 * Pi / 180
            'coordinaten zijn voor J2000. Omzetten naar huidige dag, en daarna appearent berekenen
            Call PrecessFK5(0, T + 1 / 36525, RA1, Decl1)
        End If

        Call Nutation(NutLon, NutObl, Obl, RA1, Decl1)
        
        If chkGrootstePrecisie = False Then
            Call modMoonPos.MoonPos(T + 1 / 36525, sMoon)
            Call EclToEqu(sMoon.l, sMoon.B, Obl, RA2, Decl2)
        Else
            Call Lune(TToJD(T + 1 / 36525), RA2, Decl2, dist2, dkm2, diam, phase, illum)
            Call Lune(TToJD(T - dist2 * LightTimeConst + 1 / 36525), RA2, Decl2, Dist, dkm2, diam, phase, illum)
            RA2 = RA2 * Pi / 12
            Decl2 = Decl2 * Pi / 180
            'coordinaten zijn voor J2000. Omzetten naar huidige dag, en daarna appearent berekenen
            Call PrecessFK5(0, T + 1 / 36525, RA2, Decl2)
        End If
        Call Nutation(NutLon, NutObl, Obl, RA2, Decl2)
        
        Call RiseSet(T0, deltaT, RA1, Decl1, RA, Decl, RA2, Decl2, MoonHeight, ObsLon, ObsLat, RTS)
        'DMO, 03-07-2008 indien flags aangaf dat
        If RTS.flags > 0 Then
         RTS.Rise = -1
         RTS.Setting = -1
        End If

        If RTS.Rise >= 0 Then Call objInstellingen.OpslaanGegeven(10, "rts_r1", RTS.Rise * RToD, "") Else Call objInstellingen.OpslaanGegeven(10, "rts_r1", 0, "------")
        If RTS.Transit >= 0 Then Call objInstellingen.OpslaanGegeven(10, "rts_t1", RTS.Transit * RToD, "") Else Call objInstellingen.OpslaanGegeven(10, "rts_t1", 0, "------")
        If RTS.Setting >= 0 Then Call objInstellingen.OpslaanGegeven(10, "rts_s1", RTS.Setting * RToD, "") Else Call objInstellingen.OpslaanGegeven(10, "rts_s1", 0, "------")

    End If
    
    
    '======================== ZON ========================
    If Planet = 0 Then
        SHelio.l = 0: SHelio.B = 0: SHelio.r = 0
        Call objInstellingen.OpslaanGegeven(Planet, "l_helio", 0, "")
        Call objInstellingen.OpslaanGegeven(Planet, "b_helio", 0, "")
        Call objInstellingen.OpslaanGegeven(Planet, "r_helio", 0, "")
    
        Call PlanetPosHi(Planet, T, sAarde, chkGrootstePrecisie = 0)
        Call HelioToGeo(SHelio, sAarde, SGeo)
        Call objInstellingen.OpslaanGegeven(Planet, "l_geoc", SGeo.l * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "b_geoc", SGeo.B * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "r_geoc", SGeo.r, "")
        Call PlanetPosHi(Planet, T - SGeo.r * LightTimeConst, sAarde, chkGrootstePrecisie = 0)
        Call HelioToGeo(SHelio, sAarde, SGeo)
        Call objInstellingen.OpslaanGegeven(Planet, "l_geom", SGeo.l * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "b_geom", SGeo.B * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "r_geom", SGeo.r, "")
        
        Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
        Call objInstellingen.OpslaanGegeven(Planet, "a_equm", RA * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "d_equm", Decl * RToD, "")
        
        maxhoogte = 90 - RToD * (ObsLat - Decl)
        Call objInstellingen.OpslaanGegeven(Planet, "maxhoogte", maxhoogte, "")
        
        Call PrecessFK5(T, 0, RA, Decl)
        Call objInstellingen.OpslaanGegeven(Planet, "a_equ2000", RA * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "d_equ2000", Decl * RToD, "")
        Call SterBld(RA, Decl, 0#, sSterbeeld)
        Call objInstellingen.OpslaanGegeven(Planet, "sterrenbeeld", 0, sSterbeeld)
    
        Call SunPhysEphemeris(T, SGeo.l, Obl, NutLon, SunPhysData)
        C = CarringtonRotation(KalenderNaarJD(dat))
        JDOfCarr = JDOfCarringtonRotation(C)
        Call objInstellingen.OpslaanGegeven(Planet, "l0", SunPhysData.l0 * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "b0", SunPhysData.b0 * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "p", SunPhysData.P * RToD, "")
        
        Call ConvertVSOP_FK5(T, SGeo.l, SGeo.B)
        Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA, Decl)
        Call objInstellingen.OpslaanGegeven(Planet, "a_equt", RA * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "d_equt", Decl * RToD, "")
       
        Call ParallaxHi(SolarParallax / SGeo.r, LAST, RhoCosPhi, RhoSinPhi, RA, Decl)
        Call objInstellingen.OpslaanGegeven(Planet, "a_top", RA * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "d_top", Decl * RToD, "")
        
        Call EquToHor(RA, Decl, LAST, ObsLat, Az, Alt)
        Call objInstellingen.OpslaanGegeven(Planet, "az", Az * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "hoogte", Alt * RToD, "")
        
        If (Alt > 0) Then
            dAlt = AtmRefraction(Alt)
            Call objInstellingen.OpslaanGegeven(Planet, "hoogteatm", (Alt + dAlt) * RToD, "")
        Else
            Call objInstellingen.OpslaanGegeven(Planet, "hoogteatm", Alt * RToD, "")
        End If


        ' bepalen opkomst e.d.
        Call PlanetPosHi(0, T0 - 1 / 36525, sAarde, chkGrootstePrecisie = 0)
        Call HelioToGeo(SHelio, sAarde, SGeo)
        Call EclToEqu(SGeo.l, SGeo.B, Obl, RA1, Decl1)
        
        Call PlanetPosHi(0, T0, sAarde, chkGrootstePrecisie = False)
        Call HelioToGeo(SHelio, sAarde, SGeo)
        Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
        
        Call PlanetPosHi(0, T0 + 1 / 36525, sAarde, chkGrootstePrecisie = 0)
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
        If RTS.Rise >= 0 Then Call objInstellingen.OpslaanGegeven(0, "rts_r1", RTS.Rise * RToD, "") Else Call objInstellingen.OpslaanGegeven(0, "rts_r1", 0, "------")
        If RTS.Transit >= 0 Then Call objInstellingen.OpslaanGegeven(0, "rts_t1", RTS.Transit * RToD, "") Else Call objInstellingen.OpslaanGegeven(0, "rts_t1", 0, "------")
        If RTS.Setting >= 0 Then Call objInstellingen.OpslaanGegeven(0, "rts_s1", RTS.Setting * RToD, "") Else Call objInstellingen.OpslaanGegeven(0, "rts_s1", 0, "------")
        If RTS1.Rise >= 0 Then Call objInstellingen.OpslaanGegeven(0, "rts_r2", RTS1.Rise * RToD, "") Else Call objInstellingen.OpslaanGegeven(0, "rts_r2", 0, "------")
        If RTS1.Transit >= 0 Then Call objInstellingen.OpslaanGegeven(0, "rts_t2", RTS1.Transit * RToD, "") Else Call objInstellingen.OpslaanGegeven(0, "rts_t2", 0, "------")
        If RTS1.Setting >= 0 Then Call objInstellingen.OpslaanGegeven(0, "rts_s2", RTS1.Setting * RToD, "") Else Call objInstellingen.OpslaanGegeven(0, "rts_s2", 0, "------")
        If RTS2.Rise >= 0 Then Call objInstellingen.OpslaanGegeven(0, "rts_r3", RTS2.Rise * RToD, "") Else Call objInstellingen.OpslaanGegeven(0, "rts_r3", 0, "------")
        If RTS2.Transit >= 0 Then Call objInstellingen.OpslaanGegeven(0, "rts_t3", RTS2.Transit * RToD, "") Else Call objInstellingen.OpslaanGegeven(0, "rts_t3", 0, "------")
        If RTS2.Setting >= 0 Then Call objInstellingen.OpslaanGegeven(0, "rts_s3", RTS2.Setting * RToD, "") Else Call objInstellingen.OpslaanGegeven(0, "rts_s3", 0, "------")
    
    ElseIf Planet > 0 And Planet < 9 Then
        Call PlanetPosHi(0, T, sAarde, chkGrootstePrecisie = 0)
        Call PlanetPosHi(Planet, T, SHelio, chkGrootstePrecisie = 0)
        Call objInstellingen.OpslaanGegeven(Planet, "l_helio", (SHelio.l + NutLon) * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "b_helio", SHelio.B * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "r_helio", SHelio.r, "")
    
        Call HelioToGeo(SHelio, sAarde, SGeo)
        Call objInstellingen.OpslaanGegeven(Planet, "l_geoc", SGeo.l * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "b_geoc", SGeo.B * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "r_geoc", SGeo.r, "")
    
        Call PlanetPosHi(Planet, T - SGeo.r * LightTimeConst, SHelio, chkGrootstePrecisie = 0)
        Call HelioToGeo(SHelio, sAarde, SGeo)
        Call objInstellingen.OpslaanGegeven(Planet, "l_geom", SGeo.l * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "b_geom", SGeo.B * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "r_geom", SGeo.r, "")
    
        Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
        
        maxhoogte = 90 - RToD * (ObsLat - Decl)
        Call objInstellingen.OpslaanGegeven(Planet, "maxhoogte", maxhoogte, "")
        
        Call objInstellingen.OpslaanGegeven(Planet, "a_equm", RA * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "d_equm", Decl * RToD, "")
    
        Call PrecessFK5(T, 0, RA, Decl)
        Call objInstellingen.OpslaanGegeven(Planet, "a_equ2000", RA * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "d_equ2000", Decl * RToD, "")
        Call SterBld(RA, Decl, 0#, sSterbeeld)
        Call objInstellingen.OpslaanGegeven(Planet, "sterrenbeeld", 0, sSterbeeld)
        
        Call ConvertVSOP_FK5(T, SGeo.l, SGeo.B)
        Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA, Decl)
        Call Aberration(T, Obl, FK5System, RA, Decl)
        Call objInstellingen.OpslaanGegeven(Planet, "a_equt", RA * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "d_equt", Decl * RToD, "")
       
        Call ParallaxHi(SolarParallax / SGeo.r, LAST, RhoCosPhi, RhoSinPhi, RA, Decl)
        Call objInstellingen.OpslaanGegeven(Planet, "a_top", RA * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "d_top", Decl * RToD, "")
        
        Call EquToHor(RA, Decl, LAST, ObsLat, Az, Alt)
        Call objInstellingen.OpslaanGegeven(Planet, "az", Az * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "hoogte", Alt * RToD, "")
        
        If (Alt > 0) Then
            dAlt = AtmRefraction(Alt)
            Call objInstellingen.OpslaanGegeven(Planet, "hoogteatm", (Alt + dAlt) * RToD, "")
        Else
            Call objInstellingen.OpslaanGegeven(Planet, "hoogteatm", Alt * RToD, "")
        End If

        phase = CalcPhase(SHelio.r, sAarde.r, SGeo.r)
        PhaseAngle = acos(2 * phase - 1)
        Elongation = CalcElongation(SHelio.r, sAarde.r, SGeo.r)
        If modpi(SHelio.l - SGeo.l) > 0 Then Elongation = -Elongation
        Magnitude = PlanetMagnitude(Planet, SHelio.r, SGeo.r, PhaseAngle, SaturnRingData.DeltaU, SaturnRingData.B)
        Semidiameter = PlanetSemiDiameter(Planet, SGeo.r, PolarSemiDiameter)
        Call objInstellingen.OpslaanGegeven(Planet, "semidiam", 2 * Semidiameter * RToD * SToR, "")
        Call objInstellingen.OpslaanGegeven(Planet, "fase", phase, "")
        Call objInstellingen.OpslaanGegeven(Planet, "fasehoek", PhaseAngle * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "magnitude", Magnitude, "")
        Call objInstellingen.OpslaanGegeven(Planet, "elongatie", Elongation * RToD, "")
    
        Select Case Planet
            Case Is = 4
                    Call MarsPhysEphemeris(T, SHelio, sAarde, SGeo, _
                                                Obl, NutLon, NutObl, _
                                                MarsPhysData)
                    Call objInstellingen.OpslaanGegeven(Planet, "de", MarsPhysData.DE * RToD, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "ds", MarsPhysData.DS * RToD, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "p", MarsPhysData.P * RToD, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "om", MarsPhysData.Om * RToD, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "qq", MarsPhysData.qq, "")
            Case Is = 5
                    Call objInstellingen.OpslaanGegeven(Planet, "semiequ", 2 * Semidiameter * SToR * RToD, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "semipol", 2 * PolarSemiDiameter * SToR * RToD, "")
                    
                    Call JupiterPhysEphemeris(T + deltaT / 36525 / 86400, SHelio, sAarde, SGeo, _
                                                   Obl, NutLon, NutObl, _
                                                  JupiterPhysData)
                    Call objInstellingen.OpslaanGegeven(Planet, "de", JupiterPhysData.DE * RToD, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "ds", JupiterPhysData.DS * RToD, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "p", JupiterPhysData.P * RToD, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "om1", JupiterPhysData.Om1 * RToD, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "om2", JupiterPhysData.Om2 * RToD, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "c", JupiterPhysData.C * RToD, "")
            Case Is = 6
                    Call CorrectSaturnSemiDiameter(SaturnRingData.B, PolarSemiDiameter)
                    Call objInstellingen.OpslaanGegeven(Planet, "semiequ", 2 * Semidiameter * SToR * RToD, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "semipol", 2 * PolarSemiDiameter * SToR * RToD, "")
    
                    Call SaturnRing(T, SHelio, SGeo, Obl, NutLon, NutObl, SaturnRingData)
                    Call AltSaturnRing(T, SHelio, SGeo, Obl, NutLon, NutObl, AltSaturnRingData)
                    Call objInstellingen.OpslaanGegeven(Planet, "ring_p", AltSaturnRingData.P * RToD, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "ring_p1", AltSaturnRingData.P1 * RToD, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "ring_u", AltSaturnRingData.u * RToD, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "ring_u1", AltSaturnRingData.U1 * RToD, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "ring_b", AltSaturnRingData.B * RToD, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "ring_semiasbinnen", SaturnRingData.aAxis / 3600, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "ring_ioasbinnen", SaturnRingData.ioaAxis / 3600, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "ring_oiasbinnen", SaturnRingData.oiaAxis / 3600, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "ring_iiasbinnen", SaturnRingData.iiaAxis / 3600, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "ring_idasbinnen", SaturnRingData.idaAxis / 3600, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "ring_semiasbuiten", SaturnRingData.bAxis / 3600, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "ring_ioasbuiten", SaturnRingData.iobAxis / 3600, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "ring_oiasbuiten", SaturnRingData.oibAxis / 3600, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "ring_iiasbuiten", SaturnRingData.iibAxis / 3600, "")
                    Call objInstellingen.OpslaanGegeven(Planet, "ring_idasbuiten", SaturnRingData.idbAxis / 3600, "")
                    SaturnB = SaturnRingData.B
                    SaturnDeltaU = SaturnRingData.DeltaU
        End Select
       
        ' bepalen opkomst e.d.
        Call PlanetPosHi(0, T0 - 1 / 36525, sAarde, chkGrootstePrecisie = 0)
        Call PlanetPosHi(Planet, T0 - 1 / 36525, SHelio, chkGrootstePrecisie = 0)
        Call HelioToGeo(SHelio, sAarde, SGeo)
        Call PlanetPosHi(Planet, T0 - 1 / 36525 - SGeo.r * LightTimeConst, SHelio, chkGrootstePrecisie = 0)
        Call HelioToGeo(SHelio, sAarde, SGeo)
        Call EclToEqu(SGeo.l, SGeo.B, Obl, RA1, Decl1)
        
        Call PlanetPosHi(0, T0, sAarde, chkGrootstePrecisie = 0)
        Call PlanetPosHi(Planet, T0, SHelio, chkGrootstePrecisie = 0)
        Call HelioToGeo(SHelio, sAarde, SGeo)
        Call PlanetPosHi(Planet, T0 - SGeo.r * LightTimeConst, SHelio, chkGrootstePrecisie = 0)
        Call HelioToGeo(SHelio, sAarde, SGeo)
        Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
        
        Call PlanetPosHi(0, T0 + 1 / 36525, sAarde, chkGrootstePrecisie = 0)
        Call PlanetPosHi(Planet, T0 + 1 / 36525, SHelio, chkGrootstePrecisie = 0)
        Call HelioToGeo(SHelio, sAarde, SGeo)
        Call PlanetPosHi(Planet, T0 + 1 / 36525 - SGeo.r * LightTimeConst, SHelio, chkGrootstePrecisie = 0)
        Call HelioToGeo(SHelio, sAarde, SGeo)
        Call EclToEqu(SGeo.l, SGeo.B, Obl, RA2, Decl2)
        
        Call RiseSet(T0, deltaT, RA1, Decl1, RA, Decl, RA2, Decl2, h0Planet, ObsLon, ObsLat, RTS)
        'DMO, 03-07-2008 indien flags aangaf dat
        If RTS.flags > 0 Then
         RTS.Rise = -1
         RTS.Setting = -1
        End If
        If RTS.Rise >= 0 Then Call objInstellingen.OpslaanGegeven(Planet, "rts_r1", RTS.Rise * RToD, "") Else Call objInstellingen.OpslaanGegeven(10, "rts_r1", 0, "------")
        If RTS.Transit >= 0 Then Call objInstellingen.OpslaanGegeven(Planet, "rts_t1", RTS.Transit * RToD, "") Else Call objInstellingen.OpslaanGegeven(10, "rts_t1", 0, "------")
        If RTS.Setting >= 0 Then Call objInstellingen.OpslaanGegeven(Planet, "rts_s1", RTS.Setting * RToD, "") Else Call objInstellingen.OpslaanGegeven(10, "rts_s1", 0, "------")
    ElseIf Planet = 9 Then ', dus Pluto
        'Dit is een speciaal geval. Coordinaten zijn voor 2000. Deze moeten voor de Zon worden berekend.
        'Dat is: bereken positie voor vandaag. De coordinaten omzetten naar J2000
        'dit is niet voldoende. Er moet nog een correctie plaatsvinden van Ecl. VSOP -> equ FK5-2000
        'vervolgens de positie van pluto berekenen
        Dim TAarde As TVECTOR
        Dim sZon As TSVECTOR
        Dim TPluto As TVECTOR
        SHelio.l = 0: SHelio.B = 0: SHelio.r = 0
        Call PlanetPosHi(0, T, sAarde, chkGrootstePrecisie = 0)
        Call HelioToGeo(SHelio, sAarde, SGeo)
        Call SphToRect(SGeo, TAarde)
        Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
        ' Call Reduction2000(0, RA, Decl)
        'coordinaten omzetten naar J2000
        Call PrecessFK5(T, 0, RA, Decl)
        Call EquToEcl(RA, Decl, Obliquity(0), SGeo.l, SGeo.B)
        Call SphToRect(SGeo, TAarde)
        Call EclVSOP2000_equFK52000(TAarde.x, TAarde.Y, TAarde.Z)
        Call RectToSph(TAarde, sZon)
        sAarde = SGeo
        
        Call PlanetPosHi(0, T, sAarde, chkGrootstePrecisie = 0)
        Call PlutoPos(T, SHelio)
        Call EclToRect(SHelio, Obliquity(0), TPluto)
        Dist = Sqr((TAarde.x + TPluto.x) * (TAarde.x + TPluto.x) + (TAarde.Y + TPluto.Y) * (TAarde.Y + TPluto.Y) + (TAarde.Z + TPluto.Z) * (TAarde.Z + TPluto.Z))
        Call PlutoPos(T - Dist * LightTimeConst, SHelio)
        Call EclToRect(SHelio, Obliquity(0), TPluto)
        Dist = Sqr((TAarde.x + TPluto.x) * (TAarde.x + TPluto.x) + (TAarde.Y + TPluto.Y) * (TAarde.Y + TPluto.Y) + (TAarde.Z + TPluto.Z) * (TAarde.Z + TPluto.Z))
        Call objInstellingen.OpslaanGegeven(Planet, "r_geom", Dist, "")
        RA = atan2(TPluto.Y + TAarde.Y, TPluto.x + TAarde.x)
        If RA < 0 Then
            RA = RA + Pi2
        End If
        Decl = asin((TPluto.Z + TAarde.Z) / Dist)
                Call SterBld(RA, Decl, 0#, sSterbeeld)
        Call objInstellingen.OpslaanGegeven(Planet, "sterrenbeeld", 0, sSterbeeld)
        
        maxhoogte = 90 - RToD * (ObsLat - Decl)
        Call objInstellingen.OpslaanGegeven(Planet, "maxhoogte", maxhoogte, "")
        
        Call objInstellingen.OpslaanGegeven(Planet, "a_equ2000", RA * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "d_equ2000", Decl * RToD, "")
        Call RiseSet(T0, deltaT, RA, Decl, RA, Decl, RA, Decl, h0Planet, ObsLon, ObsLat, RTS)
        'DMO, 03-07-2008 indien flags aangaf dat
        If RTS.flags > 0 Then
         RTS.Rise = -1
         RTS.Setting = -1
        End If
        
        Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA, Decl)
        Call Aberration(T, Obl, FK5System, RA, Decl)
        Call objInstellingen.OpslaanGegeven(Planet, "a_equt", RA * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "d_equt", Decl * RToD, "")
       
        Call ParallaxHi(SolarParallax / SGeo.r, LAST, RhoCosPhi, RhoSinPhi, RA, Decl)
        Call objInstellingen.OpslaanGegeven(Planet, "a_top", RA * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "d_top", Decl * RToD, "")
        
        Call EquToHor(RA, Decl, LAST, ObsLat, Az, Alt)
        Call objInstellingen.OpslaanGegeven(Planet, "az", Az * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "hoogte", Alt * RToD, "")
        
        If (Alt > 0) Then
            dAlt = AtmRefraction(Alt)
            Call objInstellingen.OpslaanGegeven(Planet, "hoogteatm", (Alt + dAlt) * RToD, "")
        Else
            Call objInstellingen.OpslaanGegeven(Planet, "hoogteatm", Alt * RToD, "")
        End If

        If RTS.Rise >= 0 Then Call objInstellingen.OpslaanGegeven(Planet, "rts_r1", RTS.Rise * RToD, "") Else Call objInstellingen.OpslaanGegeven(10, "rts_r1", 0, "------")
        If RTS.Transit >= 0 Then Call objInstellingen.OpslaanGegeven(Planet, "rts_t1", RTS.Transit * RToD, "") Else Call objInstellingen.OpslaanGegeven(10, "rts_t1", 0, "------")
        If RTS.Setting >= 0 Then Call objInstellingen.OpslaanGegeven(Planet, "rts_s1", RTS.Setting * RToD, "") Else Call objInstellingen.OpslaanGegeven(10, "rts_s1", 0, "------")
    
        phase = CalcPhase(SHelio.r, sAarde.r, Dist)
        PhaseAngle = acos(2 * phase - 1)
        Magnitude = PlanetMagnitude(9, SHelio.r, Dist, PhaseAngle, SaturnRingData.DeltaU, SaturnRingData.B)
        Elongation = CalcElongation(SHelio.r, sAarde.r, Dist)
        Semidiameter = PlanetSemiDiameter(9, Dist, PolarSemiDiameter)
        Call objInstellingen.OpslaanGegeven(Planet, "semidiam", 2 * Semidiameter * SToR * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "fase", phase, "")
        Call objInstellingen.OpslaanGegeven(Planet, "fasehoek", PhaseAngle * RToD, "")
        Call objInstellingen.OpslaanGegeven(Planet, "magnitude", Magnitude, "")
        Call objInstellingen.OpslaanGegeven(Planet, "elongatie", Elongation * RToD, "")
    End If
    
    Dim j As Double
    Dim cx As String
    Dim cy As Long
    j = 1
    Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Scheidingsteken", _
            sScheidingsteken)
    If sScheidingsteken = "" Then sScheidingsteken = "|"
            
    Do While objInstellingen.leesvolgorde(Planet, j) > 0
        cx = ""
        cy = objInstellingen.leesvolgorde(Planet, j)
        sUitvoer = sUitvoer & objInstellingen.GeefGegevenMetOpmaak(Planet, cx, cy) & sScheidingsteken
        j = j + 1
    Loop
    sUitvoer = sUitvoer & vbCrLf
    
    I = I + 1
    JD0 = JD0 + Interval_dagen
    DoEvents
Loop
End Sub


Private Sub command1_click()

End Sub

Private Sub cmdEditSettings_Click()
If Dir(txtInstellingen.Text) <> "" Then
    Call Shell("NOTEPAD " & "" & txtInstellingen.Text & "", vbNormalFocus)
End If
End Sub

Private Sub cmdPlaneet_Click()
Dim objspecialfolder As New clsSpecialFolder
Dim I As Long
Dim k As Long
Dim plnaam As Variant
Dim j As Long
Dim JD0 As Double
Dim JD As Double, JD00 As Double
Dim sUitvoer As String
Dim BeginWeekNr As Long, EindWeekNr As Long
Dim sTempFile As String
Dim nfile As Long
Dim dStapBerekening As Double
Dim sStapBerekening As String
Dim lStap As Long


#If ENGELS Then
plnaam = Array("Sun      ", "Mercury  ", "Venus    ", "Earth    ", "Mars     ", _
               "Jupiter  ", "Saturn   ", "Uranus   ", "Neptune  ", "Pluto    ", "Moon     ")
    sStapBerekening = InputBox("per hour or per day (dd.hhmm)", "Calculation")
#ElseIf FRANS Then
plnaam = Array("Soleil   ", "Mercure  ", "Vénus    ", "Terre    ", "Mars     ", _
               "Jupiter  ", "Saturne  ", "Uranus   ", "Neptune  ", "Pluto    ", "Lune     ")
    sStapBerekening = InputBox("par heure or par jour (jj.hhmm)", "Calculation")

#Else
plnaam = Array("Zon      ", "Mercurius", "Venus    ", "Aarde    ", "Mars     ", _
               "Jupiter  ", "Saturnus ", "Uranus   ", "Neptunus ", "Pluto    ", "Maan     ")
    sStapBerekening = InputBox("per uur of per dag (dd.hhmm)", "Berekening")
    
#End If
dStapBerekening = Int(Val(sStapBerekening))
sStapBerekening = Format(Frac(Val(sStapBerekening)), "00.0000")
dStapBerekening = dStapBerekening + Val(Mid(sStapBerekening, 4, 2)) / 24 + Val(Right(sStapBerekening, 2)) / 1440
If dStapBerekening = 0 Then dStapBerekening = 1

Call objInstellingen.leesopmaakbestand(txtInstellingen.Text)
Dim ddate As tDatum
Dim ObsLon As Double, ObsLat As Double, TimeZone As Double, Height As Double
Dim JD_ZT As Double, JD_WT As Double
Dim weeknr As Long
Dim dNu As tDatum
Dim dVorig As tDatum

nfile = FreeFile
sTempFile = objspecialfolder.TemporaryFolder + "\Planeten_calc"
Open sTempFile For Output As #nfile

ddate.jj = Val(Mid(txtBeginperiode, 7, 4))
ddate.mm = Val(Mid(txtBeginperiode, 4, 2))
ddate.DD = Val(Left(txtBeginperiode, 2))
JD0 = KalenderNaarJD(ddate)
'Call Zomertijd_Wintertijd(ddate.jj, JD_ZT, JD_WT)
ddate.DD = -1
dVorig = ddate

ddate.jj = Val(Mid(txtEindPeriode, 7, 4))
ddate.mm = Val(Mid(txtEindPeriode, 4, 2))
ddate.DD = Val(Left(txtEindPeriode, 2))
jde = KalenderNaarJD(ddate)

blnDoorgaan = True
Me.txtTabellen.Text = ""
For I = 0 To 10
    If I <> 3 Then
    JD = JD0
    Print #nfile, "---------------------------------------------------------------"
    #If ENGELS Then
    Print #nfile, "Survey of : " + plnaam(I)
    #ElseIf FRANS Then
    Print #nfile, "Aperçu de : " + plnaam(I)
    #Else
    Print #nfile, "Overzicht van : " + plnaam(I)
    #End If
    Print #nfile, "---------------------------------------------------------------"
    Me.txtTabellen.Text = Me.txtTabellen.Text + plnaam(I)
    j = 0
    Do While JD <= jde And blnDoorgaan
        DoEvents
        j = j + 1
        Call Zomertijd_Wintertijd(JDNaarKalender(JD).jj, JD_ZT, JD_WT)
        TimeZone = TijdCorrectie(JD + 0.001, JD_ZT, JD_WT)
        ddate = JDNaarKalender(JD)
        dNu = ddate
        dNu.DD = Int(dNu.DD)
        If dNu.DD = dVorig.DD And dNu.mm = dVorig.mm And dNu.jj = dVorig.jj Then
             Call objInstellingen.InstellingAanUit(I, "rts_r1", False)
             Call objInstellingen.InstellingAanUit(I, "rts_t1", False)
             Call objInstellingen.InstellingAanUit(I, "rts_s1", False)
             Call objInstellingen.InstellingAanUit(I, "rts_r2", False)
             Call objInstellingen.InstellingAanUit(I, "rts_t2", False)
             Call objInstellingen.InstellingAanUit(I, "rts_s2", False)
             Call objInstellingen.InstellingAanUit(I, "rts_r3", False)
             Call objInstellingen.InstellingAanUit(I, "rts_t3", False)
             Call objInstellingen.InstellingAanUit(I, "rts_s3", False)
        Else
             Call objInstellingen.InstellingAanUit(I, "rts_r1", True)
             Call objInstellingen.InstellingAanUit(I, "rts_t1", True)
             Call objInstellingen.InstellingAanUit(I, "rts_s1", True)
             Call objInstellingen.InstellingAanUit(I, "rts_r2", True)
             Call objInstellingen.InstellingAanUit(I, "rts_t2", True)
             Call objInstellingen.InstellingAanUit(I, "rts_s2", True)
             Call objInstellingen.InstellingAanUit(I, "rts_r3", True)
             Call objInstellingen.InstellingAanUit(I, "rts_t3", True)
             Call objInstellingen.InstellingAanUit(I, "rts_s3", True)
        End If
             dVorig = dNu
             sUitvoer = ""
             Call Calculate(I, ddate, ObsLon, ObsLat, TimeZone, Height, 1#, 1#, sUitvoer)
             Print #nfile, sUitvoer;
             sUitvoer = ""
             JD = JD0 + j * dStapBerekening
             ddate = JDNaarKalender(JD)

         'Me.txtTabellen.Text = Me.txtTabellen.Text + "."
         pgbBereken.Value = Min(j / ((jde - JD0) / dStapBerekening) * 100#, 100#)
         'Me.txtTabellen.Text = Me.txtTabellen.Text + sUitvoer
    Loop
    Me.txtTabellen.Text = Me.txtTabellen.Text & "... ok" & vbCrLf
    End If
Next
        
Close nfile
Open sTempFile For Binary As #nfile
sUitvoer = Space(LOF(nfile))
Get nfile, , sUitvoer
Close nfile
Me.txtTabellen.Text = sUitvoer
'Me.txtTabellen.Text = sUitvoer
    
pgbBereken.Value = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim stext As String
    If KeyCode = 67 And Shift = 2 Then
       stext = Me.txtTabellen.Text
       Clipboard.Clear
       Clipboard.SetText (stext)
    End If
End Sub

Private Sub Form_Load()
Dim ddate As tDatum
Dim BeginWeekNr As Long
Dim EindWeekNr As Long
Dim JD_ZT As Double, JD_WT As Double

#If FRANS Then
    Me.Caption = "Faire de tableaux"
    Me.lblInstellingen.Caption = "Fichier avec mise"
    Me.chkGrootstePrecisie.Caption = "&Grande Précision"
    Me.cmdBereken.Caption = "&Calculer"
    Me.cmdPlaneet.Caption = "par &Planète"
    Me.cmdEditSettings.Caption = "C&hanger fichier"
#End If
txtInstellingen.Text = App.Path & "\instel"

ddate.jj = frmPlanets.Year
ddate.mm = 1
ddate.DD = 1
BeginWeekNr = Int(ddate.jj * 100#) + 1
Call WeekDate(BeginWeekNr, ddate)
txtBeginperiode = ddate.DD & "-" & ddate.mm & "-" & ddate.jj
JD0 = KalenderNaarJD(ddate)
Call Zomertijd_Wintertijd(ddate.jj, JD_ZT, JD_WT)

EindWeekNr = Int((frmPlanets.Year + 1) * 100#) + 1
Call WeekDate(EindWeekNr, ddate)
Me.txtEindPeriode = ddate.DD & "-" & ddate.mm & "-" & ddate.jj

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
blnDoorgaan = False
End Sub

Private Sub Form_Resize()
On Error Resume Next
txtTabellen.Width = Me.Width - 245
txtTabellen.Height = Me.Height - 3030
Me.pgbBereken.top = StatusBar1.top + 75
End Sub

