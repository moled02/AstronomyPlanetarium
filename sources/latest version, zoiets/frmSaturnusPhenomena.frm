VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSaturnusPhenomena 
   Caption         =   "Elongation Saturnmoons"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9750
   Icon            =   "frmSaturnusPhenomena.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   9750
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox txtVerschijnselen 
      Height          =   5535
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9763
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmSaturnusPhenomena.frx":030A
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
   Begin MSComctlLib.ProgressBar PBVoortgang 
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   5880
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmSaturnusPhenomena"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ObsLon As Double, ObsLat As Double
Private blnDoorgaan As Boolean
Private Sub Form_Activate()
Dim sLatitude As String
Dim sLongitude As String
Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Latitude", _
            sLatitude)
ObsLat = Val(sLatitude) * Pi / 180
Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Longitude", _
            sLongitude)
ObsLon = Val(sLongitude) * Pi / 180

Elongaties_berekenen
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim stext As String
    If KeyCode = 17 And Shift = 0 Then
       stext = Me.txtVerschijnselen.Text
       Clipboard.Clear
       Clipboard.SetText (stext)
    End If
End Sub

Private Sub Form_Load()
#If FRANS Then
    Me.Caption = "Elongation de satellites de Saturne"
#End If
Call modSaturnMoon.SaturnMoonInit
End Sub


Private Sub ov(ByVal Planet As Long, ByRef ddate As tDatum, _
             ByVal ObsLon As Double, ByVal ObsLat As Double, ByVal TimeZone As Double, ByVal Height As Double, ByVal StartHeight As Double, _
             ByRef Opk As Double, ByRef Ond As Double)

Dim sAarde As TSVECTOR, SHelio As TSVECTOR, SGeo As TSVECTOR
Dim Obl As Double
Dim T0 As Double, RA As Double, Decl As Double, RA1 As Double, Decl1 As Double, RA2 As Double, Decl2 As Double
Dim deltaT As Double
Dim RTS As tRiseSetTran
Dim sLatitude As String, sLongitude As String

    
T0 = JDToT(KalenderNaarJD(ddate))
Obl = Obliquity(T0)
If Planet = 0 Then 'Zon
    Call PlanetPosHi(0, T0 - 1 / 36525, sAarde, True)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call EclToEqu(SGeo.l, SGeo.B, Obl, RA1, Decl1)
    
    Call PlanetPosHi(0, T0, sAarde, True)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
    
    Call PlanetPosHi(0, T0 + 1 / 36525, sAarde, True)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call EclToEqu(SGeo.l, SGeo.B, Obl, RA2, Decl2)
    
    Call RiseSet(T0, deltaT, RA1, Decl1, RA, Decl, RA2, Decl2, Height, ObsLon, ObsLat, RTS)
Else 'Saturnus
    Call PlanetPosHi(0, T0 - 1 / 36525, sAarde, True)
    Call PlanetPosHi(6, T0 - 1 / 36525, SHelio, True)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call PlanetPosHi(6, T0 - 1 / 36525 - SGeo.r * LightTimeConst, SHelio, True)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call EclToEqu(SGeo.l, SGeo.B, Obl, RA1, Decl1)
    
    Call PlanetPosHi(0, T0, sAarde, True)
    Call PlanetPosHi(6, T0, SHelio, True)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call PlanetPosHi(5, T0 - SGeo.r * LightTimeConst, SHelio, True)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
    
    Call PlanetPosHi(0, T0 + 1 / 36525, sAarde, True)
    Call PlanetPosHi(6, T0 + 1 / 36525, SHelio, True)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call PlanetPosHi(6, T0 + 1 / 36525 - SGeo.r * LightTimeConst, SHelio, True)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call EclToEqu(SGeo.l, SGeo.B, Obl, RA2, Decl2)
    
    Call RiseSet(T0, deltaT, RA1, Decl1, RA, Decl, RA2, Decl2, Height, ObsLon, ObsLat, RTS)
    
End If
Opk = RTS.Rise * 12 / Pi
Ond = RTS.Setting * 12 / Pi
End Sub


Function Zichtbaar(JD As Double) As Boolean
Dim ddate As tDatum
Dim Opk As Double, Ond As Double
Dim tijd As Double
Dim bTmp As Boolean
Dim hoogte As Double

    tijd = Frac(JD + 0.5) * 24
    ddate = JDNaarKalender(JD)
    ddate.DD = Int(ddate.DD)
    hoogte = 0
    
    Call ov(0, ddate, ObsLon, ObsLat, 0, hoogte, -6 * DToR, Opk, Ond)
    If Opk <= 0 Then
        Opk = 0
    End If
    If Ond <= 0 Then
        Ond = 24
    End If
    bTmp = (tijd < Opk) Or (tijd > Ond)
    Call ov(6, ddate, ObsLon, ObsLat, 0, hoogte, 5 * DToR, Opk, Ond)
    If Ond > Opk Then
      bTmp = (bTmp) And (tijd >= Opk) And (tijd <= Ond)
    Else
      bTmp = (bTmp) And ((tijd <= Ond) Or (tijd >= Opk))
    End If
    Zichtbaar = bTmp
End Function

Private Sub Bepaal_Elongatie(ByVal JD0 As Double, ByRef maangeg As tMaanGeg, ByVal maannr As Long, _
                           ByVal hoek As Double, ByRef JD As Double)

Dim u_U As Double, Time_interval As Double, ElongatieHoek As Double
Dim I As Long
Dim ddate As tDatum
Dim doorgaan As Boolean
Dim SaturnB As Double
Dim Dist As Double
Dim SHelio As TSVECTOR, SGeo As TSVECTOR, Obl As Double, NutLon As Double, NutObl As Double
Dim sAarde As TSVECTOR
Dim AltSaturnRingData As TALTSATURNRINGDATA
Dim T0 As Double
Dim sLatitude As String, sLongitude As String

  JD = JD0
  Obl = Obliquity(T0)
  Call NutationConst(T0, NutLon, NutObl)
  doorgaan = True
  Do While doorgaan
    T0 = JDToT(JD)
    Call PlanetPosHi(0, T0, sAarde, True)
    Call PlanetPosHi(6, T0, SHelio, True)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call PlanetPosHi(6, T0 - SGeo.r * LightTimeConst, SHelio, True)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    
    Call modSaturnMoon.BasisGegevens(JD, SaturnB, Dist)
     'Call modSaturnMoon.BasisGegevens(JD, SaturnB, Dist)
     Call modSatRing.AltSaturnRing(JDToT(JD), SHelio, SGeo, Obl, NutLon, NutObl, AltSaturnRingData)
     Call modSaturnMoon.sat_manen(JD, maannr, maangeg)

     If maannr < 6 Then
       u_U = maangeg.Manen(maannr).u - AltSaturnRingData.u * 180 / Pi
     Else
       u_U = maangeg.Titan.u - AltSaturnRingData.u * 180 / Pi
     End If
     ElongatieHoek = hoek - u_U

     ElongatieHoek = ElongatieHoek - 360 * Int(ElongatieHoek / 360)
     If ElongatieHoek < -180 Then ElongatieHoek = ElongatieHoek + 360
     If ElongatieHoek > 180 Then ElongatieHoek = ElongatieHoek - 360
'{ time interval in uren }
     If maannr < 6 Then Time_interval = ElongatieHoek / maangeg.Manen(maannr).n * 24 _
     Else Time_interval = ElongatieHoek / maangeg.Titan.n * 24

     Time_interval = Time_interval + SGeo.r * LightTimeConst * 876600
                                                        '{ * 36525 * 24 }
     JD = JD + Time_interval / 24
     doorgaan = Abs(Time_interval) > 0.05
  Loop
End Sub

Sub Elongaties_berekenen()
Dim maannr As Long
Dim JD_ZT As Double
Dim JD_WT As Double
Dim JD As Double
Dim JD0 As Double
Dim JDE As Double
Dim hJaar As Long
Dim maangeg As tMaanGeg
Dim ddate As tDatum

ddate.jj = frmPlanets.Year
Call WeekDate(ddate.jj * 100 + 1, ddate)
'dDate.MM = 1
'dDate.DD = 1
JD0 = KalenderNaarJD(ddate)
ddate.jj = frmPlanets.Year + 1
Call WeekDate(ddate.jj * 100 + 1, ddate)
JDE = KalenderNaarJD(ddate)
Call Zomertijd_Wintertijd(frmPlanets.Year, JD_ZT, JD_WT)
'{JD_ZT/JD_WT zijn berekend voor 0h UT}
JD_ZT = JD_ZT + 2 / 24 '{= 3h WT}
JD_WT = JD_WT + 1 / 24 '{= 3h ZT}
blnDoorgaan = True
Me.txtVerschijnselen.Text = "M1 = Mimas, M2 = Enceladus, M3 = Tethys, M4 = Dione, M5 = Rhea, M6 = Titan" & vbCrLf
For maannr = 1 To 6
    hJaar = frmPlanets.Year
    Call Zomertijd_Wintertijd(hJaar, JD_ZT, JD_WT)
    JD = JD0 - (360 / MaanBewPerDag(maannr))
    H = 0
    Do While JD < JDE And blnDoorgaan
        Me.PBVoortgang = Max(Int(100 * ((JD - JD0) / (JDE - JD0))), 0)
        DoEvents
        Call Bepaal_Elongatie(JD, maangeg, maannr, H, JD)
        If (JD > JD0) And (JD < JDE + 1) Then
             ddate = JDNaarKalender(JD)
             bZichtAlle = False
             If Zichtbaar(JD) Or bZichtAlle Then
                  If ddate.jj <> hJaar Then
                      hJaar = ddate.jj
                      Call BepaalZT_WT(ddate.jj, JD_ZT, JD_WT)
                  End If
                  ddate = JDNaarKalender(JD - TijdCorrectie(JD, JD_ZT, JD_WT))
                  st = StrDate(ddate) + " " + StrHMS(Frac(ddate.DD) * Pi2, 1)
        '          StrHMS(Frac(dDate.DD) * Pi2, 4)
                  'writeDate(Date) write(' ') writeHMS(frac(Date.dd)*pi2, 1)
                  #If FRANS Then
                    Select Case H
                         Case 0:     st = st + " H" 'Haut / Boven
                         Case 90:    st = st + " E" 'Est / Oost
                         Case 180:   st = st + " B" 'Bas / Beneden
                         Case 270:   st = st + " O" 'Est / West
                    End Select
                  #Else
                    Select Case H
                         Case 0:     st = st + " B"
                         Case 90:    st = st + " E"
                         Case 180:   st = st + " O"
                         Case 270:   st = st + " W"
                    End Select
                  #End If
                  st = st + " M" + Format(maannr, "0")
                  Me.txtVerschijnselen.Text = Me.txtVerschijnselen.Text + vbCrLf + st
             End If
        End If
        H = H + 90
        If H = 360 Then H = 0
        JD = JD + 90 / MaanBewPerDag(maannr)
    Loop
Next
End Sub
Function Max(x, Y)
If x > Y Then
    Max = x
Else
    Max = Y
End If
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
blnDoorgaan = False
End Sub

