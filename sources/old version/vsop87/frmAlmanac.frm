VERSION 5.00
Begin VB.Form frmAlmanac 
   Caption         =   "Almanac"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11355
   Icon            =   "frmAlmanac.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   11355
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtAlmanac 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   11055
   End
End
Attribute VB_Name = "frmAlmanac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(*****************************************************************************)
'(*                                                                           *)
'(*                  Copyright (c) 1991-1992 by Jeffrey Sax                   *)
'(*                            All rights reserved                            *)
'(*                        Published and Distributed by                       *)
'(*                           Willmann-Bell, Inc.                             *)
'(*                             P.O. Box 35025                                *)
'(*                        Richmond, Virginia 23235                           *)
'(*                Voice (804) 320-7016 FAX (804) 272-5920                    *)
'(*                                                                           *)
'(*                                                                           *)
'(*                NOTICE TO COMMERCIAL SOFTWARE DEVELOPERS                   *)
'(*                                                                           *)
'(*        Prior to distributing software incorporating this code             *)
'(*        you MUST write Willmann-Bell, Inc. at the above address            *)
'(*        for validation of your book's (Astronomical Algorithms             *)
'(*        by Jean Meeus) and software Serial Numbers.  No additional         *)
'(*        fees will be required BUT you MUST have the following              *)
'(*        notice at the start of your program(s) as                            *)
'(*                                                                           *)
'(*                    This program contains code                             *)
'(*              Copyright (c) 1991-1992 by Jeffrey Sax                       *)
'(*              and Distributed by Willmann-Bell, Inc.                       *)
'(*                         Serial #######                                    *)
'(*                                                                           *)
'(*****************************************************************************)
'(* Name as   ALMANAC                                                           *)
'(* Module as ALMANAC.PAS                                                       *)
'(* Type as   Main Program                                                      *)
'(* Purpose:calculate instances of various phenomena for a given Jaar.        *)
'(* Version 2.0                                                               *)
'(* Last modified as October 1, 1992                                            *)
'(*****************************************************************************)



Dim Jaar As Long
Dim JDStart As Double, JDEnd As Double, deltaT As Double, TimeZone As Double
Dim EventList(1 To 200) As TEVENT
Dim NoEvents As Integer
Dim LunarEclipse As LUNARECLIPSEDATA
Dim SolarEclipse As SOLARECLIPSEDATA
Dim JD_ZT As Double, JD_WT As Double
Dim PlanetName '(0 To 8) As String * 10
Dim Description '(1 To 28) As String * 60

Sub FindEvents(JDStart As Double, JDEnd As Double)
Dim i As Long, j As Long, Planet As Long, nEvent As Long, k0 As Long
Dim k As Long
Dim t, JD As Double
Dim ddate As tDatum
Dim x   As Double, Dummy As Double, k1 As Double, Par As Double, mDec As Double
Dim s As String * 60
Dim plPhenom As New clsPlPhenom
Dim perApg As New clsPerApg
Dim moonnode As New clsMoonNode
Dim moonphase As New clsMoonPhase
Dim maxdec As New clsMaxDec

#If FRANS Then
    PlanetName = Array("Soleil", "Mercure", "Vénus", "Terre", "Mars", "Jupiter", "Saturne", _
        "Uranus", "Neptune")
    Description = Array("", "Printemps equinoxe", "Été solstice", "Automne equinoxe", "Hiver solstice", _
        " in inférieur conjonction avec la Soleil", " in supérieur conjonction avec la Soleil", _
        " in opposition avec la Soleil", " in conjonction avec la Soleil", _
        " in grand de l'est a ", " in grand d'ouest a ", _
        "Nouvelle lune ", "Premier quartier", "Pleine lune", "Dernier quartier", _
        "Éclipse de soleil totale", "Éclipse de soleil annulaire", "Eclipse de soleil annulaire-totale", _
        "Éclipse de soleil totale non-central", "'Eclipse de soleil annulaire non-central", _
        "Éclipse de soleil partielle, grande magn. = ", "Éclipse de lune totalle, magnitude = ", _
        "Éclipse de lune partielle, magnitude = ", "Éclipse de lune penumbrale, magnitude = ", _
        "Périgée, distance = ", "Apogée, distance = ", _
        "Passage Lune dans noeud ascendant", "Passage Lune dans noeud descendant", _
        "Maximum declination = ")
#Else
    PlanetName = Array("Sun", "Mercury", "Venus", "Earth", "Mars", "Jupiter", "Saturn", _
        "Uranus", "Neptune")
    Description = Array("", "Spring equinox", "Summer solstice", "Fall equinox", "Winter solstice", _
        " in inferior conjunction with the Sun", " in superior conjunction with the Sun", _
        " in opposition with the Sun", " in conjunction with the Sun", _
        " in greatest eastern elongation as ", " in greatest western elongation as ", _
        "New Moon", "First quarter", "Full Moon", "Last quarter", _
        "Total solar eclipse", "Annular solar eclipse", "Annular-total solar eclipse", _
        "Non-central total solar eclipse", "Non-central annular solar eclipse", _
        "Partial solar eclipse, greatest magn. = ", "Total lunar eclipse, magnitude = ", _
        "Partial lunar eclipse, magnitude = ", "Penumbral lunar eclipse, magnitude = ", _
        "Perigee, distance = ", "Apogee, distance = ", _
        "Passage Moon through ascending node", "Passage Moon through descending node", _
        "Maximum declination = ")
#End If

'{ First find equinoxes and solstices }
'txtAlmanac.Text = "Looking for equinoxes and solstices..." + vbCrLf
For i = 1 To 4
    JD = EquinoxSolstice(Jaar, i - 1)
    EventList(i).JD = JD
    EventList(i).Description = Description(i)
    EventList(i).Precision = 7
Next
NoEvents = 4

'{ Find solar and lunar eclipses }
'txtAlmanac.Text = txtAlmanac.Text + "Looking for solar and lunar eclipses..." + vbCrLf
'{ Solar eclipses }
JD = JDStart - 1
Do Until JD > JDEnd
  Call NextSolarEclipse(JD, SolarEclipse)
  With SolarEclipse
    If (.JD > JDStart) And (.JD < JDEnd) Then
      nEvent = .EclipseType
      NoEvents = NoEvents + 1
      EventList(NoEvents).JD = .JD
      EventList(NoEvents).Precision = 3
      If (nEvent <> PARTIAL) Then
        EventList(NoEvents).Description = Description(14 + nEvent)
      Else
        s = Format(.Maxmag, "0.000")
        EventList(NoEvents).Description = Description(14 + nEvent) + s + " (Maximum)"
      End If
    End If
    JD = SolarEclipse.JD
  End With
Loop
'{ Lunar eclipses }
JD = JDStart - 1
Do Until JD > JDEnd
  Call NextLunarEclipse(JD, 0, LunarEclipse)
  With LunarEclipse
  If (.JD > JDStart) And (.JD < JDEnd) Then
    nEvent = .EclipseType
    NoEvents = NoEvents + 1
    EventList(NoEvents).JD = .JD
    EventList(NoEvents).Precision = 3
    Select Case nEvent
        Case TOTAL:     i = 21
        Case PARTIAL:   i = 22
        Case PENUMBRAL: i = 23
    End Select
    If i < 23 Then
      s = Format(.MagUmbra, "0.000")
    Else
      s = Format(.MagPenumbra, "0.000")
    End If
    EventList(NoEvents).Description = Trim(Description(i)) + " " + Trim(s) + " (Maximum)"
    If i <= 23 Then
        NoEvents = NoEvents + 1
        EventList(NoEvents).JD = .JD - .SpartPenumbra
        EventList(NoEvents).Precision = 3
        #If FRANS Then
            EventList(NoEvents).Description = "    Première contact avec penumbra"
        #Else
            EventList(NoEvents).Description = "    First contact with the penumbra"
        #End If
        NoEvents = NoEvents + 1
        EventList(NoEvents).JD = .JD + .SpartPenumbra
        EventList(NoEvents).Precision = 3
        #If FRANS Then
            EventList(NoEvents).Description = "    Dernière contact avec penumbra"
        #Else
            EventList(NoEvents).Description = "    Last contact with the penumbra"
        #End If
    End If
    If i <= 22 Then
        NoEvents = NoEvents + 1
        EventList(NoEvents).JD = .JD - .SpartUmbra
        EventList(NoEvents).Precision = 3
        #If FRANS Then
            EventList(NoEvents).Description = "    Première contact avec umbra"
        #Else
            EventList(NoEvents).Description = "    First contact with the umbra"
        #End If
        NoEvents = NoEvents + 1
        EventList(NoEvents).JD = .JD + .SpartUmbra
        EventList(NoEvents).Precision = 3
        #If FRANS Then
            EventList(NoEvents).Description = "    Dernière contact avec umbra"
        #Else
            EventList(NoEvents).Description = "    Last contact with the umbra"
        #End If
    End If
    If i <= 21 Then
        NoEvents = NoEvents + 1
        EventList(NoEvents).JD = .JD - .StotUmbra
        EventList(NoEvents).Precision = 3
        #If FRANS Then
            EventList(NoEvents).Description = "    Départ éclipse totale"
        #Else
            EventList(NoEvents).Description = "    Beginning of total eclipse"
        #End If
        NoEvents = NoEvents + 1
        EventList(NoEvents).JD = .JD + .StotUmbra
        EventList(NoEvents).Precision = 3
        #If FRANS Then
            EventList(NoEvents).Description = "    Fin éclipse totale"
        #Else
            EventList(NoEvents).Description = "    End of total eclipse"
        #End If
    End If
    End If
    JD = LunarEclipse.JD
    End With
Loop

'{ Next, find planetary oppositions and conjunctions }
'txtAlmanac.Text = txtAlmanac.Text + "Looking for planetary conjunctions and oppositions..." + vbCrLf
For Planet = 1 To 8
  If Planet <> 3 Then
    For nEvent = OPPOSITION To CONJUNCTION
      k = plPhenom.FirstkOfYear(Jaar, Planet, nEvent) - 1
      JD = plPhenom.ConjunctionOpposition(Planet, nEvent, k)
      Do Until (JD > JDEnd)
          If ((JD > JDStart) And (JD < JDEnd)) Then
            NoEvents = NoEvents + 1
            EventList(NoEvents).JD = JD
            If Planet > 3 Then
              j = 7 + nEvent
            Else
              j = 5 + nEvent
            End If
            EventList(NoEvents).Description = PlanetName(Planet) + Description(j)
            EventList(NoEvents).Precision = 1
          End If
          k = k + 1
         JD = plPhenom.ConjunctionOpposition(Planet, nEvent, k)
      Loop
    Next
  End If
Next

'{ Now, find times of extreme elongations of Mercury and Venus }
'txtAlmanac.Text = txtAlmanac.Text + "Looking for extreme elongations of Mercury and Venus..." + vbCrLf
For Planet = 1 To 2
  k0 = plPhenom.FirstkOfYear(Jaar, Planet, INFCONJ) - 1
  For nEvent = EASTELONGATION To WESTELONGATION
    k = k0
    JD = plPhenom.GreatestElongation(Planet, nEvent, k, x)
    Do Until (JD > JDEnd)
      If ((JD > JDStart) And (JD < JDEnd)) Then
        NoEvents = NoEvents + 1
        EventList(NoEvents).JD = JD
        j = 9 + nEvent
        EventList(NoEvents).Description = PlanetName(Planet) + Description(j)
        s = Format(x, "00.0")
        EventList(NoEvents).Description = Trim(EventList(NoEvents).Description) + " " + s
        EventList(NoEvents).Description = Trim(EventList(NoEvents).Description) + "ø"
        EventList(NoEvents).Precision = 1
      End If
      k = k + 1
    JD = plPhenom.GreatestElongation(Planet, nEvent, k, x)
    Loop
    Next
Next

'{ Next, Perigee and Apogee }
'txtAlmanac.Text = txtAlmanac.Text + "Looking for Perigee and Apogee..." + vbCrLf
k1 = perApg.FirstkPerApgOfYear(Jaar) - 1
Call perApg.perApg(k1, JD, Par)
Do Until (JD > JDEnd)
   If (JD > JDStart) And (JD < JDEnd) Then
    NoEvents = NoEvents + 1
    EventList(NoEvents).JD = JD
    If Abs(k1 - Int(k1)) < 0.2 Then
       EventList(NoEvents).Description = Description(24)
    Else
       EventList(NoEvents).Description = Description(25)
    End If
    s = Format(moonphase.MoonDistance(Par), "000000.0")
    EventList(NoEvents).Description = Trim(EventList(NoEvents).Description) + " " + Trim(s) + " km"
    EventList(NoEvents).Precision = 3
   End If
   nEvent = nEvent + 1
  k1 = k1 + 0.5
  Call perApg.perApg(k1, JD, Par)
Loop

'{ Next, Moon through nodes }
'txtAlmanac.Text = txtAlmanac.Text + "Looking Passage Moon through Nodes..." + vbCrLf
k1 = moonnode.FirstkMoonNodeOfYear(Jaar) - 1
Call moonnode.moonnode(k1, JD)
Do Until (JD > JDEnd)
   If (JD > JDStart) And (JD < JDEnd) Then
    NoEvents = NoEvents + 1
    EventList(NoEvents).JD = JD
    If Abs(k1 - Int(k1)) < 0.2 Then
       EventList(NoEvents).Description = Description(26)
    Else
       EventList(NoEvents).Description = Description(27)
    End If
    EventList(NoEvents).Precision = 3
    End If
    nEvent = nEvent + 1
    k1 = k1 + 0.5
    Call moonnode.moonnode(k1, JD)
Loop

'{ Next, Maximum Declination }
'txtAlmanac.Text = txtAlmanac.Text + "Looking for Northern Maximum Declination of Moon..." + vbCrLf
k = maxdec.FirstkMaxDecOfYear(Jaar)
Call maxdec.maxdec(k, True, JD, mDec)
Do Until JD > JDEnd
   If (JD > JDStart) And (JD < JDEnd) Then
    NoEvents = NoEvents + 1
    EventList(NoEvents).JD = JD
    EventList(NoEvents).Description = Description(28)
    s = StrHMS_DMS(mDec, 7, 0, True, False, "g", 3)
    EventList(NoEvents).Description = Trim(EventList(NoEvents).Description) + " " + s
    EventList(NoEvents).Precision = 3
    End If
    nEvent = nEvent + 1
    k = k + 1
    Call maxdec.maxdec(k, True, JD, mDec)
Loop

'txtAlmanac.Text = txtAlmanac.Text + "Looking for Maximum Southern Declination of Moon..." + vbCrLf
k = maxdec.FirstkMaxDecOfYear(Jaar)
Call maxdec.maxdec(k, False, JD, mDec)
Do Until JD > JDEnd
   If (JD > JDStart) And (JD < JDEnd) Then
    NoEvents = NoEvents + 1
    EventList(NoEvents).JD = JD
    EventList(NoEvents).Description = Description(28)
    s = StrHMS_DMS(mDec, 7, 0, True, False, "g", 3)
    EventList(NoEvents).Description = Trim(EventList(NoEvents).Description) + " " + s
    EventList(NoEvents).Precision = 3
   End If
   nEvent = nEvent + 1
   k = k + 1
    Call maxdec.maxdec(k, False, JD, mDec)
Loop

'{ Next, lunar phases }
'txtAlmanac.Text = txtAlmanac.Text + "Looking for lunar phases..." + vbCrLf
k = Round(Int(4 * (Jaar - 2000) * 12.3685 - 2.1))
nEvent = k Mod 4
k = Int(k / 4)
If nEvent < 0 Then
    nEvent = nEvent + 4
    k = k - 1
End If
JD = moonphase.moonphase(k, nEvent)
Do Until JD > JDEnd
  If (JD > JDStart) And (JD < JDEnd) Then
    NoEvents = NoEvents + 1
    EventList(NoEvents).JD = JD
    EventList(NoEvents).Description = Description(11 + nEvent)
    EventList(NoEvents).Precision = 3
    End If
  nEvent = nEvent + 1
  If nEvent > LASTQUARTER Then
    nEvent = NEWMOON
    k = k + 1
  End If
  JD = moonphase.moonphase(k, nEvent)
Loop
End Sub

Sub SortEvents()
Dim nEvent As TEVENT
Dim i As Integer, j As Integer, iMin As Integer
Dim Min As Double
'txtAlmanac.Text = txtAlmanac.Text + "Now sorting the " + Str(NoEvents) + " astronomical events..." + vbCrLf
'{ Sorting this few events, it doublely isn't worth using a QuickSort }
For i = 1 To NoEvents - 1
  For j = i + 1 To NoEvents
    If (EventList(j).JD < EventList(i).JD) Then
      nEvent = EventList(i)
      EventList(i) = EventList(j)
      EventList(j) = nEvent
    End If
  Next
Next
End Sub

Sub PrintEvents()
Dim i As Integer
Dim deltaT As Double
Dim ddate As tDatum
Dim ch As String * 1

Call Zomertijd_Wintertijd(Jaar, JD_ZT, JD_WT)

deltaT = ApproxDeltaT(JDToT((JDStart + JDEnd) / 2)) / 86400
For i = 1 To NoEvents
  TimeZone = TijdCorrectie(EventList(i).JD, JD_ZT, JD_WT)
'{  TimeZone=0}
  ddate = JDNaarKalender(EventList(i).JD - deltaT - TimeZone)
  sTime = StrHMS_DMS((ddate.DD - Int(ddate.DD)) * 360, EventList(i).Precision, 0, False, False, "h", 2)
  If EventList(i).Precision = 1 Then sTime = sTime + "      "
  If EventList(i).Precision = 3 Then sTime = sTime + "   "
'  sTime = StrHMS((dDate.DD - Int(dDate.DD)) * 2 * pi, 3)
  sDate = StrDate(ddate)
  txtAlmanac.Text = txtAlmanac.Text + sDate + " " + sTime
  txtAlmanac.Text = txtAlmanac.Text + " " + EventList(i).Description + vbCrLf
Next
End Sub

Sub Initialize()
Jaar = Val(frmPlanets.Year)
'{TimeZone = Readdouble('Enter time zone (LT - UT) as ') / 24}
JDStart = JaarNaarJD0(Jaar) + 1 - 1 / 24# '{ Add 1 as we need Jan 1st, not Jan 0.0 ! }
JDEnd = JaarNaarJD0(Jaar + 1) + 1 - 1 / 24#
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim stext As String
    If KeyCode = 67 And Shift = 2 Then
       stext = Me.txtAlmanac
       Clipboard.Clear
       Clipboard.SetText (stext)
    End If
End Sub

Private Sub Form_Load()
#If FRANS Then
    Me.Caption = "Almanach"
#End If
Initialize
Call FindEvents(JDStart, JDEnd)
SortEvents
PrintEvents
End Sub

