Attribute VB_Name = "modWeek"
'(********************************************************
'* Weeknummers volgens ISO-standaard 2015                *
'* 23 juli 1994 (c) D.A.M. Molenkamp                     *
'*********************************************************)


Function Schrikkeljaar(dDate As tDatum) As Boolean

With dDate
  If (.jj Mod 400 = 0) Or ((.jj Mod 100 <> 0) And (.jj Mod 4 = 0)) Then
      Schrikkeljaar = True
  Else
      Schrikkeljaar = False
  End If
End With
End Function


Function DayOfYear2(dDate As tDatum) As Long
    DayOfYear2 = Int(KalenderNaarJD(dDate) - JaarNaarJD0(dDate.jj) + 0.001)
End Function

Function WeekOfYear(dDate As tDatum) As Long


Dim hDate As tDatum
Dim nJan1_ditjaar, nWeeknr, nJan1_vorigjaar, nDagditjaar, nJaar, _
    nAantdagen_vorigjaar, tmp As Long

    hDate = dDate
    hDate.mm = 1
    hDate.DD = 1
    nJaar = dDate.jj
    nJan1_ditjaar = DagVanWeek(KalenderNaarJD(hDate))
    nDagditjaar = DayOfYear2(dDate)

    If nJan1_ditjaar <= 4 Then '{jan1 valt op ma, di, wo of do}
        nWeeknr = 1 + Int((nDagditjaar + nJan1_ditjaar - 2) / 7)
        hDate.jj = dDate.jj + 1
        If (nWeeknr = 53) And (DagVanWeek(KalenderNaarJD(hDate)) <= 4) Then
             nWeeknr = 1
             nJaar = nJaar + 1
        End If
    Else   '{nJan1_ditjaar>4  dus vr, za, zo}
       nWeeknr = Int((nDagditjaar + nJan1_ditjaar - 2) / 7)
       Select Case nWeeknr
            Case 53
               nWeeknr = 1
               nJaar = nJaar + 1
            Case 0
               hDate.jj = dDate.jj - 1
               nJan1_vorigjaar = DagVanWeek(KalenderNaarJD(hDate))
               If Schrikkeljaar(hDate) Then
                   nAantdagen_vorigjaar = 366
               Else
                   nAantdagen_vorigjaar = 365
               End If
               nWeeknr = Int((nAantdagen_vorigjaar + nDagditjaar + nJan1_vorigjaar - 2) / 7)
               If nJan1_vorigjaar <= 4 Then
                  nWeeknr = nWeeknr + 1
               End If
               nJaar = nJaar - 1
       End Select
    End If
WeekOfYear = Int(100# * nJaar) + nWeeknr
End Function


Sub WeekDate(Week As Long, ByRef dDate As tDatum)

Dim weeknr, jr, dat1jan, dagweek1jan, Week1janvjaar As Long
Dim hDate As tDatum, JD As Double

    hDate = dDate
    weeknr = Week Mod 100
    jr = Int(Week / 100)

    hDate.jj = jr + 1
    hDate.mm = 1
    hDate.DD = 1

    Week1janvjaar = WeekOfYear(hDate)
    If (weeknr < 0) Or (weeknr > 53) Or ((weeknr = 53) And (Not (Week1janvjaar = Week))) Then
       dDate.DD = -1
    Else
        hDate.jj = jr
        JD = KalenderNaarJD(hDate)
        dagweek1jan = DagVanWeek(JD)

        If dagweek1jan <= 4 Then '{1jan in dit jaar, anders in vorig jaar}
          JD = JD + 1 - dagweek1jan + 7 * (weeknr - 1)
        Else
          JD = JD + 1 - dagweek1jan + 7 * weeknr
        End If
        dDate = JDNaarKalender(JD)
    End If
End Sub

Sub Zomertijd_Wintertijd(Jaar As Long, ByRef JD_ZT As Double, ByRef JD_WT As Double)

Dim Datum As tDatum, dagweek As Long

    Datum.jj = Jaar
    Datum.mm = 3 '{de laatste zondag van maart}
    Datum.DD = 31
'    Datum.H = 0
'    Datum.m = 0
'    Datum.s = 0

    JD_ZT = KalenderNaarJD(Datum)
    dagweek = DagVanWeek(JD_ZT)
    If dagweek < 7 Then
      JD_ZT = JD_ZT - dagweek
    End If
    Datum.jj = Jaar
    If Jaar < 1996 Then
      Datum.mm = 9 '{de laatste zondag van september}
      Datum.DD = 30
    Else
      Datum.mm = 10 '{de laatste zondag van oktober}
      Datum.DD = 31
    End If
    JD_WT = KalenderNaarJD(Datum)
    dagweek = DagVanWeek(JD_WT)
    If dagweek < 7 Then
      JD_WT = JD_WT - dagweek
    End If
End Sub

Function TijdCorrectie(JD, JD_ZT, JD_WT As Double) As Double
    If (JD > JD_ZT) And (JD <= JD_WT) Then
      TijdCorrectie = -2 / 24#
    Else
      TijdCorrectie = -1 / 24#
    End If
End Function

Sub BepaalZT_WT(Jaar As Long, ByRef JD_ZT As Double, ByRef JD_WT As Double)
Dim Datum As tDatum, dagweek As Long

    Datum.jj = Jaar
    Datum.mm = 3 '{de laatste zondag van maart}
    Datum.DD = 31

    JD_ZT = KalenderNaarJD(Datum)
    dagweek = DagVanWeek(JD_ZT)
    If dagweek < 7 Then
      JD_ZT = JD_ZT - dagweek
    End If
    Datum.jj = Jaar
    If Jaar < 1996 Then
      Datum.mm = 9 '{de laatste zondag van september}
      Datum.DD = 30
    Else
      Datum.mm = 10 '{de laatste zondag van oktober}
      Datum.DD = 31
    End If
    JD_WT = KalenderNaarJD(Datum)
    dagweek = DagVanWeek(JD_WT)
    If dagweek < 7 Then
      JD_WT = JD_WT - dagweek
    End If
    JD_ZT = JD_ZT + 2 / 24 '{= 3h WT}
    JD_WT = JD_WT + 1 / 24 '{= 3h ZT}
End Sub

