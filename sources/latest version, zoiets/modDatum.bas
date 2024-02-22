Attribute VB_Name = "modDatum"
'datumroutines, gebaseerd op de juliaansedatum die veel in de astronomie wordt gebruikt.
'startdatum 1 januari -4712, 12h (dus 's middags)
'Datum-routines : v1.0
'datum          : 10 november 2000
'auteur         : Dominique Molenkamp
'iso-norm 2015
Type tDatum
    jj As Long
    MM As Integer
    DD As Double
End Type
Function KalenderNaarJD(Datum As tDatum) As Double
Dim A As Integer, B As Integer, m As Integer, j As Integer
Dim D As Double

D = Datum.DD
If (Datum.MM > 2) Then
    j = Datum.jj
    m = Datum.MM
    
Else
    j = Datum.jj - 1
    m = Datum.MM + 12
End If

A = Int(j / 100)

If (Datum.jj < 1582) _
Or ((Datum.jj = 1582) And ((Datum.MM < 10) _
Or ((Datum.MM = 10) And (Datum.DD <= 4)))) Then
    B = 0
Else
    B = 2 - A + Int(A / 4)
End If
KalenderNaarJD = Int(365.25 * (j + 4716)) + Int(30.6001 * (m + 1)) + D + B - 1524.5
End Function

Function JaarNaarJD0(ByVal Jaar As Integer) As Double
Dim A As Integer, j As Integer

j = Jaar - 1
If (Jaar > 1582) Then
    A = Int(j / 100)
Else
    A = 2
End If
JaarNaarJD0 = Int(365.25 * j) - A + Int(A / 4) + 1721424.5
End Function

Function JDNaarKalender(JD As Double) As tDatum

Dim A As Double, B As Double, f As Double
Dim alpha As Integer, c As Integer, E As Integer
Dim D As Long, Z As Long
Dim Datum As tDatum

Z = Int(JD + 0.5)
f = (JD + 0.5) - Z
If Z < 2299161# Then
    A = Z
Else
    alpha = Int((Z - 1867216.25) / 36524.25)
    A = Z + 1 + alpha - Int(alpha / 4)
End If
B = A + 1524
c = Int((B - 122.1) / 365.25)
D = Int(365.25 * c)
E = Int((B - D) / 30.6001)
Datum.DD = B - D - Int(30.6001 * E) + f
If E < 14 Then
    Datum.MM = E - 1
Else
    Datum.MM = E - 13
End If
If Datum.MM > 2 Then
    Datum.jj = c - 4716
Else
    Datum.jj = c - 4715
End If
JDNaarKalender = Datum
End Function

Function DagVanWeek(ByVal JD As Double) As Integer
JD = Int(JD + 0.5)
DagVanWeek = Int(JD - Int(JD / 7) * 7) + 1
End Function


Function DagVanJaar(Datum As tDatum) As Integer

Dim k As Integer

If Schrikkeljaar(Datum.jj) Then
    k = 1
Else
    k = 2
End If
DagVanJaar = Int((275 * Datum.MM / 9)) - k * Int((Datum.MM + 9) / 12) + Int(Datum.DD) - 30
End Function

Function DatumPasen(Jaar As Long) As tDatum
Dim Datum As tDatum
Dim A As Integer, B As Integer, c As Integer, D As Integer, E As Integer, f As Integer, G As Integer, H As Integer, i As Integer, k As Integer, l As Integer, m As Integer, tmp As Integer

Datum.jj = Jaar
If Jaar < 1583 Then
'  { Juliaanse Pasen }
    A = Jaar Mod 4
    B = Jaar Mod 7
    c = Jaar Mod 19
    D = (19 * c + 15) Mod 30
    E = (2 * A + 4 * B - D + 34) Mod 7
    tmp = D + E + 114
Else
'  { Gregoriaanse Pasen }
    A = Jaar Mod 19
    B = Int(Jaar / 100)
    c = Jaar Mod 100
    D = Int(B / 4)
    E = B Mod 4
    f = Int((B + 8) / 25)
    G = Int((B - f + 1) / 3)
    H = (19 * A + B - D - G + 15) Mod 30
    i = Int(c / 4)
    k = c Mod 4
    l = (32 + 2 * E + 2 * i - H - k) Mod 7
    m = Int((A + 11 * H + 22 * l) / 451)
    tmp = H + l - 7 * m + 114
End If
Datum.MM = Int(tmp / 31)
Datum.DD = (tmp Mod 31) + 1
DatumPasen = Datum
End Function
Function Schrikkeljaar(Jaar As Long) As Boolean
If (Jaar Mod 4 = 0) And ((Jaar < 1583) Or (Jaar Mod 100 <> 0) _
                                               Or (Jaar Mod 400 = 0)) Then
    Schrikkeljaar = True
Else
    Schrikkeljaar = False
End If
End Function


Function DagVanJaar2(Datum As tDatum) As Integer
    DagVanJaar2 = Int(KalenderNaarJD(Datum)) - JaarNaarJD0(Datum.jj) + 0.001
End Function

Function WeekVanJaar(Datum As tDatum) As Long


Dim hDatum As tDatum
Dim nJan1_ditjaar As Integer, nWeeknr As Integer, nJan1_vorigjaar As Integer, nDagditjaar As Integer, nJaar As Integer
Dim nAantdagen_vorigjaar As Integer, tmp As Integer

hDatum = Datum
hDatum.MM = 1
hDatum.DD = 1
nJaar = Datum.jj
nJan1_ditjaar = DagVanWeek(KalenderNaarJD(hDatum))
nDagditjaar = DagVanJaar2(Datum)

If nJan1_ditjaar <= 4 Then '{jan1 valt op ma, di, wo of do}
    nWeeknr = 1 + Int((nDagditjaar + nJan1_ditjaar - 2) / 7)
    hDatum.jj = Datum.jj + 1
    If (nWeeknr = 53) And (DagVanWeek(KalenderNaarJD(hDatum)) <= 4) Then
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
        hDatum.jj = Datum.jj - 1
        nJan1_vorigjaar = DagVanWeek(KalenderNaarJD(hDatum))
        If Schrikkeljaar(hDatum.jj) Then
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
WeekVanJaar = Int(100# * nJaar) + nWeeknr
End Function


Function DatumWeek(Week As Long) As tDatum

Dim weeknr As Long, jr As Long, dat1jan As Long, dagweek1jan As Long, Week1janvjaar As Long
Dim Datum As tDatum, hDatum As tDatum
Dim JD As Double
weeknr = Week Mod 100
jr = Int(Week / 100)

hDatum.jj = jr + 1
hDatum.MM = 1
hDatum.DD = 1
Datum = hDatum
    
Week1janvjaar = WeekVanJaar(hDatum)
If (weeknr < 0) Or (weeknr > 53) Or _
   ((weeknr = 53) And (Not (Week1janvjaar = Week))) Then
    Date.DD = -1
Else
    hDatum.jj = jr
    JD = KalenderNaarJD(hDatum)
    dagweek1jan = DagVanWeek(JD)
    If dagweek1jan <= 4 Then '{1 jan in dit jaar, anders in vorig jaar}
        JD = JD + 1 - dagweek1jan + 7 * (weeknr - 1)
    Else
        JD = JD + 1 - dagweek1jan + 7 * weeknr
    End If
    Datum = JDNaarKalender(JD)
End If
DatumWeek = Datum
End Function
Function JDToT(JD As Double) As Double
JDToT = (JD - 2451545) / 36525
End Function
Function TToJD(T As Double) As Double
TToJD = T * 36525 + 2451545#
End Function

Sub test()
Dim Datum As tDatum
Dim weekjr As Long, JD As Double
Datum.jj = 1
Datum.MM = 2
Datum.DD = 28
JD = KalenderNaarJD(Datum)
'jd = jd + 1
'JD = 0
Datum = JDNaarKalender(JD)
MsgBox Datum.jj & "-" & Datum.MM & "-" & Datum.DD

weekjr = WeekVanJaar(Datum)
MsgBox weekjr
MsgBox DagVanJaar(Datum)
Datum = DatumWeek(weekjr)
MsgBox Datum.jj & "-" & Datum.MM & "-" & Datum.DD
Datum = DatumPasen(2001)
MsgBox Datum.jj & "-" & Datum.MM & "-" & Datum.DD

Dim A(2) As Variant
Dim x As Variant
Dim Y As Variant
A(1) = Array(1, 2, 3, 4, 5, 6)
A(2) = Array(9)
x = A(1)
Y = A(2)
End Sub





