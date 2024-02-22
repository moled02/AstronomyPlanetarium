Attribute VB_Name = "modtest"
Type tSter
'    saonum As String * 6
'    sterbld As String * 3
    a As Double
    d As Double
    M As Integer
    flamsteed As Byte
    bayer As String * 2
End Type
Type tlijn
'    saonum As String * 6
'    sterbld As String * 3
    ster1 As tSter
    ster2 As tSter
End Type
Const Pi = 3.14159265358979
Sub T()
Dim objPerApg As New clsPlPhenom
Dim JDE As Double, Par As Double
'MsgBox objPerApg.ConjunctionOpposition(6, 1, 121)
'MsgBox objPerApg.FirstkOfYear(1883, 2, 0)
'MsgBox objPerApg.GreatestElongation(1, 1, -20, Par)

'MsgBox objPerApg.ConjunctionOpposition(2, 0, -74)
'Debug.Print JDE, Par
MsgBox EquinoxSolstice(1962, 1)

Dim T As Double, BessElmt As tBessElmt, AuxElmt As tAuxElmt, dBess As tDiffBess, PredData As tPredData, _
                   Extr As tExtremes


'Call Extremes(-0.388751996349528, BessElmt, AuxElmt, DBess, PredData, Extr)
Call modEclipse.TestEclipse(-0.388751996349528)
Call modEclipse.TestEclipse(-0.388752566735113)
End Sub

Sub Main()
Dim n As Double
Dim s(34)
s(0) = "  0  9h45.9m  16°26' 10h19.4m  11°35'"
s(1) = "  1  9h47.9m  16°14' 10h19.5m  11°35'"
s(2) = "  2  9h49.8m  16°02' 10h19.6m  11°34'"
s(3) = "  3  9h51.7m  15°50' 10h19.7m  11°34'"
s(4) = "  4  9h53.6m  15°38' 10h19.8m  11°33'"
s(5) = "  5  9h55.5m  15°25' 10h19.9m  11°33'"
s(6) = "  6  9h57.4m  15°13' 10h20.0m  11°32'"
s(7) = "  7  9h59.3m  15°00' 10h20.1m  11°31'"
s(8) = "  8 10h01.2m  14°48' 10h20.2m  11°31'"
s(9) = "  9 10h03.0m  14°35' 10h20.3m  11°30'"
s(10) = " 10 10h04.9m  14°23' 10h20.4m  11°30'"
s(11) = " 11 10h06.8m  14°10' 10h20.5m  11°29'"
s(12) = " 12 10h08.6m  13°57' 10h20.6m  11°28'"
s(13) = " 13 10h10.5m  13°44' 10h20.7m  11°28'"
s(14) = " 14 10h12.4m  13°32' 10h20.8m  11°27'"
s(15) = " 15 10h14.2m  13°19' 10h20.9m  11°27'"
s(16) = " 16 10h16.1m  13°06' 10h21.0m  11°26'"
s(17) = " 17 10h17.9m  12°53' 10h21.1m  11°26'"
s(18) = " 18 10h19.8m  12°40' 10h21.2m  11°25'"
s(19) = " 19 10h21.6m  12°27' 10h21.3m  11°24'"
s(20) = " 20 10h23.4m  12°14' 10h21.4m  11°24'"
s(21) = " 21 10h25.3m  12°01' 10h21.5m  11°23'"
s(22) = " 22 10h27.1m  11°47' 10h21.6m  11°23'"
s(23) = " 23 10h28.9m  11°34' 10h21.7m  11°22'"
s(24) = " 24 10h30.7m  11°21' 10h21.8m  11°22'"
s(25) = " 25 10h32.5m  11°08' 10h21.9m  11°21'"
s(26) = " 26 10h34.4m  10°54' 10h22.0m  11°20'"
s(27) = " 27 10h36.2m  10°41' 10h22.1m  11°20'"
s(28) = " 28 10h38.0m  10°28' 10h22.2m  11°19'"
s(29) = " 29 10h39.8m  10°14' 10h22.3m  11°19'"
s(30) = " 30 10h41.6m  10°01' 10h22.4m  11°18'"
s(31) = " 31 10h43.4m   9°47' 10h22.4m  11°17'"
s(32) = " 32 10h45.2m   9°34' 10h22.5m  11°17'"
s(33) = " 33 10h46.9m   9°20' 10h22.6m  11°16'"
s(34) = " 34 10h48.7m   9°07' 10h22.7m  11°16'"


For i = 0 To 34
    r1 = (Val(Mid(s(i), 5, 2)) + Val(Mid(s(i), 8, 4)) / 60) * Pi / 180
    d1 = (Val(Mid(s(i), 15, 2)) + Val(Mid(s(i), 18, 2)) / 60) * Pi / 180
    r2 = (Val(Mid(s(i), 22, 2)) + Val(Mid(s(i), 25, 4)) / 60) * Pi / 180
    d2 = (Val(Mid(s(i), 32, 2)) + Val(Mid(s(i), 35, 2)) / 60) * Pi / 180
'    Selection.TypeText s(i) & vbTab & distance(r1, d1, r2, d2) * 180 / pi * 60 & vbTab & distance(r1, d1, r2, d2) * 180 / pi * 60 - (30.71667 - 7 / 60) / 2
    'Selection.TypeParagraph
Next
MsgBox Interpol3(9.29736105934427, -1.37911102336353, -6.51719622906835, 1)
n = 0
Call Nulpunt(-1.37911102336353, -6.51719622906835, 1.34600855991885, n, 1)
MsgBox n
Call Interpol3(-1.37911102336353, -6.51719622906835, 1.34600855991885, n)
End Sub
Function distance(r1, d1, r2, d2) As Double
distance = acos(Sin(d1) * Sin(d2) + Cos(d1) * Cos(d2) * Cos(r1 - r2))
End Function

Function Nulpunt(y1 As Double, y2 As Double, y3 As Double, n As Double, nStap As Double)
Dim yx As Double
Debug.Print y1 & vbTab & y2 & vbTab & y3 & vbTab & n
If Abs(y2) < 0.00001 Then
    Exit Function
End If
If sign(y1) <> sign(y2) Then
    n = n - nStap / 2
    yx = Interpol3(y1, y2, y3, -0.5)
    Call Nulpunt(y1, yx, y2, n, nStap / 2)
Else
    n = n + nStap / 2
    yx = Interpol3(y1, y2, y3, 0.5)
    Call Nulpunt(y2, yx, y3, n, nStap / 2)
End If
End Function
Function sign(x As Double) As Long
If x >= 0 Then
    sign = 1
Else
    sign = -1
End If
End Function





Sub modConversie()
Dim ster As tSter
Dim lijn As tlijn
Dim sRegel As String
Open "N:\DM\astro\vsop87\sterren.bin" For Output As #2
Close

Open "N:\DM\astro\vsop87\sterren.asc" For Input As #1
Open "N:\DM\astro\vsop87\sterren.bin" For Random As #2 Len = LenB(ster)
While Not EOF(1)
    Line Input #1, sRegel
'    ster.saonum = ""
'    ster.sterbld = ""
    ster.a = Val(Mid(sRegel, 3, 10))
    ster.d = Val(Mid(sRegel, 14, 10))
    ster.M = Val(Mid(sRegel, 25, 3))
    ster.flamsteed = Val(Mid(sRegel, 29, 3))
    ster.bayer = Mid(sRegel, 33, 2)
    Put #2, , ster
Wend
Close

Open "N:\DM\astro\vsop87\sterlijn.bin" For Output As #2
Close

Open "N:\DM\astro\vsop87\Sterln3.asc" For Input As #1
Open "N:\DM\astro\vsop87\sterlijn.bin" For Random As #2 Len = LenB(lijn)
While Not EOF(1)
    Line Input #1, sRegel
'    ster.saonum = ""
'    ster.sterbld = ""
    With lijn.ster1
        .a = Val(Mid(sRegel, 13, 9))
        .d = Val(Mid(sRegel, 22, 9))
        .M = Val(Mid(sRegel, 33, 3))
        .flamsteed = Val(Mid(sRegel, 37, 3))
        .bayer = Mid(sRegel, 40, 2)
    End With
    With lijn.ster2
        .a = Val(Mid(sRegel, 56, 9))
        .d = Val(Mid(sRegel, 65, 9))
        .M = Val(Mid(sRegel, 76, 3))
        .flamsteed = Val(Mid(sRegel, 80, 3))
        .bayer = Mid(sRegel, 83, 2)
    End With
    Put #2, , lijn
Wend
Close

Open "N:\DM\astro\vsop87\sterlijn.bin" For Random As #2 Len = LenB(lijn)
While Not EOF(2)
    Get #2, , lijn
    Debug.Print lijn.ster1.a
Wend
End Sub



Function NauwkeurigerTijdstipMaansverduistering(ByVal JD As Double, ByVal stype As String) As Double
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

T = JDToT(JD)
Call PositieZonMaan(T + 1 / 876600, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime)
BessElmt2.x = modpi((RkM - (RkZ + Pi))) * Cos(DecM)
eps = 0.25 * modpi((RkM - (RkZ + Pi))) * Sin(2 * -DecZ) * Sin((RkM - (RkZ + Pi)))
BessElmt2.y = modpi(DecM + DecZ + eps)
'Call Bess_elmts(RkM, DecM, ParM, RkZ + Pi, -DecZ, ParZ, RZ, AppTime, BessElmt2)
BessElmt2.x = BessElmt2.x * 180 / Pi * 3600: BessElmt2.y = BessElmt2.y * 180 / Pi * 3600

Call PositieZonMaan(T, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, AppTime)
sParZ0 = asin(0.272274 * Sin(ParM))
BessElmt1.x = modpi((RkM - (RkZ + Pi))) * Cos(DecM)
eps = 0.25 * modpi((RkM - (RkZ + Pi))) * Sin(2 * -DecZ) * Sin((RkM - (RkZ + Pi)))
BessElmt1.y = modpi(DecM + DecZ + eps)
' Call Bess_elmts(RkM, DecM, ParM, RkZ + Pi, -DecZ, ParZ, RZ, AppTime, BessElmt1)
BessElmt1.x = BessElmt1.x * 180 / Pi * 3600: BessElmt1.y = BessElmt1.y * 180 / Pi * 3600
Call DiffBess(BessElmt1, BessElmt2, dBess)

If stype = "PB" Or stype = "PE" Then
    lf = 1.02 * (0.99834 * ParM + 959.63 / 3600 * Pi / 180 / RZ)
    lL = lf + sParZ0
ElseIf stype = "UB" Or stype = "UE" Then
    lf = 1.02 * (0.99834 * ParM - 959.63 / 3600 * Pi / 180 / RZ)
    lL = lf + sParZ0
ElseIf stype = "TB" Or stype = "TE" Then
    lf = 1.02 * (0.99834 * ParM - 959.63 / 3600 * Pi / 180 / RZ)
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

