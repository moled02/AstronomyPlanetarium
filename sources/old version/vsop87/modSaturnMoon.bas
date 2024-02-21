Attribute VB_Name = "modSaturnMoon"
Option Explicit
Type tManen1_5
                     e0 As Double
                     n As Double
                     O  As Double
                     gamma As Double
                     Pi1 As Double
                     e As Double
                     a As Double
                     dl As Double
                     u As Double
                     r_a As Double
                     r As Double
                     nd As Double
                     l1 As Double
                     M As Double
                     d As Double
                     l As Double
                     x As Double
                     y As Double
                     g As Double
End Type
Type tmaanTitan
                     a As Double
                     e As Double
                     Pi As Double
                     n As Double
                     l As Double
                     O As Double
                     nn0 As Double
                     i0 As Double
                     st0 As Double
                     e0 As Double
                     p0 As Double
                     l0 As Double
                     r As Double
                     d_w As Double
                     fi1 As Double
                     gamma As Double
                     dea As Double
                     dl As Double
                     dma As Double
                     sgdd As Double
                     dgamma As Double
                     g As Double
                     g0 As Double
                     de As Double
                     dst As Double
                     di As Double
                     dp As Double
                     u As Double
                     middan As Double
                     middlong As Double
                     v As Double
                     x As Double
                     y As Double
End Type
Type tMaanGeg
        Manen(1 To 5) As tManen1_5
        Titan  As tmaanTitan
End Type

Private maangeg   As tMaanGeg
Private dDate    As tDatum
Private Planet As Long, Earth   As Long
Private ObsLon As Double, ObsLat As Double, Height As Double
Private T As Double, DTOfUT As Double, DeltaT As Double, LAST   As Double
Private Obl As Double, NutLon As Double, NutObl As Double, TimeZone   As Double
Private SHelio As TSVECTOR, SAarde As TSVECTOR, SGeo As TSVECTOR
Private RA As Double, Decl As Double
Private SaturnDeltaU As Double, SaturnB As Double, Magnitude   As Double
Public SaturnRingData As TSATURNRINGDATA
Public AltSaturnRingData As TALTSATURNRINGDATA
Private JD As Double, JD0 As Double, JDE As Double, JD_ZT As Double, JD_WT   As Double
Private i As Long, maannr As Long, h As Long
Public SemiDiameter As Double
Public PolarSemiDiameter As Double
Public MaanBewPerDag As Variant

Sub SaturnMoonInit()

MaanBewPerDag = Array(0, 381.9944421, 262.7319405, 190.69795, _
                        131.5349729, 79.690081, 22.57701508)
End Sub

Function fnmod(x As Double) As Double
    fnmod = x - 360 * Int(x / 360)
End Function

Sub ElementToSigp_P(ByVal B, ByVal u_U As Double, Sig As Double, p_P As Double)

Dim sSigsp_P As Double, sSigcp_P As Double, cSig As Double
     sSigsp_P = Sin(u_U)
     sSigcp_P = Sin(B) * Cos(u_U)
     cSig = Cos(B) * Cos(u_U)

     p_P = atan2(sSigsp_P, sSigcp_P)
     Sig = atan2(sSigsp_P / Sin(p_P), cSig)
End Sub

Sub sat_manen(ByVal JD As Double, ByVal maannr As Long, maangeg As tMaanGeg)

Dim t1 As Double, t2 As Double, t3 As Double, t4 As Double, T As Double, p11 As Double, ch1 As Double
Dim ch2 As Double, ch3 As Double, om As Double, w As Double, i As Double, phi1 As Double, afst As Double
Dim B As Double, i1 As Double
Dim k As Long
Dim s As Double, Sig As Double, p_P       As Double
Dim om_om1si1 As Double, i_i1 As Double, Xsi1 As Double, ea     As Double
Dim Xsi As Double, O_Om0 As Double, Om_Om0 As Double, GroteG As Double, l0_Pi0 As Double, Om_Om1 As Double

With maangeg
   T = JDToT(JD)
   i1 = (28.075216 - T * (0.012998 - T * 0.000004)) * DToR
   om = (169.50847 + T * (1.394681 + T * 0.000412)) * DToR
   w = AltSaturnRingData.w
  afst = SGeo.r

  p11 = Pi / 180
  t1 = (JD - 2415020 - 0.31352) / 365.24219878
  t2 = t1 + (1900 - 1889.25)
  t3 = t1 + (1900 - 1866.27)
  t4 = t1 + (1900 - 1879.59)
  ch1 = fnmod(5.0864 * (t3)) * p11
  ch2 = fnmod((63.75 + 32.51 * t2)) * p11
  ch3 = fnmod((117.28 + 93.14 * t2)) * p11

  Select Case maannr
    Case 1
       .Manen(1).e0 = 127.09166666
       .Manen(1).n = 381.994442
       .Manen(1).O = fnmod(56.1 - 365.23 * t2)
       .Manen(1).gamma = 1.5166666667
       .Manen(1).Pi1 = fnmod(105 + 365.6 * t2)
       .Manen(1).e = 0.0201
       .Manen(1).a = 255.89
       .Manen(1).dl = -44.39 * Sin(ch1) - 0.764 * Sin(3 * ch1)
    Case 2
       .Manen(2).e0 = 199.43
       .Manen(2).n = 262.7319405
       .Manen(2).O = fnmod(52 - 152.7 * t2)
       .Manen(2).gamma = 0.023333333333
       .Manen(2).Pi1 = fnmod(308.38 + 123.43 * t2)
       .Manen(2).e = 0.00444
       .Manen(2).a = 328.29
       .Manen(2).dl = 0.23983333333 * Sin(ch2) + 0.23433333333 * Sin(ch3)
    Case 3
       .Manen(3).e0 = 284.47166667
       .Manen(3).n = 190.69795
       .Manen(3).O = fnmod(110.39 - 72.25 * t2)
       .Manen(3).gamma = 1.0926666667
       .Manen(3).Pi1 = 0
       .Manen(3).e = 0
       .Manen(3).a = 406.4
       .Manen(3).dl = 2.065 * Sin(ch1) + 0.036 * Sin(3 * ch1)
    Case 4
       .Manen(4).e0 = 253.86666667
       .Manen(4).n = 131.5349729
       .Manen(4).O = fnmod(201 - 31 * t2)
       .Manen(4).gamma = 0.023333333333
       .Manen(4).Pi1 = fnmod(173.4 + 30.75 * t2)
       .Manen(4).e = 0.00221
       .Manen(4).a = 520.51
       .Manen(4).dl = -0.0155 * Sin(ch2) - 0.015166666667 * Sin(ch3)
  End Select
  If maannr < 5 Then
  ' MIMAS ENCELADUS TETHYS DIONE }
    With .Manen(maannr)
         .l1 = fnmod(.e0) + fnmod(.n * (JD - 2411093) + .dl)
         .M = .l1 - .Pi1
         .d = .O + (w - om) / p11
         .l = .l1 + (w - om) / p11
         .u = .l + 1 / p11 * .e * (2 * Sin((.M * p11)) + 5 / 4 * .e * Sin(2 * .M * p11))

         .r_a = 1 + .e * (0.5 * .e - Cos(.M * p11) - 0.5 * .e * Cos(2 * .M * p11))
         .r = .r_a + .a
         Call ElementToSigp_P(AltSaturnRingData.B, .u * p11 - AltSaturnRingData.u, Sig, p_P)
         s = .a / afst * Sin(Sig)
         .x = s * Sin(p_P + AltSaturnRingData.P)
         .y = s * Cos(p_P + AltSaturnRingData.P)
    End With
  End If
  'RHEA}
  If maannr = 5 Then
     With .Manen(maannr)
         .e0 = 358.395
         .n = 79.6900881
         om_om1si1 = 1 / 60 * (20.49 * Sin(p11 * fnmod(344.09 - 10.2 * t2)) - 0.38 + 1 * Sin(p11 * fnmod(48.5 - 0.5 * t2)))
         i_i1 = 1 / 60 * (20.49 * Cos(p11 * fnmod(344.09 - 10.2 * t2)) - 2.79 + 1 * Cos(p11 * fnmod(48.5 - 0.5 * t2)))
         .Pi1 = 276.25 + 0.53 * t2 + 17.64 * Sin(p11 * fnmod(9.5 * (t1 + 20.41)))
         .e = 0.00098 + 0.0003 * Cos(p11 * fnmod(9.5 * (t1 + 20.41)))
         .a = 726.89
         Xsi1 = atan2(om_om1si1, i_i1) / p11
         .d = Xsi1 + w / p11 + om_om1si1 / tan(i1) * (1 - 0.5 * Sin(Xsi1 * p11) * Sin(Xsi1 * p11))
         .g = om_om1si1 / Sin(Xsi1 * p11)
         .l1 = fnmod(.e0) + fnmod(.n * (JD - 2411093))
         .M = .l1 - .Pi1
         .l = .l1 + (w - om) / p11
         .u = .l + 1 / p11 * .e * (2 * Sin((.M * p11)) + 5 / 4 * .e * Sin(2 * .M * p11))
         .r_a = 1 + .e * (0.5 * .e - Cos(.M * p11) - 0.5 * .e * Cos(2 * .M * p11))
         .r = .r_a + .a
         Call ElementToSigp_P(AltSaturnRingData.B, .u * p11 - AltSaturnRingData.u, Sig, p_P)
         s = .a / afst * Sin(Sig)
         .x = s * Sin(p_P + AltSaturnRingData.P)
         .y = s * Cos(p_P + AltSaturnRingData.P)
     End With
  End If
' TITAN}
  If maannr = 6 Then
     With .Titan
         T = (JD - 2415020) / 36525
         .nn0 = 0.001483716
         .i0 = 2.4923 - 0.00039 * T
         .st0 = 112.7836 + 0.8732 * T
         .e0 = 0.05589 - 0.000346 * T
         .p0 = 91.0891 + 1.9584 * T + 0.0008 * T * T
         .l0 = 266.5653 + 1223.5099 * T + 0.0003 * T * T
         t3 = (JD - 2411368) / 365.25
         Om_Om0 = fnmod(55.1687 + 0.00521 * t3 + 0.65 * Sin(p11 * (40.69 - 0.506 * t3)))
         l0_Pi0 = fnmod(53.3378 + 12.215515 * t3)
         Om_Om1 = fnmod(-0.11345 + 0.65 * Sin(p11 * (40.69 - 0.506 * t3)))
         i = fnmod(27.43883 + 0.30583 * Cos(p11 * (40.69 - 0.506 * t3)) - 0.00013 * t3)
         Xsi = atan2(Sin(.i0 * p11) * Sin(Om_Om0 * p11), Sin(i * p11) * Cos(.i0 * p11) - Cos(i * p11) * Sin(.i0 * p11) * Cos(Om_Om0 * p11)) / p11
         O_Om0 = atan2(Sin(i * p11) * Sin(Om_Om0 * p11), -Cos(i * p11) * Sin(.i0 * p11) + _
                      Sin(i * p11) * Cos(.i0 * p11) * Cos(Om_Om0 * p11)) / p11
         GroteG = atan2(Sin(.i0 * p11) * Sin(Om_Om0 * p11) / Sin(Xsi * p11), Cos(i * p11) * Cos(.i0 * p11) + _
                      Sin(i * p11) * Sin(.i0 * p11) * Cos(Om_Om0 * p11)) / p11
         .g0 = 103.199
         .g = .g0
         For k = 1 To 5
             .g = 108.2633 + 0.50956 * t3 - 0.65 * Sin(p11 * (40.69 - 0.506 * t3)) - Xsi + 0.36667 * (Sin(p11 * 2 * .g) - Sin(2 * p11 * .g0))
         Next
         .a = 1684.35
         .e = 0.0291 + 0.000186 * (Cos(2 * p11 * .g0) - Cos(2 * p11 * .g))
         .Pi = fnmod(276.12833333 + 0.5235 * t3) + 0.36666666667 * (Sin(2 * .g * p11) - Sin(2 * p11 * .g0))
         .n = 22.57701508
         .l = 260.40433333 + fnmod(.n * (JD - 2411368)) + 0.073166666667 * Sin(p11 * (40.69 - 0.506 * t3))
         .O = O_Om0 + .st0
         .de = 15 / 8 * .nn0 * .e * Cos(2 * p11 * (.l0 - .Pi))
         .dst = 3 / 8 * .nn0 * Sin(GroteG * p11) / Sin(i * p11) * Sin(p11 * (2 * .l0 - 2 * .O + Xsi))
         .di = .dst * Sin(i * p11)
         .dp = 15 / 8 * .nn0 * Sin(2 * p11 * (.l0 - .Pi))
         .d_w = atan2(Sin(i * p11) * Sin(Om_Om1 * p11), -Cos(i * p11) * Sin(i1) + Sin(i * p11) * Cos(i1) * Cos(Om_Om1 * p11))
         .fi1 = atan2(Sin(i1) * Sin(Om_Om1 * p11), Sin(i * p11) * Cos(i1) - Cos(i * p11) * Sin(i1) * Cos(Om_Om1 * p11))
         .gamma = atan2(1 / Sin(.fi1) * Sin(i1) * Sin(Om_Om1 * p11), Cos(i * p11) * Cos(i1) + Sin(i * p11) * Sin(i1) * Cos(Om_Om1 * p11))
         .dea = -3 * .nn0 * (.e0 * Sin(p11 * (l0_Pi0)) + 3 / 4 * .e0 * .e0 * Sin(2 * p11 * (l0_Pi0)) + 15 / 16 * .e * .e * Sin(2 * p11 * (l0_Pi0)) + 3 / 16 * _
                     Sin(GroteG * p11) * Sin(GroteG * p11) * Sin(2 * p11 * (.l0 - .O)))
         .dl = .dea - 2 * Sin(0.5 * i * p11) * Sin(0.5 * i1) * .dst
         .dma = .dea - .dp
         .sgdd = Sin(i * p11) * Cos(.fi1) * .dst - Sin(.fi1) * .di
         .dgamma = Sin(i * p11) * Sin(.fi1) * .dst + Cos(.fi1) * .di
         .e = .e + .de
         .middan = .l - Pi + .dma / p11 + ApproxDeltaT(JDToT(JD)) * .n / 86400
         .middlong = .l + 1 / p11 * (.d_w + w - p11 * (Om_Om0 + .st0) - .fi1) + .dl + ApproxDeltaT(JDToT(JD)) * .n / 86400
         .v = Kepler(.middan * p11, .e) / p11
         .u = .middlong + .v - .middan
         Call ElementToSigp_P(AltSaturnRingData.B, .u * p11 - AltSaturnRingData.u, Sig, p_P)
         .r = .a * (1 - .e * .e) / (1 + .e * Cos(.v * p11))
         s = .a / afst * Sin(Sig)
         .x = s * Sin(p_P + AltSaturnRingData.P)
         .y = s * Cos(p_P + AltSaturnRingData.P)
   End With
End If
End With
End Sub

Sub BasisGegevens(JD As Double, SaturnB As Double, Dist As Double)

T = JDToT(JD)

Obl = Obliquity(T)
Call NutationConst(T, NutLon, NutObl)
LAST = SiderealTime(T) + NutLon * Cos(Obl) - ObsLon

' Main Calculations }
Call PlanetPosHi(0, T, SAarde, True)
Call PlanetPosHi(6, T, SHelio, True)
Call HelioToGeo(SHelio, SAarde, SGeo)
Dist = SGeo.r
Call PlanetPosHi(6, T - Dist * LightTimeConst, SHelio, True)
Call HelioToGeo(SHelio, SAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)

Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA, Decl)
Call SaturnRing(T, SHelio, SGeo, Obl, NutLon, NutObl, SaturnRingData)
Call AltSaturnRing(T, SHelio, SGeo, Obl, NutLon, NutObl, AltSaturnRingData)

Call Aberration(T, Obl, FK5System, RA, Decl)
SaturnB = SaturnRingData.B
SemiDiameter = PlanetSemiDiameter(6, SGeo.r, PolarSemiDiameter)
Call CorrectSaturnSemiDiameter(SaturnB, PolarSemiDiameter)
End Sub

Sub SaturnusGeg(ByVal JD As Double, u, r, x, y, SaturnB As Double)

Dim i  As Long
Dim maangeg As tMaanGeg
Dim alf As Double, rx As Double
Dim Dist As Double
Call BasisGegevens(JD, SaturnB, Dist)
Dim ty As Double
Dim tx As Double

For maannr = 1 To 6
     'er moet een tijdcorrectie worden toegepast
     'omdat Saturnus vrij ver weg staat moet de tijd dat het licht hierover doet worden
     'meegenomen. In effect betekent dit dat wij de maantjes zien op een tijdstip waar ze
     'afh. van Dist eerst stonden
     Call sat_manen(JD - Dist * LightTimeConst * 36525, maannr, maangeg)

     If maannr < 6 Then
             x(maannr) = maangeg.Manen(maannr).x
             y(maannr) = maangeg.Manen(maannr).y
             u(maannr) = maangeg.Manen(maannr).u
             r(maannr) = maangeg.Manen(maannr).r
     Else
             x(maannr) = maangeg.Titan.x
             y(maannr) = maangeg.Titan.y
             u(maannr) = maangeg.Titan.u
             r(maannr) = maangeg.Titan.r
     End If
       
     rx = Sqr(x(maannr) * x(maannr) + y(maannr) * y(maannr))
     tx = x(maannr)
     ty = y(maannr)
     alf = atan2(ty, tx)
     alf = alf + AltSaturnRingData.P
     x(maannr) = Cos(alf) * rx
     y(maannr) = Sin(alf) * rx
Next
End Sub

