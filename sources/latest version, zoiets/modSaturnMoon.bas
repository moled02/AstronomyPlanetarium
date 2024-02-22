Attribute VB_Name = "modSaturnMoon"
Option Explicit
Const p11 = Pi / 180
Type tMain
    t1 As Double
    t2 As Double
    t3 As Double
    t4 As Double
    t5 As Double
    t6 As Double
    t7 As Double
    t8 As Double
    t9 As Double
    t10 As Double
    t11 As Double
    w0 As Double
    W1 As Double
    W2 As Double
    W3 As Double
    W4 As Double
    W5 As Double
    W6 As Double
    W7 As Double
    W8 As Double
    sl As Double
    cl As Double
    s2 As Double
    c2 As Double
    e1 As Double
End Type
Type tABC
    A As Double
    B As Double
    C As Double
End Type
Public Type tMaan
    l As Double
    r As Double
    g As Double
    Om As Double
    x As Double
    Y As Double
    Z As Double
    ABC(4) As tABC
End Type
Type tManen1_5
                     E0 As Double
                     n As Double
                     O  As Double
                     Gamma As Double
                     Pi1 As Double
                     E As Double
                     A As Double
                     dl As Double
                     u As Double
                     r_a As Double
                     r As Double
                     nD As Double
                     l1 As Double
                     M As Double
                     D As Double
                     l As Double
                     x As Double
                     Y As Double
                     g As Double
End Type
Type tmaanTitan
                     A As Double
                     E As Double
                     Pi As Double
                     n As Double
                     l As Double
                     O As Double
                     nn0 As Double
                     i0 As Double
                     st0 As Double
                     E0 As Double
                     p0 As Double
                     L0 As Double
                     r As Double
                     d_w As Double
                     fi1 As Double
                     Gamma As Double
                     dea As Double
                     dl As Double
                     dma As Double
                     sgdd As Double
                     dgamma As Double
                     g As Double
                     g0 As Double
                     DE As Double
                     dst As Double
                     di As Double
                     dp As Double
                     u As Double
                     middan As Double
                     middlong As Double
                     v As Double
                     x As Double
                     Y As Double
End Type
Type tMaanGeg
        manen(1 To 5) As tManen1_5
        Titan  As tmaanTitan
End Type
Type tSubroutine
    P1 As Double
    g As Double
    a1 As Double
    a2 As Double
    E As Double
    P As Double
    n As Double
    l1 As Double
    i As Double
    Om As Double
    A As Double
    lambda As Double
    gamme As Double
    W As Double
    r As Double
End Type
Private maingeg As tMain
Private maangeg   As tMaanGeg
Private ddate    As tDatum
Private Planet As Long, Earth   As Long
Private ObsLon As Double, ObsLat As Double, Height As Double
Private T As Double, DtofUT As Double, deltaT As Double, LAST   As Double
Private Obl As Double, NutLon As Double, NutObl As Double, TimeZone   As Double
Private SHelio As TSVECTOR, sAarde As TSVECTOR, SGeo As TSVECTOR
Private RA As Double, Decl As Double
Private SaturnDeltaU As Double, SaturnB As Double, Magnitude   As Double
Public SaturnRingData As TSATURNRINGDATA
Public AltSaturnRingData As TALTSATURNRINGDATA
Private JD As Double, JD0 As Double, jde As Double, JD_ZT As Double, JD_WT   As Double
Private i As Long, maannr As Long, H As Long
Public Semidiameter As Double
Public PolarSemiDiameter As Double
Public MaanBewPerDag As Variant
Private manen(9) As tMaan
Sub SaturnMoonInit()

MaanBewPerDag = Array(0, 381.9944421, 262.7319405, 190.69795, _
                        131.5349729, 79.690081, 22.57701508)
End Sub

Sub FillMainGeg(jde As Double)
    With maingeg
        .t1 = jde - 2411093#
        .t2 = .t1 / 365.25
        .t3 = (jde - 2433282.423) / 365.25 + 1950#
        .t4 = jde - 2411368#
        .t5 = .t4 / 365.25
        .t6 = jde - 2415020#
        .t7 = .t6 / 36525#
        .t8 = .t6 / 365.25
        .t9 = (jde - 2442000.5) / 365.25
        .t10 = jde - 2409786#
        .t11 = .t10 / 36525#
        .w0 = 5.095 * (.t3 - 1866.39)
        .W1 = 74.4 + 32.39 * .t2
        .W2 = 134.3 + 92.62 * .t2
        .W3 = 42# - 0.5118 * .t5
        .W4 = 276.59 + 0.5118 * .t5
        .W5 = 267.2635 + 1222.1136 * .t7
        .W6 = 175.4762 + 1221.5515 * .t7
        .W7 = 2.4891 + 0.002435 * .t7
        .W8 = 113.35 - 0.2597 * .t7
        .sl = sind(28.0818)
        .s2 = sind(168.8112)
        .cl = cosd(28.0817)
        .c2 = cosd(168.8112)
        .e1 = 0.05589 - 0.000346 * .t7
    End With
End Sub
Sub fillMaangeg()
    Call fillMimas_1(manen(1))
    Call fillEnceladus_2(manen(2))
    Call fillTehtys_3(manen(3))
    Call fillDione_4(manen(4))
    Call fillRhea_5(manen(5))
    Call fillTitan_6(manen(6))
    Call fillHyperion_7(manen(7))
    Call fillIapetus_8(manen(8))
End Sub
Sub fillMimas_1(Mimas As tMaan)
Dim l As Double
Dim P As Double
Dim M As Double
Dim C As Double

l = 127.64 + 381.994497 * maingeg.t1 - 43.57 * sind(maingeg.w0) - 0.72 * sind(3 * maingeg.w0) - 0.02144 * sind(5 * maingeg.w0)
P = 106.1 + 365.549 * maingeg.t2
M = l - P
C = 2.18287 * sind(M) + 0.025988 * sind(2 * M) + 0.00043 * sind(3 * M)
Mimas.l = l + C
Mimas.r = 3.06879 / (1 + 0.01905 * cosd(M + C))
Mimas.g = 1.563
Mimas.Om = 54.5 - 365.072 * maingeg.t2
End Sub
Sub fillEnceladus_2(Enceladus As tMaan)
Dim l As Double
Dim P As Double
Dim M As Double
Dim C As Double

l = 200.317 + 262.7319002 * maingeg.t1 + 0.25667 * sind(maingeg.W1) + 0.20883 * sind(maingeg.W2)
P = 309.107 + 123.44121 * maingeg.t2
M = l - P
C = 0.55577 * sind(M) + 0.00168 * sind(2 * M)
Enceladus.l = l + C
Enceladus.r = 3.94118 / (1 + 0.00485 * cosd(M + C))
Enceladus.g = 0.0262
Enceladus.Om = 348# - 151.95 * maingeg.t2
End Sub
Sub fillTehtys_3(Tethys As tMaan)
Tethys.l = 285.306 + 190.69791226 * maingeg.t1 + 2.063 * sind(maingeg.w0) + 0.03409 * sind(3 * maingeg.w0) + 0.001015 * sind(5 * maingeg.w0)
Tethys.r = 4.880998
Tethys.g = 1.0976
Tethys.Om = 111.33 - 72.2441 * maingeg.t2
End Sub
Sub fillDione_4(Dione As tMaan)
Dim l As Double
Dim P As Double
Dim M As Double
Dim C As Double

l = 254.712 + 131.53493193 * maingeg.t1 - 0.0215 * sind(maingeg.W1) - 0.01733 * sind(maingeg.W2)
P = 174.8 + 30.82 * maingeg.t2
M = l - P
C = 0.24717 * sind(M) + 0.00033 * sind(2 * M)
Dione.l = l + C
Dione.r = 6.24871 / (1 + 0.002157 * cosd(M + C))
Dione.g = 0.0139
Dione.Om = 232 - 30.27 * maingeg.t2
End Sub
Sub fillRhea_5(Rhea As tMaan)
Dim subr As tSubroutine
With subr
    .P1 = 342.7 + 10.057 * maingeg.t2
    .a1 = 0.000265 * sind(.P1) + 0.001 * sind(maingeg.W4)
    .a2 = 0.000265 * cosd(.P1) + 0.001 * cosd(maingeg.W4)
    .E = Sqr(.a1 * .a1 + .a2 * .a2)
    .P = atan2d(.a1, .a2)
    .n = 345 - 10.057 * maingeg.t2
    .l1 = 359.244 + 79.6900472 * maingeg.t1 + 0.086754 * sind(.n)
    .i = 28.0362 + 0.346898 * cosd(.n) + 0.0193 * cosd(maingeg.W3)
    .Om = 168.8034 + 0.736936 * sind(.n) + 0.041 * sind(maingeg.W3)
    .A = 8.725924
End With
Call subroutine(subr)
With Rhea
    .l = subr.lambda
    .Om = subr.W
    .r = subr.r
    .g = subr.gamme
End With
End Sub
Sub fillTitan_6(Titan As tMaan)
Dim l As Double
Dim i1 As Double
Dim o1 As Double
Dim Q As Double, e1 As Double, delta As Double
Dim b1 As Double, b2 As Double, u As Double, Om1 As Double
Dim g0 As Double, Phi As Double, s As Double, Om As Double
Dim Tmp As Integer

Dim subr As tSubroutine
With subr
    l = 261.1582 + 22.57697855 * maingeg.t4 + 0.074025 * sind(maingeg.W3)
    i1 = 27.45141 + 0.295999 * cosd(maingeg.W3)
    Om1 = 168.66925 + 0.628808 * sind(maingeg.W3)
    .a1 = sind(maingeg.W7) * sind(Om1 - maingeg.W8)
    .a2 = cosd(maingeg.W7) * sind(i1) - sind(maingeg.W7) * cosd(i1) * cosd(Om1 - maingeg.W8)
    g0 = 102.8623
    Phi = atan2d(.a1, .a2)
    s = Sqr(.a1 * .a1 + .a2 * .a2)
    .g = maingeg.W4 - Om1 - Phi
    For Tmp = 1 To 3
        Om = maingeg.W4 + 0.37515 * (sind(2 * .g) - sind(2 * g0))
        .g = Om - Om1 - Phi
    Next
    e1 = 0.029092 + 0.00019048 * (cosd(2 * .g) - cosd(2 * g0))
    Q = 2 * (maingeg.W5 - Om)
    b1 = sind(i1) * sind(Om1 - maingeg.W8)
    b2 = cosd(maingeg.W7) * sind(i1) * cosd(Om1 - maingeg.W8) - sind(maingeg.W7) * cosd(i1)
    delta = atan2d(b1, b2) + maingeg.W8
    .E = e1 + 0.002778797 * e1 * cosd(Q)
    .P = Om + 0.159215 * sind(Q)
    u = 2 * maingeg.W5 - 2 * delta + Phi
    H = 0.9375 * e1 * e1 * sind(Q) + 0.1875 * s * s * sind(2 * (maingeg.W5 - delta))
    .l1 = l - 0.254744 * (e1 * sind(maingeg.W6) + 0.75 * e1 * e1 * sind(2 * maingeg.W6) + H)
    .i = i1 + 0.031843 * s * cosd(u)
    .Om = Om1 + 0.031843 * s * sind(u) / sind(i1)
    .A = 20.216193
End With

Call subroutine(subr)
With Titan
    .l = subr.lambda
    .Om = subr.W
    .r = subr.r
    .g = subr.gamme
End With
End Sub
Sub fillHyperion_7(Hyperion As tMaan)
Dim nu As Double, xsi As Double, delta As Double, delta1 As Double
Dim asat As Double, bsat As Double, csat As Double, Om As Double
Dim Phi As Double, x As Double
Dim subr As tSubroutine
With subr
    nu = 92.39 + 0.5621071 * maingeg.t6
    xsi = 148.19 - 19.18 * maingeg.t8
    delta = 184.8 - 35.41 * maingeg.t9
    delta1 = delta - 7.5
    asat = 176 + 12.22 * maingeg.t8
    bsat = 8 + 24.44 * maingeg.t8
    csat = bsat + 5
    Om = 69.898 - 18.67088 * maingeg.t8
    Phi = 2 * (Om - maingeg.W5)
    x = 94.9 - 2.292 * maingeg.t8
    .A = 24.50601 - 0.08686 * cosd(nu) - 0.00166 * cosd(xsi + nu) + 0.00175 * cosd(xsi - nu)
    .E = 0.103458 - 0.004099 * cosd(nu) - 0.00167 * cosd(xsi + nu) _
         + 0.000235 * cosd(xsi - nu) + 0.02303 * cosd(xsi) - 0.00212 * cosd(2 * xsi) _
         + 0.000151 * cosd(3 * xsi) + 0.00013 * cosd(Phi)
    .P = Om + 0.15648 * sind(x) - 0.4457 * sind(nu) - 0.2657 * sind(xsi + nu) _
         - 0.3573 * sind(xsi - nu) - 12.872 * sind(xsi) + 1.668 * sind(2 * xsi) _
         - 0.2419 * sind(3 * xsi) - 0.07 * sind(Phi)
    .l1 = 177.047 + 16.91993829 * maingeg.t6 + 0.15648 * sind(x) + 9.142 * sind(nu) _
          + 0.007 * sind(2 * nu) - 0.014 * sind(3 * nu) + 0.2275 * sind(xsi + nu) _
          + 0.2112 * sind(xsi - nu) - 0.26 * sind(xsi) - 0.0098 * sind(2 * xsi) _
          - 0.013 * sind(asat) + 0.017 * sind(bsat) - 0.0303 * sind(Phi)
    .i = 27.3347 + 0.643486 * cosd(x) + 0.315 * cosd(maingeg.W3) + 0.018 * cosd(delta) - 0.018 * cosd(csat)
    .Om = 168.6812 + 1.40136 * cosd(x) + 0.68599 * sind(maingeg.W3) - 0.0392 * sind(csat) + 0.0366 * sind(delta1)
End With
Call subroutine(subr)
With Hyperion
    .l = subr.lambda
    .Om = subr.W
    .r = subr.r
    .g = subr.gamme
End With
End Sub
Sub fillIapetus_8(Japetus As tMaan)
Dim Om1 As Double, Phi As Double, delta As Double, Phi1 As Double, Phi2 As Double
Dim e1 As Double, w0 As Double, mu As Double, i1 As Double
Dim g As Double, g1 As Double, lS As Double, gs As Double, lt As Double, gt As Double
Dim u1 As Double, u2 As Double, u3 As Double, u4 As Double, u5 As Double
Dim W As Double, W1 As Double
Dim l As Double, lk As Double
Dim subr As tSubroutine
With subr
    l = 261.1582 + 22.57697855 * maingeg.t4
    W1 = 91.796 + 0.562 * maingeg.t7
    Phi = 4.367 - 0.195 * maingeg.t7
    delta = 146.819 - 3.198 * maingeg.t7
    Phi1 = 60.47 + 1.521 * maingeg.t7
    Phi2 = 205.055 - 2.091 * maingeg.t7
    e1 = 0.028298 + 0.001156 * maingeg.t11
    w0 = 352.91 + 11.71 * maingeg.t11
    mu = 76.3852 + 4.53795125 * maingeg.t10
    i1 = 18.4602 - 0.9518 * maingeg.t11 - 0.072 * maingeg.t11 * maingeg.t11 + 0.0054 * maingeg.t11 * maingeg.t11 * maingeg.t11
    Om1 = 143.198 - 3.919 * maingeg.t11 + 0.116 * maingeg.t11 * maingeg.t11 + 0.008 * maingeg.t11 * maingeg.t11 * maingeg.t11
    lk = mu - w0
    g = w0 - Om1 - Phi
    g1 = w0 - Om1 - Phi1
    lS = maingeg.W5 - W1
    gs = W1 - delta
    lt = l - maingeg.W4
    gt = maingeg.W4 - Phi2
    u1 = 2 * (lk + g - lS - gs)
    u2 = lk + g1 - lt - gt
    u3 = lk + 2 * (g - lS - gs)
    u4 = lt + gt - g1
    u5 = 2 * (lS + gs)
    .A = 58.935028 + 0.004638 * cosd(u1) + 0.058222 * cosd(u2)
    .E = e1 - 0.0014097 * cosd(g1 - gt) + 0.0003733 * cosd(u5 - 2 * g) _
            + 0.000118 * cosd(u3) + 0.0002408 * cosd(lk) _
            + 0.0002849 * cosd(lk + u2) + 0.000619 * cosd(u4)
    W = 0.08077 * sind(g1 - gt) + 0.02139 * sind(u5 - 2 * g) - 0.00676 * sind(u3) _
        + 0.0138 * sind(lk) + 0.01632 * sind(lk + u2) + 0.03547 * sind(u4)
    .P = w0 + W / e1
    .l1 = mu - 0.04299 * sind(u2) - 0.00789 * sind(u1) - 0.06312 * sind(lS) _
             - 0.00295 * sind(2 * lS) - 0.02231 * sind(u5) + 0.0065 * sind(u5 + Phi)
    .i = i1 + 0.04204 * cosd(u5 + Phi) + 0.00235 * cosd(lk + g1 + lt + gt + Phi1) + 0.0036 * cosd(u2 + Phi1)
    W1 = 0.04204 * sind(u5 + Phi) + 0.00235 * sind(lk + g1 + lt + gt + Phi1) + 0.00358 * sind(u2 + Phi1)
    .Om = Om1 + W1 / sind(i1)
    
End With
Call subroutine(subr)
With Japetus
    .l = subr.lambda
    .Om = subr.W
    .r = subr.r
    .g = subr.gamme
End With
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
Dim ch2 As Double, ch3 As Double, Om As Double, W As Double, i As Double, Phi1 As Double, afst As Double
Dim B As Double, i1 As Double
Dim k As Long
Dim s As Double, Sig As Double, p_P       As Double
Dim om_om1si1 As Double, i_i1 As Double, Xsi1 As Double, ea     As Double
Dim xsi As Double, O_Om0 As Double, Om_Om0 As Double, GroteG As Double, l0_Pi0 As Double, Om_Om1 As Double

With maangeg
   T = JDToT(JD)
   i1 = (28.075216 - T * (0.012998 - T * 0.000004)) * DToR
   Om = (169.50847 + T * (1.394681 + T * 0.000412)) * DToR
   W = AltSaturnRingData.W
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
       .manen(1).E0 = 127.09166666
       .manen(1).n = 381.994442
       .manen(1).O = fnmod(56.1 - 365.23 * t2)
       .manen(1).Gamma = 1.5166666667
       .manen(1).Pi1 = fnmod(105 + 365.6 * t2)
       .manen(1).E = 0.0201
       .manen(1).A = 255.89
       .manen(1).dl = -44.39 * Sin(ch1) - 0.764 * Sin(3 * ch1)
    Case 2
       .manen(2).E0 = 199.43
       .manen(2).n = 262.7319405
       .manen(2).O = fnmod(52 - 152.7 * t2)
       .manen(2).Gamma = 0.023333333333
       .manen(2).Pi1 = fnmod(308.38 + 123.43 * t2)
       .manen(2).E = 0.00444
       .manen(2).A = 328.29
       .manen(2).dl = 0.23983333333 * Sin(ch2) + 0.23433333333 * Sin(ch3)
    Case 3
       .manen(3).E0 = 284.47166667
       .manen(3).n = 190.69795
       .manen(3).O = fnmod(110.39 - 72.25 * t2)
       .manen(3).Gamma = 1.0926666667
       .manen(3).Pi1 = 0
       .manen(3).E = 0
       .manen(3).A = 406.4
       .manen(3).dl = 2.065 * Sin(ch1) + 0.036 * Sin(3 * ch1)
    Case 4
       .manen(4).E0 = 253.86666667
       .manen(4).n = 131.5349729
       .manen(4).O = fnmod(201 - 31 * t2)
       .manen(4).Gamma = 0.023333333333
       .manen(4).Pi1 = fnmod(173.4 + 30.75 * t2)
       .manen(4).E = 0.00221
       .manen(4).A = 520.51
       .manen(4).dl = -0.0155 * Sin(ch2) - 0.015166666667 * Sin(ch3)
  End Select
  If maannr < 5 Then
  ' MIMAS ENCELADUS TETHYS DIONE }
    With .manen(maannr)
         .l1 = fnmod(.E0) + fnmod(.n * (JD - 2411093) + .dl)
         .M = .l1 - .Pi1
         .D = .O + (W - Om) / p11
         .l = .l1 + (W - Om) / p11
         .u = .l + 1 / p11 * .E * (2 * Sin((.M * p11)) + 5 / 4 * .E * Sin(2 * .M * p11))

         .r_a = 1 + .E * (0.5 * .E - Cos(.M * p11) - 0.5 * .E * Cos(2 * .M * p11))
         .r = .r_a + .A
         Call ElementToSigp_P(AltSaturnRingData.B, .u * p11 - AltSaturnRingData.u, Sig, p_P)
         s = .A / afst * Sin(Sig)
         .x = s * Sin(p_P + AltSaturnRingData.P)
         .Y = s * Cos(p_P + AltSaturnRingData.P)
    End With
  End If
  'RHEA}
  If maannr = 5 Then
     With .manen(maannr)
         .E0 = 358.395
         .n = 79.6900881
         om_om1si1 = 1 / 60 * (20.49 * Sin(p11 * fnmod(344.09 - 10.2 * t2)) - 0.38 + 1 * Sin(p11 * fnmod(48.5 - 0.5 * t2)))
         i_i1 = 1 / 60 * (20.49 * Cos(p11 * fnmod(344.09 - 10.2 * t2)) - 2.79 + 1 * Cos(p11 * fnmod(48.5 - 0.5 * t2)))
         .Pi1 = 276.25 + 0.53 * t2 + 17.64 * Sin(p11 * fnmod(9.5 * (t1 + 20.41)))
         .E = 0.00098 + 0.0003 * Cos(p11 * fnmod(9.5 * (t1 + 20.41)))
         .A = 726.89
         Xsi1 = atan2(om_om1si1, i_i1) / p11
         .D = Xsi1 + W / p11 + om_om1si1 / tan(i1) * (1 - 0.5 * Sin(Xsi1 * p11) * Sin(Xsi1 * p11))
         .g = om_om1si1 / Sin(Xsi1 * p11)
         .l1 = fnmod(.E0) + fnmod(.n * (JD - 2411093))
         .M = .l1 - .Pi1
         .l = .l1 + (W - Om) / p11
         .u = .l + 1 / p11 * .E * (2 * Sin((.M * p11)) + 5 / 4 * .E * Sin(2 * .M * p11))
         .r_a = 1 + .E * (0.5 * .E - Cos(.M * p11) - 0.5 * .E * Cos(2 * .M * p11))
         .r = .r_a + .A
         Call ElementToSigp_P(AltSaturnRingData.B, .u * p11 - AltSaturnRingData.u, Sig, p_P)
         s = .A / afst * Sin(Sig)
         .x = s * Sin(p_P + AltSaturnRingData.P)
         .Y = s * Cos(p_P + AltSaturnRingData.P)
     End With
  End If
' TITAN}
  If maannr = 6 Then
     With .Titan
         T = (JD - 2415020) / 36525
         .nn0 = 0.001483716
         .i0 = 2.4923 - 0.00039 * T
         .st0 = 112.7836 + 0.8732 * T
         .E0 = 0.05589 - 0.000346 * T
         .p0 = 91.0891 + 1.9584 * T + 0.0008 * T * T
         .L0 = 266.5653 + 1223.5099 * T + 0.0003 * T * T
         t3 = (JD - 2411368) / 365.25
         Om_Om0 = fnmod(55.1687 + 0.00521 * t3 + 0.65 * Sin(p11 * (40.69 - 0.506 * t3)))
         l0_Pi0 = fnmod(53.3378 + 12.215515 * t3)
         Om_Om1 = fnmod(-0.11345 + 0.65 * Sin(p11 * (40.69 - 0.506 * t3)))
         i = fnmod(27.43883 + 0.30583 * Cos(p11 * (40.69 - 0.506 * t3)) - 0.00013 * t3)
         xsi = atan2(Sin(.i0 * p11) * Sin(Om_Om0 * p11), Sin(i * p11) * Cos(.i0 * p11) - Cos(i * p11) * Sin(.i0 * p11) * Cos(Om_Om0 * p11)) / p11
         O_Om0 = atan2(Sin(i * p11) * Sin(Om_Om0 * p11), -Cos(i * p11) * Sin(.i0 * p11) + _
                      Sin(i * p11) * Cos(.i0 * p11) * Cos(Om_Om0 * p11)) / p11
         GroteG = atan2(Sin(.i0 * p11) * Sin(Om_Om0 * p11) / Sin(xsi * p11), Cos(i * p11) * Cos(.i0 * p11) + _
                      Sin(i * p11) * Sin(.i0 * p11) * Cos(Om_Om0 * p11)) / p11
         .g0 = 103.199
         .g = .g0
         For k = 1 To 5
             .g = 108.2633 + 0.50956 * t3 - 0.65 * Sin(p11 * (40.69 - 0.506 * t3)) - xsi + 0.36667 * (Sin(p11 * 2 * .g) - Sin(2 * p11 * .g0))
         Next
         .A = 1684.35
         .E = 0.0291 + 0.000186 * (Cos(2 * p11 * .g0) - Cos(2 * p11 * .g))
         .Pi = fnmod(276.12833333 + 0.5235 * t3) + 0.36666666667 * (Sin(2 * .g * p11) - Sin(2 * p11 * .g0))
         .n = 22.57701508
         .l = 260.40433333 + fnmod(.n * (JD - 2411368)) + 0.073166666667 * Sin(p11 * (40.69 - 0.506 * t3))
         .O = O_Om0 + .st0
         .DE = 15 / 8 * .nn0 * .E * Cos(2 * p11 * (.L0 - .Pi))
         .dst = 3 / 8 * .nn0 * Sin(GroteG * p11) / Sin(i * p11) * Sin(p11 * (2 * .L0 - 2 * .O + xsi))
         .di = .dst * Sin(i * p11)
         .dp = 15 / 8 * .nn0 * Sin(2 * p11 * (.L0 - .Pi))
         .d_w = atan2(Sin(i * p11) * Sin(Om_Om1 * p11), -Cos(i * p11) * Sin(i1) + Sin(i * p11) * Cos(i1) * Cos(Om_Om1 * p11))
         .fi1 = atan2(Sin(i1) * Sin(Om_Om1 * p11), Sin(i * p11) * Cos(i1) - Cos(i * p11) * Sin(i1) * Cos(Om_Om1 * p11))
         .Gamma = atan2(1 / Sin(.fi1) * Sin(i1) * Sin(Om_Om1 * p11), Cos(i * p11) * Cos(i1) + Sin(i * p11) * Sin(i1) * Cos(Om_Om1 * p11))
         .dea = -3 * .nn0 * (.E0 * Sin(p11 * (l0_Pi0)) + 3 / 4 * .E0 * .E0 * Sin(2 * p11 * (l0_Pi0)) + 15 / 16 * .E * .E * Sin(2 * p11 * (l0_Pi0)) + 3 / 16 * _
                     Sin(GroteG * p11) * Sin(GroteG * p11) * Sin(2 * p11 * (.L0 - .O)))
         .dl = .dea - 2 * Sin(0.5 * i * p11) * Sin(0.5 * i1) * .dst
         .dma = .dea - .dp
         .sgdd = Sin(i * p11) * Cos(.fi1) * .dst - Sin(.fi1) * .di
         .dgamma = Sin(i * p11) * Sin(.fi1) * .dst + Cos(.fi1) * .di
         .E = .E + .DE
         .middan = .l - Pi + .dma / p11 + ApproxDeltaT(JDToT(JD)) * .n / 86400
         .middlong = .l + 1 / p11 * (.d_w + W - p11 * (Om_Om0 + .st0) - .fi1) + .dl + ApproxDeltaT(JDToT(JD)) * .n / 86400
         .v = Kepler(.middan * p11, .E) / p11
         .u = .middlong + .v - .middan
         Call ElementToSigp_P(AltSaturnRingData.B, .u * p11 - AltSaturnRingData.u, Sig, p_P)
         .r = .A * (1 - .E * .E) / (1 + .E * Cos(.v * p11))
         s = .A / afst * Sin(Sig)
         .x = s * Sin(p_P + AltSaturnRingData.P)
         .Y = s * Cos(p_P + AltSaturnRingData.P)
   End With
End If
End With
End Sub

Sub subroutine(ByRef subgeg As tSubroutine)
Dim M As Double
Dim C As Double
Dim H As Double
Dim u As Double
Dim Phi As Double

With subgeg
    M = .l1 - .P
    C = (2 * .E + .E * .E * .E * (-0.25 + 0.0520833333 * .E * .E)) * sind(M) _
       + .E * .E * (1.25 - 0.458333333 * .E * .E) * sind(2 * M) _
       + .E * .E * .E * (1.083333333 - 0.671875 * .E * .E) * sind(3 * M) _
       + .E * .E * .E * .E * (1.072917 * sind(4 * M) + 1.142708 * .E * sind(5 * M))
    C = C / p11
    .r = .A * (1 - .E * .E) / (1 + .E * cosd(M + C))
    .g = .Om - 168.8112
    .a1 = sind(.i) * sind(.g)
    .a2 = maingeg.cl * sind(.i) * cosd(.g) - maingeg.sl * cosd(.i)
    .gamme = asind(Sqr(.a1 * .a1 + .a2 * .a2))
    u = atan2d(.a1, .a2)
    .W = 168.8112 + u
    H = maingeg.cl * sind(.i) - maingeg.sl * cosd(.i) * cosd(.g)
    Phi = atan2d(maingeg.sl * sind(.g), H)
    .lambda = .l1 + C + u - .g - Phi
End With
End Sub

Sub CalculateXYZ()

Dim i As Long
Dim u As Double, W As Double, D As Double, perspectiveFactor As Double

For i = 1 To 8
    With manen(i)
        u = .l - .Om
        W = .Om - 168.8112
        .x = .r * (cosd(u) * cosd(W) - sind(u) * cosd(.g) * sind(W))
        .Y = .r * (sind(u) * cosd(W) * cosd(.g) + cosd(u) * sind(W))
        .Z = .r * sind(u) * sind(.g)
    End With
Next
With manen(9)
    .x = 0
    .Y = 0
    .Z = 1
End With
For i = 1 To 9
    With manen(i)
        .ABC(1).A = .x
        .ABC(1).B = maingeg.cl * .Y - maingeg.sl * .Z
        .ABC(1).C = maingeg.sl * .Y + maingeg.cl * .Z
        .ABC(2).A = maingeg.c2 * .ABC(1).A - maingeg.s2 * .ABC(1).B
        .ABC(2).B = maingeg.s2 * .ABC(1).A + maingeg.c2 * .ABC(1).B
        .ABC(2).C = .ABC(1).C
        .ABC(3).A = .ABC(2).A * Sin(SGeo.l) - .ABC(2).B * Cos(SGeo.l)
        .ABC(3).B = .ABC(2).A * Cos(SGeo.l) + .ABC(2).B * Sin(SGeo.l)
        .ABC(3).C = .ABC(2).C
        .ABC(4).A = .ABC(3).A
        .ABC(4).B = .ABC(3).B * Cos(SGeo.B) + .ABC(3).C * Sin(SGeo.B)
        .ABC(4).C = .ABC(3).C * Cos(SGeo.B) - .ABC(3).B * Sin(SGeo.B)
    End With
Next
With manen(9)
    D = atan2(.ABC(4).A, .ABC(4).C)
End With
Dim diffLightFact As Variant
diffLightFact = Array(0, 20947, 23715, 26382, 29876, 35313, 53800, 59222, 91820)
For i = 1 To 8
    With manen(i)
        .x = .ABC(4).A * Cos(D) - .ABC(4).C * Sin(D)
        .Y = .ABC(4).A * Sin(D) + .ABC(4).C * Cos(D)
        .Z = .ABC(4).B
        
        .x = .x + Abs(.Z) / diffLightFact(i) * ssqr(1 - (.x / .r) * (.x / .r))
        
        perspectiveFactor = SGeo.r / (SGeo.r + .Z / 2475)
        .x = .x * perspectiveFactor
        .Y = .Y * perspectiveFactor
        .Z = .Z
    End With
Next
End Sub
Function ssqr(x As Double) As Double
    If x < 0 Then x = 0
    ssqr = Sqr(x)
End Function
Sub BasisGegevens(JD As Double, SaturnB As Double, Dist As Double)

T = JDToT(JD)

Obl = Obliquity(T)
Call NutationConst(T, NutLon, NutObl)
LAST = SiderealTime(T) + NutLon * Cos(Obl) - ObsLon

' Main Calculations }
Call PlanetPosHi(0, T, sAarde, False)
Call PlanetPosHi(6, T, SHelio, False)
Call HelioToGeo(SHelio, sAarde, SGeo)
Dist = SGeo.r
Call PlanetPosHi(6, T - Dist * LightTimeConst, SHelio, True)
Call HelioToGeo(SHelio, sAarde, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)

Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RA, Decl)
Call SaturnRing(T, SHelio, SGeo, Obl, NutLon, NutObl, SaturnRingData)
Call AltSaturnRing(T, SHelio, SGeo, Obl, NutLon, NutObl, AltSaturnRingData)

Call Aberration(T, Obl, FK5System, RA, Decl)
SaturnB = SaturnRingData.B
Semidiameter = PlanetSemiDiameter(6, SGeo.r, PolarSemiDiameter)
Call CorrectSaturnSemiDiameter(SaturnB, PolarSemiDiameter)
End Sub

Sub SaturnusGeg(ByVal JD As Double, u, r, x, Y, SaturnB As Double)

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
             x(maannr) = maangeg.manen(maannr).x
             Y(maannr) = maangeg.manen(maannr).Y
             u(maannr) = maangeg.manen(maannr).u
             r(maannr) = maangeg.manen(maannr).r
     Else
             x(maannr) = maangeg.Titan.x
             Y(maannr) = maangeg.Titan.Y
             u(maannr) = maangeg.Titan.u
             r(maannr) = maangeg.Titan.r
     End If
       
     rx = Sqr(x(maannr) * x(maannr) + Y(maannr) * Y(maannr))
     tx = x(maannr)
     ty = Y(maannr)
     alf = atan2(ty, tx)
     alf = alf + AltSaturnRingData.P
     x(maannr) = Cos(alf) * rx
     Y(maannr) = Sin(alf) * rx
Next
End Sub





Sub CalcSaturnMoons(jde As Double, ByRef satmanen() As tMaan)
Dim Dist As Double
Call BasisGegevens(jde, SaturnB, Dist)
jde = jde - Dist * LightTimeConst * 36525
Dim T As Double
T = JDToT(jde)
Call PrecessEcliptic(T, TB1950, SGeo.l, SGeo.B)
Call FillMainGeg(jde)
Call fillMaangeg
CalculateXYZ
Dim i As Long
For i = 1 To 9
    satmanen(i) = manen(i)
Next
End Sub
