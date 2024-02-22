Attribute VB_Name = "modJsatsHi"
'(*****************************************************************************)
'(* Module: JSATSHI.PAS                                                       *)
'(* Version 2.0                                                               *)
'(* Last modified: October 1, 1992                                            *)
'(*****************************************************************************)
Public Const DUMMY_SATELLITE = 5

Private lCoeff As Variant
Private pCoeff As Variant
Private omCoeff As Variant
'Private lCoeff(4, 1) As Double
'Private pCoeff(4, 1) As Double
'Private omCoeff(4, 1) As Double

Private l(4) As Double, dl(4) As Double, P(4) As Double, Om(4) As Double
Private Gamma As Double, PhiL As Double, Psi As Double, G As Double, Gd As Double, BigPi   As Double
Private Tcur   As Double

Sub JSats(T As Double)

Dim i As Long
Dim smallt As Double

lCoeff = Array(Array(0, 0), Array(106.07947, 203.488955432), Array(175.72938, 101.37472455), _
         Array(120.55434, 50.31760911), Array(84.44868, 21.571071314))
pCoeff = Array(Array(0, 0), Array(58.3329, 0.16103936), Array(132.8959, 0.04647985), _
         Array(187.2887, 0.0071274), Array(335.3418, 0.00183998))
omCoeff = Array(Array(0, 0), Array(311.0793, -0.1327943), Array(100.5099, -0.03263047), _
         Array(119.1688, -0.00717704), Array(322.5729, -0.00175934))

'If (Tcur <> T) Then
    smallt = (T * 36525# + 8544.5)
    Gamma = 0.33033 * Sin((163.679 + 0.0010512 * smallt) * DToR)
    Gamma = Gamma + 0.03439 * Sin((34.486 - 0.0161731 * smallt) * DToR)
    PhiL = (191.8132 + 0.17390023 * smallt) * DToR
    Psi = (316.5182 - 0.00000208 * smallt) * DToR
    G = (30.23756 + 0.0830925701 * smallt + Gamma) * DToR
    Gd = (31.97853 + 0.0334597339 * smallt) * DToR
    BigPi = 13.469942 * DToR
    For i = 1 To 4
      l(i) = modpi2((lCoeff(i)(0) + smallt * lCoeff(i)(1)) * DToR)
      P(i) = (pCoeff(i)(0) + smallt * pCoeff(i)(1)) * DToR
      Om(i) = (omCoeff(i)(0) + smallt * omCoeff(i)(1)) * DToR
    Next
    Tcur = T
'End If
End Sub

Sub Io(T As Double, ByRef v As TSVECTOR)
' { Io }
  Call JSats(T)
  dl(1) = dl1
  l(1) = l(1) + dl(1)
  v.l = l(1)
  v.B = b1
  l(1) = l(1) - dl(1)
  v.r = r1
End Sub

Private Function dl1() As Double

  Dim dl As Double
    dl = _
      0.47259 * Sin(2 * (l(1) - l(2))) _
      - 0.0348 * Sin(P(3) - P(4)) _
      - 0.01756 * Sin(P(1) + P(3) - 2 * BigPi - 2 * G) _
      + 0.0108 * Sin(l(2) - 2 * l(3) + P(3)) _
      + 0.00757 * Sin(PhiL)
    dl = dl + _
      0.00663 * Sin(l(2) - 2 * l(3) + P(4)) _
      + 0.00453 * Sin(l(1) - P(3)) _
      + 0.00453 * Sin(l(2) - 2 * l(3) + P(2)) _
      - 0.00354 * Sin(l(1) - l(2)) _
      - 0.00317 * Sin(2 * Psi - 2 * BigPi)
    dl = dl + _
      -0.00269 * Sin(l(2) - 2 * l(3) + P(1)) _
      + 0.00263 * Sin(l(1) - P(4)) _
      + 0.00186 * Sin(l(1) - P(1)) _
      - 0.00186 * Sin(G) _
      + 0.00167 * Sin(P(2) - P(3))
    dl = dl + _
      0.00158 * Sin(4 * (l(1) - l(2))) _
      - 0.00155 * Sin(l(1) - l(3)) _
      - 0.00142 * Sin(Psi + Om(3) - 2 * BigPi - 2 * G) _
      - 0.00115 * Sin(2 * (l(1) - 2 * l(2) + Om(2))) _
      + 0.00089 * Sin(P(2) - P(4))
    dl = dl + _
      0.00084 * Sin(Om(2) - Om(3)) _
      + 0.00084 * Sin(l(1) + P(3) - 2 * BigPi - 2 * G) _
      + 0.00053 * Sin(Psi - Om(2))
    dl1 = dl * DToR
End Function

Private Function b1() As Double

Dim tanb As Double
    tanb = _
      0.0006502 * Sin(l(1) - Om(1)) _
      + 0.0001835 * Sin(l(1) - Om(2)) _
      + 0.0000329 * Sin(l(1) - Om(3)) _
      - 0.0000311 * Sin(l(1) - Psi) _
      + 0.0000093 * Sin(l(1) - Om(4)) _
      + 0.0000075 * Sin(3 * l(1) - 4 * l(2) - 1.9927 * dl(1) + Om(2)) _
      + 0.0000046 * Sin(l(1) + Psi - 2 * BigPi - 2 * G)
    b1 = Atn(tanb)
End Function

Private Function r1() As Double

Dim dr As Double
    dr = _
      -0.0041339 * Cos(2 * (l(1) - l(2))) _
      - 0.0000395 * Cos(l(1) - P(3)) _
      - 0.0000214 * Cos(l(1) - P(4)) _
      + 0.000017 * Cos(l(1) - l(2))
    dr = dr + _
      -0.000013 * Cos(4 * (l(1) - l(2))) _
      + 0.0000106 * Cos(l(1) - l(3)) _
      - 0.0000162 * Cos(l(1) - P(1)) _
      - 0.0000063 * Cos(l(1) + P(3) - 2 * BigPi - 2 * G)
    r1 = 5.9073 * (1 + dr)
End Function

Sub Europa(T As Double, ByRef v As TSVECTOR)
' { Europa }
  Call JSats(T)
  dl(2) = dl2
  l(2) = l(2) + dl(2)
  v.l = l(2)
  v.B = b2
  l(2) = l(2) - dl(2)
  v.r = r2
End Sub

Private Function dl2() As Double

Dim dl As Double
    dl = _
      1.06476 * Sin(2 * (l(2) - l(3))) _
      + 0.04253 * Sin(l(1) - 2 * l(2) + P(3)) _
      + 0.03579 * Sin(l(2) - P(3)) _
      + 0.02383 * Sin(l(1) - 2 * l(2) + P(4)) _
      + 0.01977 * Sin(l(2) - P(4)) _
      - 0.01843 * Sin(PhiL)
    dl = dl + _
      0.01299 * Sin(P(3) - P(4)) _
      - 0.01142 * Sin(l(2) - l(3)) _
      - 0.01058 * Sin(G) _
      + 0.01078 * Sin(l(2) - P(2)) _
      + 0.00327 * Sin(Psi - 2 * G + Om(3) - 2 * BigPi) _
      + 0.0087 * Sin(l(2) - 2 * l(3) + P(2))
    dl = dl + _
      -0.00775 * Sin(Psi - BigPi) _
      + 0.00524 * Sin(2 * (l(1) - l(2))) _
      - 0.0046 * Sin(l(1) - l(3)) _
      + 0.0045 * Sin(l(2) - 2 * l(3) + P(1)) _
      - 0.00296 * Sin(P(1) + P(3) - 2 * BigPi - 2 * G) _
      - 0.00151 * Sin(2 * G)
    dl = dl + _
      0.00146 * Sin(Psi - Om(3)) _
      + 0.00125 * Sin(Psi - Om(4)) _
      - 0.00117 * Sin(l(1) - 2 * l(3) + P(3)) _
      - 0.00095 * Sin(2 * (l(2) - Om(2))) _
      + 0.00086 * Sin(2 * (l(1) - 2 * l(2) + Om(2))) _
      - 0.00086 * Sin(5 * Gd - 2 * G + 52.225 * DToR)
    dl = dl + _
      -0.00078 * Sin(l(2) - l(4)) _
      - 0.00064 * Sin(l(1) - 2 * l(3) + P(4)) _
      - 0.00063 * Sin(3 * l(3) - 7 * l(4) + 4 * P(4)) _
      + 0.00061 * Sin(P(1) - P(4)) _
      + 0.00058 * Sin(2 * (Psi - BigPi - G)) _
      + 0.00058 * Sin(Om(3) - Om(4))
    dl = dl + _
      0.00056 * Sin(2 * (l(2) - l(4))) _
      + 0.00055 * Sin(2 * (l(1) - l(3))) _
      + 0.00052 * Sin(3 * l(3) - 7 * l(4) + P(3) + 3 * P(4)) _
      - 0.00043 * Sin(l(1) - P(3)) _
      + 0.00042 * Sin(P(3) - P(2)) _
      + 0.00041 * Sin(5 * (l(2) - l(3)))
    dl = dl + _
      0.00041 * Sin(P(4) - BigPi) _
      + 0.00032 * Sin(Om(2) - Om(3)) _
      + 0.00032 * Sin(2 * (l(3) - G - BigPi)) _
      + 0.00029 * Sin(P(1) - P(3)) _
      + 0.00038 * Sin(l(2) - P(1))
    dl2 = dl * DToR
End Function

Private Function b2() As Double

Dim tanb As Double
    tanb = _
      0.0081275 * Sin(l(2) - Om(2)) _
      + 0.0004512 * Sin(l(2) - Om(3)) _
      - 0.0003286 * Sin(l(2) - Psi) _
      + 0.0001164 * Sin(l(2) - Om(4)) _
      + 0.0000273 * Sin(l(1) - 2 * l(3) + 1.0146 * dl(2) + Om(2))
    tanb = tanb + _
      0.0000143 * Sin(l(2) + Psi - 2 * BigPi - 2 * G) _
      - 0.0000143 * Sin(l(2) - Om(1)) _
      + 0.0000035 * Sin(l(2) - Psi + G) _
      - 0.0000028 * Sin(l(1) - 2 * l(3) + 1.0146 * dl(2) + Om(3))
    b2 = Atn(tanb)
End Function

Private Function r2() As Double

Dim dr As Double
    dr = _
      0.0093847 * Cos(l(1) - l(2)) _
      - 0.0003114 * Cos(l(2) - P(3)) _
      - 0.0001738 * Cos(l(2) - P(4)) _
      - 0.0000941 * Cos(l(2) - P(2)) _
      + 0.0000553 * Cos(l(2) - l(3)) _
      + 0.0000523 * Cos(l(1) - l(3))
    dr = dr + _
      -0.000029 * Cos(2 * (l(1) - l(2))) _
      + 0.0000166 * Cos(2 * (l(2) - Om(2))) _
      + 0.0000107 * Cos(l(1) - 2 * l(3) + P(3)) _
      - 0.0000102 * Cos(l(2) - P(1)) _
      - 0.0000091 * Cos(2 * (l(1) - l(3)))
    r2 = 9.39912 * (1 + dr)
End Function


Sub Ganymede(T As Double, ByRef v As TSVECTOR)
'{ Ganymede }
  Call JSats(T)
  dl(3) = dl3
  l(3) = l(3) + dl(3)
  v.l = l(3)
  v.B = b3
  l(3) = l(3) - dl(3)
  v.r = r3
End Sub

Private Function dl3() As Double

Dim dl As Double
    dl = _
      0.16477 * Sin(l(3) - P(3)) _
      + 0.09062 * Sin(l(3) - P(4)) _
      - 0.06907 * Sin(l(2) - l(3)) _
      + 0.03786 * Sin(P(3) - P(4)) _
      + 0.01844 * Sin(2 * (l(3) - l(4))) _
      - 0.0134 * Sin(G)
    dl = dl + _
      -0.0067 * Sin(2 * (Psi - BigPi)) _
      + 0.00703 * Sin(l(2) - 2 * l(3) + P(3)) _
      - 0.0054 * Sin(l(3) - l(4)) _
      - 0.00409 * Sin(l(2) - 2 * l(3) + P(2)) _
      + 0.00379 * Sin(l(2) - 2 * l(3) + P(4)) _
      + 0.00481 * Sin(P(1) + P(3) - 2 * BigPi - 2 * G)
    dl = dl + _
      0.00235 * Sin(Psi - Om(3)) _
      + 0.00198 * Sin(Psi - Om(4)) _
      + 0.0018 * Sin(PhiL) _
      + 0.00124 * Sin(l(1) - l(3)) _
      - 0.00119 * Sin(5 * Gd - 2 * G + 52.225 * DToR) _
      + 0.00109 * Sin(l(1) - l(2))
    dl = dl + _
      0.00129 * Sin(3 * (l(3) - l(4))) _
      - 0.00099 * Sin(3 * l(3) - 7 * l(4) + 4 * P(4)) _
      - 0.00029 * Sin(Om(3) + Psi - 2 * BigPi - 2 * G) _
      + 0.00091 * Sin(Om(3) - Om(4)) _
      + 0.00081 * Sin(3 * l(3) - 7 * l(4) + P(3) + 3 * P(4)) _
      - 0.00076 * Sin(2 * l(2) - 3 * l(3) + P(3))
    dl = dl + _
      0.00069 * Sin(P(4) - BigPi) _
      - 0.00058 * Sin(2 * l(3) - 3 * l(4) + P(4)) _
      + 0.00057 * Sin(l(3) + P(3) - 2 * BigPi - 2 * G) _
      - 0.00057 * Sin(l(3) - 2 * l(4) + P(4)) _
      - 0.00052 * Sin(P(2) - P(3)) _
      - 0.00052 * Sin(l(2) - 2 * l(3) + P(1))
    dl = dl + _
      0.00048 * Sin(l(3) - 2 * l(4) + P(3)) _
      - 0.00045 * Sin(2 * l(2) - 3 * l(3) + P(4)) _
      - 0.00041 * Sin(P(2) - P(4)) _
      - 0.00038 * Sin(2 * G) _
      - 0.00033 * Sin(P(3) - P(4) + Om(3) - Om(4)) _
      - 0.00032 * Sin(3 * l(3) - 7 * l(4) + 2 * P(3) + 2 * P(4))
    dl = dl + _
      0.0003 * Sin(4 * (l(3) - l(4))) _
      + 0.00029 * Sin(l(3) + P(4) - 2 * BigPi - 2 * G) _
      + 0.00026 * Sin(l(3) - BigPi - G) _
      + 0.00024 * Sin(l(2) - 3 * l(3) + 2 * l(4)) _
      + 0.00021 * Sin(2 * (l(3) - BigPi - G)) _
      - 0.00021 * Sin(l(3) - P(2)) _
      + 0.00017 * Sin(2 * (l(3) - P(3)))
  dl3 = dl * DToR
End Function

Private Function b3() As Double

Dim tanb As Double
    tanb = _
      0.0032364 * Sin(l(3) - Om(3)) _
      - 0.0016911 * Sin(l(3) - Psi) _
      + 0.0006849 * Sin(l(3) - Om(4)) _
      - 0.0002806 * Sin(l(3) - Om(2)) _
      + 0.0000321 * Sin(l(3) + Psi - 2 * BigPi - 2 * G) _
      + 0.0000051 * Sin(l(3) - Psi + G)
    tanb = tanb + _
      -0.0000045 * Sin(l(3) - Psi - G) _
      - 0.0000045 * Sin(l(3) + Psi - 2 * BigPi) _
      + 0.0000037 * Sin(l(3) + Psi - 2 * BigPi - 3 * G) _
      + 0.000003 * Sin(2 * l(2) - 3 * l(3) + 4.03 * dl(3) + Om(2)) _
      - 0.0000021 * Sin(2 * l(2) - 3 * l(3) + 4.03 * dl(3) + Om(3))
    b3 = Atn(tanb)
End Function

Private Function r3() As Double
Dim dr As Double
    dr = _
      -0.0014377 * Cos(l(3) - P(3)) _
      - 0.0007904 * Cos(l(3) - P(4)) _
      + 0.0006342 * Cos(l(2) - l(3)) _
      - 0.0001758 * Cos(2 * (l(3) - l(4))) _
      + 0.0000294 * Cos(l(3) - l(4))
    dr = dr + _
      -0.0000153 * Cos(l(1) - l(2)) _
      + 0.0000155 * Cos(l(1) - l(3)) _
      - 0.0000156 * Cos(3 * (l(3) - l(4))) _
      + 0.000007 * Cos(2 * l(2) - 3 * l(3) + P(3)) _
      - 0.0000051 * Cos(l(3) + P(3) - 2 * BigPi - 2 * G)
    r3 = 14.9924 * (1 + dr)
End Function


Sub Callisto(T As Double, ByRef v As TSVECTOR)
' { Callisto }
  Call JSats(T)
  dl(4) = dl4
  l(4) = l(4) + dl(4)
  v.l = l(4)
  v.B = b4
  l(4) = l(4) - dl(4)
  v.r = r4
End Sub

Private Function dl4() As Double

Dim dl As Double
    dl = _
      0.84109 * Sin(l(4) - P(4)) _
      + 0.03429 * Sin(P(4) - P(3)) _
      - 0.03305 * Sin(2 * (Psi - BigPi)) _
      - 0.03211 * Sin(G) _
      - 0.0186 * Sin(l(4) - P(3)) _
      + 0.01182 * Sin(Psi - Om(4)) _
      + 0.00622 * Sin(l(4) + P(4) - 2 * G - 2 * BigPi)
    dl = dl + _
      0.00385 * Sin(2 * (l(4) - P(4))) _
      - 0.00284 * Sin(5 * Gd - 2 * G + 52.225 * DToR) _
      - 0.00233 * Sin(2 * (Psi - P(4))) _
      - 0.00223 * Sin(l(3) - l(4)) _
      - 0.00208 * Sin(l(4) - BigPi) _
      + 0.00177 * Sin(Psi + Om(4) - 2 * P(4))
    dl = dl + _
      0.00134 * Sin(P(4) - BigPi) _
      + 0.00125 * Sin(2 * (l(4) - G - BigPi)) _
      - 0.00117 * Sin(2 * G) _
      - 0.00112 * Sin(2 * (l(3) - l(4))) _
      + 0.00106 * Sin(3 * l(3) - 7 * l(4) + 4 * P(4)) _
      + 0.00102 * Sin(l(4) - G - BigPi)
    dl = dl + _
      0.00096 * Sin(2 * l(4) - Psi - Om(4)) _
      + 0.00087 * Sin(2 * (Psi - Om(4))) _
      - 0.00087 * Sin(3 * l(3) - 7 * l(4) + P(3) + 3 * P(4)) _
      + 0.00085 * Sin(l(3) - 2 * l(4) + P(4)) _
      - 0.00081 * Sin(2 * (l(4) - Psi)) _
      + 0.00071 * Sin(l(4) + P(4) - 2 * BigPi - 3 * G)
    dl = dl + _
      0.0006 * Sin(l(1) - l(4)) _
      - 0.00056 * Sin(Psi - Om(3)) _
      - 0.00055 * Sin(l(3) - 2 * l(4) + P(3)) _
      + 0.00051 * Sin(l(2) - l(4)) _
      + 0.00042 * Sin(2 * (Psi - G - BigPi)) _
      + 0.00039 * Sin(2 * (P(4) - Om(4)))
    dl = dl + _
      0.00036 * Sin(Psi + BigPi - P(4) - Om(4)) _
      + 0.00035 * Sin(2 * Gd - G + 188.37 * DToR) _
      - 0.00035 * Sin(l(4) - P(4) + 2 * BigPi - 2 * Psi) _
      - 0.00032 * Sin(l(4) + P(4) - 2 * BigPi - G) _
      + 0.0003 * Sin(3 * l(3) - 7 * l(4) + 2 * P(3) + 2 * P(4)) _
      + 0.0003 * Sin(2 * Gd - 2 * G + 149.15 * DToR)
    dl = dl + _
      0.00028 * Sin(l(4) - P(4) + 2 * Psi - 2 * BigPi) _
      - 0.00028 * Sin(2 * (l(4) - Om(4))) _
      - 0.00027 * Sin(P(3) - P(4) + Om(3) - Om(4)) _
      - 0.00026 * Sin(5 * Gd - 3 * G + 188.37 * DToR) _
      + 0.00025 * Sin(Om(4) - Om(3)) _
      - 0.00025 * Sin(l(2) - 3 * l(3) + 2 * l(4))
    dl = dl + _
      -0.00023 * Sin(3 * (l(3) - l(4))) _
      + 0.00021 * Sin(2 * l(4) - 2 * BigPi - 3 * G) _
      - 0.00021 * Sin(2 * l(3) - 3 * l(4) + P(4)) _
      + 0.00019 * Sin(l(4) - P(4) - G) _
      - 0.00019 * Sin(2 * l(4) - P(3) - P(4)) _
      - 0.00018 * Sin(l(4) - P(4) + G) _
      - 0.00016 * Sin(l(4) + P(3) - 2 * BigPi - 2 * G)
    dl4 = dl * DToR
End Function

Private Function b4() As Double
Dim tanb As Double
    tanb = _
      -0.0076579 * Sin(l(4) - Psi) _
      + 0.0044148 * Sin(l(4) - Om(4)) _
      - 0.0005106 * Sin(l(4) - Om(3)) _
      + 0.0000773 * Sin(l(4) + Psi - 2 * BigPi - 2 * G)
    tanb = tanb _
      + 0.0000104 * Sin(l(4) - Psi + G) _
      - 0.0000102 * Sin(l(4) - Psi - G) _
      + 0.0000088 * Sin(l(4) + Psi - 2 * BigPi - 3 * G) _
      - 0.0000038 * Sin(l(4) + Psi - 2 * BigPi - G)
    b4 = Atn(tanb)
End Function

Private Function r4() As Double
Dim dr As Double
    dr = _
      -0.0073391 * Cos(l(4) - P(4)) _
      + 0.000162 * Cos(l(4) - P(3)) _
      + 0.0000974 * Cos(l(3) - l(4)) _
      - 0.0000541 * Cos(l(4) + P(4) - 2 * BigPi - 2 * G) _
      - 0.0000269 * Cos(2 * (l(4) - P(4)))
    dr = dr + _
     0.0000182 * Cos(l(4) - BigPi) _
     + 0.0000177 * Cos(2 * (l(3) - l(4))) _
     - 0.0000167 * Cos(2 * l(4) - Psi - Om(4)) _
     + 0.0000167 * Cos(Psi - Om(4)) _
     - 0.0000155 * Cos(2 * (l(4) - BigPi - G)) _
     + 0.0000142 * Cos(2 * (l(4) - Psi))
    dr = dr + _
      0.0000104 * Cos(l(1) - l(4)) _
      + 0.0000092 * Cos(l(2) - l(4)) _
      - 0.0000089 * Cos(l(4) - BigPi - G) _
      - 0.0000062 * Cos(l(4) + P(4) - 2 * BigPi - 3 * G) _
      + 0.0000048 * Cos(2 * (l(4) - Om(4)))
    r4 = 26.3699 * (1 + dr)
End Function


'(*****************************************************************************)
'(* Name:    JSatEclipticPosition                                             *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Calculate the ecliptic position of a satellite of Jupiter.       *)
'(* Arguments:                                                                *)
'(*   n : number of satellite (Io := 1, Callisto := 4)                        *)
'(*   T : time in centuries since J2000.0                                     *)
'(*   v : TVECTOR to hold the coordinates                                     *)
'(*****************************************************************************)

Sub JSatEclipticPosition(ByVal n As Long, ByVal T As Double, ByRef v As TVECTOR)

Dim Psi As Double, Arg As Double, Omega As Double, C As Double, s As Double, r   As Double
Dim vsSatellite As TSVECTOR
  Psi = (316.50043 - 0.075972 * T) * DToR
  If (n = DUMMY_SATELLITE) Then
    v.x = 0
    v.Y = 0
    v.Z = 1
    r = 1
  Else
    Select Case n
        Case 1: Call Io(T, vsSatellite)
        Case 2: Call Europa(T, vsSatellite)
        Case 3: Call Ganymede(T, vsSatellite)
        Case 4: Call Callisto(T, vsSatellite)
    End Select
    vsSatellite.l = vsSatellite.l - Psi
    r = vsSatellite.r
    Call SphToRect(vsSatellite, v)
  End If

  Arg = (3.120262 + 0.0006 * T) * DToR
  C = Cos(Arg)
  s = Sin(Arg)
  Call XRot(v, C, s, v)

  Omega = (100.464441 + T * (1.020955 + T * (0.00040117 + T * 0.000000569))) * DToR
  '{ General precession }
  Arg = (T - TB1950)
  Arg = Arg * (1.3966626 + Arg * 0.0003088) * DToR
  Arg = Psi + Arg - Omega
  C = Cos(Arg)
  s = Sin(Arg)
  Call ZRot(v, C, s, v)

  Arg = (1.30327 + T * (-0.0054966 + T * (0.00000465 - T * 0.000000004))) * DToR
  C = Cos(Arg)
  s = Sin(Arg)
  Call XRot(v, C, s, v)

  Arg = Omega
  C = Cos(Arg)
  s = Sin(Arg)
  Call ZRot(v, C, s, v)
End Sub

'(*****************************************************************************)
'(* Name:    JSatViewFrom                                                     *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Calculate the 'objectocentric' position of a satellite of        *)
'(*          Jupiter                                                          *)
'(* Arguments:                                                                *)
'(*   n : number of satellite (Io = 1, Callisto = 4)                          *)
'(*   v : TVECTOR holding the ecliptical coordinates                          *)
'(*   vsOrigin : spherical coordinates of Jupiter relative to object          *)
'(*   vDummy : TVECTOR holding the coordinates of the fifth 'dummy' satellite *)
'(*   bScale : non-zero if Y-coordinate should be scaled to account for the   *)
'(*            flattening of Jupiter's disk                                   *)
'(*   w : TVECTOR to hold the final X, Y, Z coordinates                       *)
'(*****************************************************************************)

Sub JSatViewFrom(ByVal n As Long, ByRef v As TVECTOR, ByRef vsOrigin As TSVECTOR, _
          ByRef vdummy As TVECTOR, bScale As Boolean, ByRef W As TVECTOR, bPerspective As Boolean)

Dim Arg As Double, C As Double, s As Double, r As Double, rDummy  As Double
Dim k As Variant ' K(4) As Double

  k = Array(0, 17295#, 21819#, 27558#, 36548#)
  Arg = PI / 2 - vsOrigin.l
  C = Cos(Arg)
  s = Sin(Arg)
  Call ZRot(v, C, s, W)

  Arg = W.Z
  W.Z = W.Y
  W.Y = Arg

  r = Sqr(v.x * v.x + v.Y * v.Y + v.Z * v.Z)
  Arg = vsOrigin.B
  C = Cos(Arg)
  s = Sin(Arg)
  Call XRot(W, C, s, W)

  If (n <> DUMMY_SATELLITE) Then
    '{ 'Rectification' }
    rDummy = Sqr(vdummy.x * vdummy.x + vdummy.Y * vdummy.Y)
    C = vdummy.Y / rDummy
    s = vdummy.x / rDummy
    Call ZRot(W, C, s, W)

    '{ Scaling }
    If bScale Then W.Y = W.Y * 1.071374

    '{ Light time correction }
    Arg = W.x / r
    W.x = W.x + Abs(W.Z) / k(n) * Sqr(1 - Arg * Arg)

    '{ Perspective effect }
    If bPerspective Then
        Arg = vsOrigin.r / (vsOrigin.r + W.Z / 2095)
        W.x = W.x * Arg
        W.Y = W.Y * Arg
    End If
  End If
End Sub

