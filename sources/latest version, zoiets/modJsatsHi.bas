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

Private L(4) As Double, dl(4) As Double, P(4) As Double, Om(4) As Double
Private Gamma As Double, PhiL As Double, Psi As Double, g As Double, Gd As Double, BigPi   As Double
Private Tcur   As Double

Sub JSats(T As Double)

Dim I As Long
Dim smallt As Double

lCoeff = Array(Array(0, 0), Array(106.07719, 203.48895579), Array(175.73161, 101.374724735), _
         Array(120.55883, 50.317609207), Array(84.44459, 21.571071177))
pCoeff = Array(Array(0, 0), Array(97.0881, 0.16138586), Array(154.8663, 0.04726307), _
         Array(188.184, 0.00712734), Array(335.2868, 0.00184))
omCoeff = Array(Array(0, 0), Array(312.3346, -0.13279386), Array(100.4411, -0.03263064), _
         Array(119.1942, -0.00717703), Array(322.6186, -0.00175934))

'If (Tcur <> T) Then
    smallt = (T * 36525# + 8544.5)
    Gamma = 0.33033 * Sin((163.679 + 0.0010512 * smallt) * DToR)
    Gamma = Gamma + 0.03439 * Sin((34.486 - 0.0161731 * smallt) * DToR)
    PhiL = (199.6766 + 0.173919 * smallt) * DToR
    Psi = (316.5182 - 0.00000208 * smallt) * DToR
    g = (30.23756 + 0.0830925701 * smallt + Gamma) * DToR
    Gd = (31.97853 + 0.0334597339 * smallt) * DToR
    BigPi = 13.469942 * DToR
    For I = 1 To 4
      L(I) = modpi2((lCoeff(I)(0) + smallt * lCoeff(I)(1)) * DToR)
      P(I) = (pCoeff(I)(0) + smallt * pCoeff(I)(1)) * DToR
      Om(I) = (omCoeff(I)(0) + smallt * omCoeff(I)(1)) * DToR
    Next
    Tcur = T
'End If
End Sub

Sub IO(T As Double, ByRef v As TSVECTOR)
' { Io }
  Call JSats(T)
  dl(1) = dl1
  L(1) = L(1) + dl(1)
  v.L = L(1)
  v.B = b1
  L(1) = L(1) - dl(1)
  v.r = r1
End Sub

Private Function dl1() As Double

  Dim dl As Double
'    dl = _
      0.47259 * Sin(2 * (L(1) - L(2))) _
      - 0.0348 * Sin(P(3) - P(4)) _
      - 0.01756 * Sin(P(1) + P(3) - 2 * BigPi - 2 * g) _
      + 0.0108 * Sin(L(2) - 2 * L(3) + P(3)) _
      + 0.00757 * Sin(PhiL)
'     dl = dl + _
      0.00663 * Sin(L(2) - 2 * L(3) + P(4)) _
      + 0.00453 * Sin(L(1) - P(3)) _
      + 0.00453 * Sin(L(2) - 2 * L(3) + P(2)) _
      - 0.00354 * Sin(L(1) - L(2)) _
      - 0.00317 * Sin(2 * Psi - 2 * BigPi)
'    dl = dl + _
      -0.00269 * Sin(L(2) - 2 * L(3) + P(1)) _
      + 0.00263 * Sin(L(1) - P(4)) _
      + 0.00186 * Sin(L(1) - P(1)) _
      - 0.00186 * Sin(g) _
      + 0.00167 * Sin(P(2) - P(3))
'    dl = dl + _
      0.00158 * Sin(4 * (L(1) - L(2))) _
      - 0.00155 * Sin(L(1) - L(3)) _
      - 0.00142 * Sin(Psi + Om(3) - 2 * BigPi - 2 * g) _
      - 0.00115 * Sin(2 * (L(1) - 2 * L(2) + Om(2))) _
      + 0.00089 * Sin(P(2) - P(4))
'    dl = dl + _
      0.00084 * Sin(Om(2) - Om(3)) _
      + 0.00084 * Sin(L(1) + P(3) - 2 * BigPi - 2 * g) _
      + 0.00053 * Sin(Psi - Om(2))
         dl = _
      0.47259 * Sin(2 * (L(1) - L(2))) _
    - 0.03478 * Sin(P(3) - P(4)) _
    + 0.01081 * Sin(L(2) - 2 * L(3) + P(3)) _
    + 0.00738 * Sin(PhiL) _
    + 0.00713 * Sin(L(2) - 2 * L(3) + P(2))

    dl = dl + _
    -0.00674 * Sin(P(1) + P(3) - 2 * BigPi - 2 * g) _
    + 0.00666 * Sin(L(2) - 2 * L(3) + P(4)) _
    + 0.00445 * Sin(L(1) - P(3)) _
    - 0.00354 * Sin(L(1) - L(2)) _
    - 0.00317 * Sin(2 * Psi - 2 * BigPi)

    dl = dl + _
    0.00265 * Sin(L(1) - P(4)) _
    - 0.00186 * Sin(g) _
    + 0.00162 * Sin(P(2) - P(3)) _
    + 0.00158 * Sin(4 * (L(1) - L(2))) _
    - 0.00155 * Sin(L(1) - L(3)) _

    dl = dl + _
    -0.00138 * Sin(Psi + Om(3) - 2 * BigPi - 2 * g) _
    - 0.00115 * Sin(2 * (L(1) - 2 * L(2) + Om(2))) _
    + 0.00089 * Sin(P(2) - P(4))

    dl = dl + _
    0.00083 * Sin(Om(2) - Om(3)) _
    + 0.00085 * Sin(L(1) + P(3) - 2 * BigPi - 2 * g) _
    + 0.00053 * Sin(Psi - Om(2))
    dl1 = dl * DToR
End Function

Private Function b1() As Double

Dim tanb As Double
'    tanb = _
      0.0006502 * Sin(L(1) - Om(1)) _
      + 0.0001835 * Sin(L(1) - Om(2)) _
      + 0.0000329 * Sin(L(1) - Om(3)) _
      - 0.0000311 * Sin(L(1) - Psi) _
      + 0.0000093 * Sin(L(1) - Om(4)) _
      + 0.0000075 * Sin(3 * L(1) - 4 * L(2) - 1.9927 * dl(1) + Om(2)) _
      + 0.0000046 * Sin(L(1) + Psi - 2 * BigPi - 2 * g)
          tanb = _
    0.0006393 * Sin(L(1) - Om(1)) _
    + 0.0001825 * Sin(L(1) - Om(2)) _
    + 0.0000329 * Sin(L(1) - Om(3)) _
    - 0.0000311 * Sin(L(1) - Psi) _
    + 0.0000093 * Sin(L(1) - Om(4)) _
    + 0.0000075 * Sin(3 * L(1) - 4 * L(2) - 1.9927 * dl(1) + Om(2)) _
    + 0.0000046 * Sin(L(1) + Psi - 2 * BigPi - 2 * g)
    b1 = Atn(tanb)
End Function

Private Function r1() As Double

Dim dr As Double
'    dr = _
      -0.0041339 * Cos(2 * (L(1) - L(2))) _
      - 0.0000395 * Cos(L(1) - P(3)) _
      - 0.0000214 * Cos(L(1) - P(4)) _
      + 0.000017 * Cos(L(1) - L(2))
'    dr = dr + _
      -0.000013 * Cos(4 * (L(1) - L(2))) _
      + 0.0000106 * Cos(L(1) - L(3)) _
      - 0.0000162 * Cos(L(1) - P(1)) _
      - 0.0000063 * Cos(L(1) + P(3) - 2 * BigPi - 2 * g)
'    r1 = 5.9073 * (1 + dr)
    dr = _
      -0.0041339 * Cos(2 * (L(1) - L(2))) _
      - 0.0000387 * Cos(L(1) - P(3)) _
      - 0.0000214 * Cos(L(1) - P(4)) _
      + 0.000017 * Cos(L(1) - L(2))

    dr = dr + _
     -0.0000131 * Cos(4 * (L(1) - L(2))) _
     + 0.0000106 * Cos(L(1) - L(3)) _
     - 0.0000066 * Cos(L(1) + P(3) - 2 * BigPi - 2 * g)
    r1 = 5.90569 * (1 + dr)
End Function

Sub Europa(T As Double, ByRef v As TSVECTOR)
' { Europa }
  Call JSats(T)
  dl(2) = dl2
  L(2) = L(2) + dl(2)
  v.L = L(2)
  v.B = b2
  L(2) = L(2) - dl(2)
  v.r = r2
End Sub

Private Function dl2() As Double

Dim dl As Double
    dl = _
     1.06476 * Sin(2 * (L(2) - L(3))) _
     + 0.04256 * Sin(L(1) - 2 * L(2) + P(3)) _
     + 0.03581 * Sin(L(2) - P(3)) _
     + 0.02395 * Sin(L(1) - 2 * L(2) + P(4)) _
     + 0.01984 * Sin(L(2) - P(4)) _
     - 0.01778 * Sin(PhiL)

    dl = dl + _
      0.01654 * Sin(L(2) - P(2)) _
      + 0.01334 * Sin(L(2) - 2 * L(3) + P(2)) _
      + 0.01294 * Sin(P(3) - P(4)) _
      - 0.01142 * Sin(L(2) - L(3)) _
      - 0.01057 * Sin(g) _
      - 0.00775 * Sin(2 * (Psi - BigPi)) _
      + 0.00524 * Sin(2 * (L(1) - L(2))) _
      - 0.0046 * Sin(L(1) - L(3))
 

    dl = dl + _
      0.00316 * Sin(Psi - 2 * g + Om(3) - 2 * BigPi) _
      - 0.00203 * Sin(P(1) + P(3) - 2 * BigPi - 2 * g) _
      + 0.00146 * Sin(Psi - Om(3)) _
      - 0.00145 * Sin(2 * g) _
      + 0.00125 * Sin(Psi - Om(4))

    dl = dl + _
      -0.00115 * Sin(L(1) - 2 * L(3) + P(3)) _
      - 0.00094 * Sin(2 * (L(2) - Om(2))) _
      + 0.00086 * Sin(2 * (L(1) - 2 * L(2) + Om(2))) _
      - 0.00086 * Sin(5 * Gd - 2 * g + 52.225 * DToR)

    dl = dl + _
      -0.00078 * Sin(L(2) - L(4)) _
      - 0.00064 * Sin(3 * L(3) - 7 * L(4) + 4 * P(4)) _
      + 0.00064 * Sin(P(1) - P(4)) _
      - 0.00063 * Sin(L(1) - 2 * L(3) + P(4)) _
      + 0.00058 * Sin(Om(3) - Om(4)) _
      + 0.00056 * Sin(2 * (Psi - BigPi - g))

    dl = dl + _
      0.00056 * Sin(2 * (L(2) - L(4))) _
      + 0.00055 * Sin(2 * (L(1) - L(3))) _
      + 0.00052 * Sin(3 * L(3) - 7 * L(4) + P(3) + 3 * P(4)) _
      - 0.00043 * Sin(L(1) - P(3)) _
      + 0.00041 * Sin(5 * (L(2) - L(3)))

    dl = dl + _
      0.00041 * Sin(P(4) - BigPi) _
      + 0.00032 * Sin(Om(2) - Om(3)) _
      + 0.00032 * Sin(2 * (L(3) - g - BigPi))
'     dl = _
      1.06476 * Sin(2 * (L(2) - L(3))) _
      + 0.04253 * Sin(L(1) - 2 * L(2) + P(3)) _
      + 0.03579 * Sin(L(2) - P(3)) _
      + 0.02383 * Sin(L(1) - 2 * L(2) + P(4)) _
      + 0.01977 * Sin(L(2) - P(4)) _
      - 0.01843 * Sin(PhiL)
  '   dl = dl + _
      0.01299 * Sin(P(3) - P(4)) _
      - 0.01142 * Sin(L(2) - L(3)) _
      - 0.01058 * Sin(g) _
      + 0.01078 * Sin(L(2) - P(2)) _
      + 0.00327 * Sin(Psi - 2 * g + Om(3) - 2 * BigPi) _
      + 0.0087 * Sin(L(2) - 2 * L(3) + P(2))
'     dl = dl + _
      -0.00775 * Sin(Psi - BigPi) _
      + 0.00524 * Sin(2 * (L(1) - L(2))) _
      - 0.0046 * Sin(L(1) - L(3)) _
      + 0.0045 * Sin(L(2) - 2 * L(3) + P(1)) _
      - 0.00296 * Sin(P(1) + P(3) - 2 * BigPi - 2 * g) _
      - 0.00151 * Sin(2 * g)
    ' dl = dl + _
      0.00146 * Sin(Psi - Om(3)) _
      + 0.00125 * Sin(Psi - Om(4)) _
      - 0.00117 * Sin(L(1) - 2 * L(3) + P(3)) _
      - 0.00095 * Sin(2 * (L(2) - Om(2))) _
      + 0.00086 * Sin(2 * (L(1) - 2 * L(2) + Om(2))) _
      - 0.00086 * Sin(5 * Gd - 2 * g + 52.225 * DToR)
    ' dl = dl + _
      -0.00078 * Sin(L(2) - L(4)) _
      - 0.00064 * Sin(L(1) - 2 * L(3) + P(4)) _
      - 0.00063 * Sin(3 * L(3) - 7 * L(4) + 4 * P(4)) _
      + 0.00061 * Sin(P(1) - P(4)) _
      + 0.00058 * Sin(2 * (Psi - BigPi - g)) _
      + 0.00058 * Sin(Om(3) - Om(4))
    ' dl = dl + _
      0.00056 * Sin(2 * (L(2) - L(4))) _
      + 0.00055 * Sin(2 * (L(1) - L(3))) _
      + 0.00052 * Sin(3 * L(3) - 7 * L(4) + P(3) + 3 * P(4)) _
      - 0.00043 * Sin(L(1) - P(3)) _
      + 0.00042 * Sin(P(3) - P(2)) _
      + 0.00041 * Sin(5 * (L(2) - L(3)))
    ' dl = dl + _
      0.00041 * Sin(P(4) - BigPi) _
      + 0.00032 * Sin(Om(2) - Om(3)) _
      + 0.00032 * Sin(2 * (L(3) - g - BigPi)) _
      + 0.00029 * Sin(P(1) - P(3)) _
      + 0.00038 * Sin(L(2) - P(1))
    dl2 = dl * DToR
End Function

Private Function b2() As Double

Dim tanb As Double
'     tanb = _
      0.0081275 * Sin(L(2) - Om(2)) _
      + 0.0004512 * Sin(L(2) - Om(3)) _
      - 0.0003286 * Sin(L(2) - Psi) _
      + 0.0001164 * Sin(L(2) - Om(4)) _
      + 0.0000273 * Sin(L(1) - 2 * L(3) + 1.0146 * dl(2) + Om(2))
'     tanb = tanb + _
      0.0000143 * Sin(L(2) + Psi - 2 * BigPi - 2 * g) _
      - 0.0000143 * Sin(L(2) - Om(1)) _
      + 0.0000035 * Sin(L(2) - Psi + g) _
      - 0.0000028 * Sin(L(1) - 2 * L(3) + 1.0146 * dl(2) + Om(3))
          tanb = _
     0.0081004 * Sin(L(2) - Om(2)) _
     + 0.0004512 * Sin(L(2) - Om(3)) _
     - 0.0003284 * Sin(L(2) - Psi) _
     + 0.000116 * Sin(L(2) - Om(4)) _
     + 0.0000272 * Sin(L(1) - 2 * L(3) + 1.0146 * dl(2) + Om(2))

    tanb = tanb + _
    -0.0000144 * Sin(L(2) - Om(1)) _
    + 0.0000143 * Sin(L(2) + Psi - 2 * BigPi - 2 * g) _
    + 0.0000035 * Sin(L(2) - Psi + g) _
    - 0.0000028 * Sin(L(1) - 2 * L(3) + 1.0146 * dl(2) + Om(3))
    b2 = Atn(tanb)
End Function

Private Function r2() As Double

Dim dr As Double
'    dr = _
      0.0093847 * Cos(L(1) - L(2)) _
      - 0.0003114 * Cos(L(2) - P(3)) _
      - 0.0001738 * Cos(L(2) - P(4)) _
      - 0.0000941 * Cos(L(2) - P(2)) _
      + 0.0000553 * Cos(L(2) - L(3)) _
      + 0.0000523 * Cos(L(1) - L(3))
'    dr = dr + _
      -0.000029 * Cos(2 * (L(1) - L(2))) _
      + 0.0000166 * Cos(2 * (L(2) - Om(2))) _
      + 0.0000107 * Cos(L(1) - 2 * L(3) + P(3)) _
      - 0.0000102 * Cos(L(2) - P(1)) _
      - 0.0000091 * Cos(2 * (L(1) - L(3)))
   ' r2 = 9.39912 * (1 + dr)
        dr = _
      0.0093848 * Cos(L(1) - L(2)) _
      - 0.0003116 * Cos(L(2) - P(3)) _
      - 0.0001744 * Cos(L(2) - P(4)) _
      - 0.0001442 * Cos(L(2) - P(2)) _
      + 0.0000553 * Cos(L(2) - L(3)) _
      + 0.0000523 * Cos(L(1) - L(3))

    dr = dr + _
      -0.000029 * Cos(2 * (L(1) - L(2))) _
      + 0.0000164 * Cos(2 * (L(2) - Om(2))) _
      + 0.0000107 * Cos(L(1) - 2 * L(3) + P(3)) _
      - 0.0000102 * Cos(L(2) - P(1)) _
      - 0.0000091 * Cos(2 * (L(1) - L(3)))

    r2 = 9.39657 * (1 + dr)
End Function


Sub Ganymede(T As Double, ByRef v As TSVECTOR)
'{ Ganymede }
  Call JSats(T)
  dl(3) = dl3
  L(3) = L(3) + dl(3)
  v.L = L(3)
  v.B = b3
  L(3) = L(3) - dl(3)
  v.r = r3
End Sub

Private Function dl3() As Double

Dim dl As Double
'     dl = _
      0.16477 * Sin(L(3) - P(3)) _
      + 0.09062 * Sin(L(3) - P(4)) _
      - 0.06907 * Sin(L(2) - L(3)) _
      + 0.03786 * Sin(P(3) - P(4)) _
      + 0.01844 * Sin(2 * (L(3) - L(4))) _
      - 0.0134 * Sin(g)
'     dl = dl + _
      -0.0067 * Sin(2 * (Psi - BigPi)) _
      + 0.00703 * Sin(L(2) - 2 * L(3) + P(3)) _
      - 0.0054 * Sin(L(3) - L(4)) _
      - 0.00409 * Sin(L(2) - 2 * L(3) + P(2)) _
      + 0.00379 * Sin(L(2) - 2 * L(3) + P(4)) _
      + 0.00481 * Sin(P(1) + P(3) - 2 * BigPi - 2 * g)
'     dl = dl + _
      0.00235 * Sin(Psi - Om(3)) _
      + 0.00198 * Sin(Psi - Om(4)) _
      + 0.0018 * Sin(PhiL) _
      + 0.00124 * Sin(L(1) - L(3)) _
      - 0.00119 * Sin(5 * Gd - 2 * g + 52.225 * DToR) _
      + 0.00109 * Sin(L(1) - L(2))
'     dl = dl + _
      0.00129 * Sin(3 * (L(3) - L(4))) _
      - 0.00099 * Sin(3 * L(3) - 7 * L(4) + 4 * P(4)) _
      - 0.00029 * Sin(Om(3) + Psi - 2 * BigPi - 2 * g) _
      + 0.00091 * Sin(Om(3) - Om(4)) _
      + 0.00081 * Sin(3 * L(3) - 7 * L(4) + P(3) + 3 * P(4)) _
      - 0.00076 * Sin(2 * L(2) - 3 * L(3) + P(3))
'     dl = dl + _
      0.00069 * Sin(P(4) - BigPi) _
      - 0.00058 * Sin(2 * L(3) - 3 * L(4) + P(4)) _
      + 0.00057 * Sin(L(3) + P(3) - 2 * BigPi - 2 * g) _
      - 0.00057 * Sin(L(3) - 2 * L(4) + P(4)) _
      - 0.00052 * Sin(P(2) - P(3)) _
      - 0.00052 * Sin(L(2) - 2 * L(3) + P(1))
'     dl = dl + _
      0.00048 * Sin(L(3) - 2 * L(4) + P(3)) _
      - 0.00045 * Sin(2 * L(2) - 3 * L(3) + P(4)) _
      - 0.00041 * Sin(P(2) - P(4)) _
      - 0.00038 * Sin(2 * g) _
      - 0.00033 * Sin(P(3) - P(4) + Om(3) - Om(4)) _
      - 0.00032 * Sin(3 * L(3) - 7 * L(4) + 2 * P(3) + 2 * P(4))
'     dl = dl + _
      0.0003 * Sin(4 * (L(3) - L(4))) _
      + 0.00029 * Sin(L(3) + P(4) - 2 * BigPi - 2 * g) _
      + 0.00026 * Sin(L(3) - BigPi - g) _
      + 0.00024 * Sin(L(2) - 3 * L(3) + 2 * L(4)) _
      + 0.00021 * Sin(2 * (L(3) - BigPi - g)) _
      - 0.00021 * Sin(L(3) - P(2)) _
      + 0.00017 * Sin(2 * (L(3) - P(3)))
         dl = _
      0.1649 * Sin(L(3) - P(3)) _
      + 0.09081 * Sin(L(3) - P(4)) _
      - 0.06907 * Sin(L(2) - L(3)) _
      + 0.03784 * Sin(P(3) - P(4)) _
      + 0.01846 * Sin(2 * (L(3) - L(4))) _
      - 0.0134 * Sin(g)

    dl = dl + _
      -0.01014 * Sin(2 * (Psi - BigPi)) _
      + 0.00704 * Sin(L(2) - 2 * L(3) + P(3)) _
      - 0.0062 * Sin(L(2) - 2 * L(3) + P(2)) _
      - 0.00541 * Sin(L(3) - L(4)) _
      + 0.00381 * Sin(L(2) - 2 * L(3) + P(4))
    
    dl = dl + _
      0.00235 * Sin(Psi - Om(3)) _
      + 0.00198 * Sin(Psi - Om(4)) _
      + 0.00176 * Sin(PhiL) _
      + 0.0013 * Sin(3 * (L(3) - L(4))) _
      + 0.00125 * Sin(L(1) - L(3)) _
      - 0.00119 * Sin(5 * Gd - 2 * g + 52.225 * DToR) _
      + 0.00109 * Sin(L(1) - L(2))

    dl = dl + _
      -0.001 * Sin(3 * L(3) - 7 * L(4) + 4 * P(4)) _
      + 0.00091 * Sin(Om(3) - Om(4)) _
      + 0.0008 * Sin(3 * L(3) - 7 * L(4) + P(3) + 3 * P(4)) _
      - 0.00075 * Sin(2 * L(2) - 3 * L(3) + P(3))

    dl = dl + _
      0.00072 * Sin(P(1) + P(3) - 2 * BigPi - 2 * g) _
      + 0.00069 * Sin(P(4) - BigPi) _
      - 0.00058 * Sin(2 * L(3) - 3 * L(4) + P(4)) _
      - 0.00057 * Sin(L(3) - 2 * L(4) + P(4)) _
      + 0.00056 * Sin(L(3) + P(3) - 2 * BigPi - 2 * g) _
      - 0.00052 * Sin(L(2) - 2 * L(3) + P(1)) _
      - 0.0005 * Sin(P(2) - P(3))

    dl = dl + _
      0.00048 * Sin(L(3) - 2 * L(4) + P(3)) _
      - 0.00045 * Sin(2 * L(2) - 3 * L(3) + P(4)) _
      - 0.00041 * Sin(P(2) - P(4)) _
      - 0.00038 * Sin(2 * g) _
      - 0.00037 * Sin(P(3) - P(4) + Om(3) - Om(4)) _
      - 0.00032 * Sin(3 * L(3) - 7 * L(4) + 2 * P(3) + 2 * P(4))

    dl = dl + _
      0.0003 * Sin(4 * (L(3) - L(4))) _
      + 0.00029 * Sin(L(3) + P(4) - 2 * BigPi - 2 * g) _
      - 0.00028 * Sin(Om(3) + Psi - 2 * BigPi - 2 * g) _
      + 0.00026 * Sin(L(3) - BigPi - g) _
      + 0.00024 * Sin(L(2) - 3 * L(3) + 2 * L(4)) _
      + 0.00021 * Sin(2 * (L(3) - BigPi - g)) _
      - 0.00021 * Sin(L(3) - P(2)) _
      + 0.00017 * Sin(2 * (L(3) - P(3)))
      
  dl3 = dl * DToR
End Function

Private Function b3() As Double

Dim tanb As Double
'    tanb = _
      0.0032364 * Sin(L(3) - Om(3)) _
      - 0.0016911 * Sin(L(3) - Psi) _
      + 0.0006849 * Sin(L(3) - Om(4)) _
      - 0.0002806 * Sin(L(3) - Om(2)) _
      + 0.0000321 * Sin(L(3) + Psi - 2 * BigPi - 2 * g) _
      + 0.0000051 * Sin(L(3) - Psi + g)
 '   tanb = tanb + _
      -0.0000045 * Sin(L(3) - Psi - g) _
      - 0.0000045 * Sin(L(3) + Psi - 2 * BigPi) _
      + 0.0000037 * Sin(L(3) + Psi - 2 * BigPi - 3 * g) _
      + 0.000003 * Sin(2 * L(2) - 3 * L(3) + 4.03 * dl(3) + Om(2)) _
      - 0.0000021 * Sin(2 * L(2) - 3 * L(3) + 4.03 * dl(3) + Om(3))
      
          tanb = _
      0.0032402 * Sin(L(3) - Om(3)) _
      - 0.0016911 * Sin(L(3) - Psi) _
      + 0.0006847 * Sin(L(3) - Om(4)) _
      - 0.0002797 * Sin(L(3) - Om(2)) _
      + 0.0000321 * Sin(L(3) + Psi - 2 * BigPi - 2 * g) _
      + 0.0000051 * Sin(L(3) - Psi + g)
      
    tanb = tanb + _
      -0.0000045 * Sin(L(3) - Psi - g) _
      - 0.0000045 * Sin(L(3) + Psi - 2 * BigPi) _
      + 0.0000037 * Sin(L(3) + Psi - 2 * BigPi - 3 * g) _
      + 0.000003 * Sin(2 * L(2) - 3 * L(3) + 4.03 * dl(3) + Om(2)) _
      - 0.0000021 * Sin(2 * L(2) - 3 * L(3) + 4.03 * dl(3) + Om(3))
    b3 = Atn(tanb)
End Function

Private Function r3() As Double
Dim dr As Double
'    dr = _
      -0.0014377 * Cos(L(3) - P(3)) _
      - 0.0007904 * Cos(L(3) - P(4)) _
      + 0.0006342 * Cos(L(2) - L(3)) _
      - 0.0001758 * Cos(2 * (L(3) - L(4))) _
      + 0.0000294 * Cos(L(3) - L(4))
 '   dr = dr + _
      -0.0000153 * Cos(L(1) - L(2)) _
      + 0.0000155 * Cos(L(1) - L(3)) _
      - 0.0000156 * Cos(3 * (L(3) - L(4))) _
      + 0.000007 * Cos(2 * L(2) - 3 * L(3) + P(3)) _
      - 0.0000051 * Cos(L(3) + P(3) - 2 * BigPi - 2 * g)
  '  r3 = 14.9924 * (1 + dr)
        dr = _
      -0.0014388 * Cos(L(3) - P(3)) _
      - 0.0007919 * Cos(L(3) - P(4)) _
      + 0.0006342 * Cos(L(2) - L(3)) _
      - 0.0001761 * Cos(2 * (L(3) - L(4))) _
      + 0.0000294 * Cos(L(3) - L(4))

    dr = dr + _
      -0.0000156 * Cos(3 * (L(3) - L(4))) _
      + 0.0000156 * Cos(L(1) - L(3)) _
      - 0.0000153 * Cos(L(1) - L(2)) _
      + 0.000007 * Cos(2 * L(2) - 3 * L(3) + P(3)) _
      - 0.0000051 * Cos(L(3) + P(3) - 2 * BigPi - 2 * g)

    r3 = 14.98832 * (1 + dr)
End Function


Sub Callisto(T As Double, ByRef v As TSVECTOR)
' { Callisto }
  Call JSats(T)
  dl(4) = dl4
  L(4) = L(4) + dl(4)
  v.L = L(4)
  v.B = b4
  L(4) = L(4) - dl(4)
  v.r = r4
End Sub

Private Function dl4() As Double

Dim dl As Double
'    dl = _
      0.84109 * Sin(L(4) - P(4)) _
      + 0.03429 * Sin(P(4) - P(3)) _
      - 0.03305 * Sin(2 * (Psi - BigPi)) _
      - 0.03211 * Sin(g) _
      - 0.0186 * Sin(L(4) - P(3)) _
      + 0.01182 * Sin(Psi - Om(4)) _
      + 0.00622 * Sin(L(4) + P(4) - 2 * g - 2 * BigPi)
'    dl = dl + _
      0.00385 * Sin(2 * (L(4) - P(4))) _
      - 0.00284 * Sin(5 * Gd - 2 * g + 52.225 * DToR) _
      - 0.00233 * Sin(2 * (Psi - P(4))) _
      - 0.00223 * Sin(L(3) - L(4)) _
      - 0.00208 * Sin(L(4) - BigPi) _
      + 0.00177 * Sin(Psi + Om(4) - 2 * P(4))
'    dl = dl + _
      0.00134 * Sin(P(4) - BigPi) _
      + 0.00125 * Sin(2 * (L(4) - g - BigPi)) _
      - 0.00117 * Sin(2 * g) _
      - 0.00112 * Sin(2 * (L(3) - L(4))) _
      + 0.00106 * Sin(3 * L(3) - 7 * L(4) + 4 * P(4)) _
      + 0.00102 * Sin(L(4) - g - BigPi)
'    dl = dl + _
      0.00096 * Sin(2 * L(4) - Psi - Om(4)) _
      + 0.00087 * Sin(2 * (Psi - Om(4))) _
      - 0.00087 * Sin(3 * L(3) - 7 * L(4) + P(3) + 3 * P(4)) _
      + 0.00085 * Sin(L(3) - 2 * L(4) + P(4)) _
      - 0.00081 * Sin(2 * (L(4) - Psi)) _
      + 0.00071 * Sin(L(4) + P(4) - 2 * BigPi - 3 * g)
'    dl = dl + _
      0.0006 * Sin(L(1) - L(4)) _
      - 0.00056 * Sin(Psi - Om(3)) _
      - 0.00055 * Sin(L(3) - 2 * L(4) + P(3)) _
      + 0.00051 * Sin(L(2) - L(4)) _
      + 0.00042 * Sin(2 * (Psi - g - BigPi)) _
      + 0.00039 * Sin(2 * (P(4) - Om(4)))
'    dl = dl + _
      0.00036 * Sin(Psi + BigPi - P(4) - Om(4)) _
      + 0.00035 * Sin(2 * Gd - g + 188.37 * DToR) _
      - 0.00035 * Sin(L(4) - P(4) + 2 * BigPi - 2 * Psi) _
      - 0.00032 * Sin(L(4) + P(4) - 2 * BigPi - g) _
      + 0.0003 * Sin(3 * L(3) - 7 * L(4) + 2 * P(3) + 2 * P(4)) _
      + 0.0003 * Sin(2 * Gd - 2 * g + 149.15 * DToR)
'    dl = dl + _
      0.00028 * Sin(L(4) - P(4) + 2 * Psi - 2 * BigPi) _
      - 0.00028 * Sin(2 * (L(4) - Om(4))) _
      - 0.00027 * Sin(P(3) - P(4) + Om(3) - Om(4)) _
      - 0.00026 * Sin(5 * Gd - 3 * g + 188.37 * DToR) _
      + 0.00025 * Sin(Om(4) - Om(3)) _
      - 0.00025 * Sin(L(2) - 3 * L(3) + 2 * L(4))
'    dl = dl + _
      -0.00023 * Sin(3 * (L(3) - L(4))) _
      + 0.00021 * Sin(2 * L(4) - 2 * BigPi - 3 * g) _
      - 0.00021 * Sin(2 * L(3) - 3 * L(4) + P(4)) _
      + 0.00019 * Sin(L(4) - P(4) - g) _
      - 0.00019 * Sin(2 * L(4) - P(3) - P(4)) _
      - 0.00018 * Sin(L(4) - P(4) + g) _
      - 0.00016 * Sin(L(4) + P(3) - 2 * BigPi - 2 * g)
         dl = _
      0.84287 * Sin(L(4) - P(4)) _
      + 0.03431 * Sin(P(4) - P(3)) _
      - 0.03305 * Sin(2 * (Psi - BigPi)) _
      - 0.03211 * Sin(g) _
      - 0.01862 * Sin(L(4) - P(3)) _
      + 0.01186 * Sin(Psi - Om(4)) _
      + 0.00623 * Sin(L(4) + P(4) - 2 * g - 2 * BigPi)

    dl = dl + _
      0.00387 * Sin(2 * (L(4) - P(4))) _
      - 0.00284 * Sin(5 * Gd - 2 * g + 52.225 * DToR) _
      - 0.00234 * Sin(2 * (Psi - P(4))) _
      - 0.00223 * Sin(L(3) - L(4)) _
      - 0.00208 * Sin(L(4) - BigPi) _
      + 0.00178 * Sin(Psi + Om(4) - 2 * P(4))

    dl = dl + _
      0.00134 * Sin(P(4) - BigPi) _
      + 0.00125 * Sin(2 * (L(4) - g - BigPi)) _
      - 0.00117 * Sin(2 * g) _
      - 0.00112 * Sin(2 * (L(3) - L(4))) _
      + 0.00107 * Sin(3 * L(3) - 7 * L(4) + 4 * P(4)) _
      + 0.00102 * Sin(L(4) - g - BigPi)

    dl = dl + _
      0.00096 * Sin(2 * L(4) - Psi - Om(4)) _
      + 0.00087 * Sin(2 * (Psi - Om(4))) _
      - 0.00085 * Sin(3 * L(3) - 7 * L(4) + P(3) + 3 * P(4)) _
      + 0.00085 * Sin(L(3) - 2 * L(4) + P(4)) _
      - 0.00081 * Sin(2 * (L(4) - Psi)) _
      + 0.00071 * Sin(L(4) + P(4) - 2 * BigPi - 3 * g)

    dl = dl + _
      0.00061 * Sin(L(1) - L(4)) _
      - 0.00056 * Sin(Psi - Om(3)) _
      - 0.00054 * Sin(L(3) - 2 * L(4) + P(3)) _
      + 0.00051 * Sin(L(2) - L(4)) _
      + 0.00042 * Sin(2 * (Psi - g - BigPi)) _
      + 0.00039 * Sin(2 * (P(4) - Om(4)))

    dl = dl + _
      0.00036 * Sin(Psi + BigPi - P(4) - Om(4)) _
      + 0.00035 * Sin(2 * Gd - g + 188.37 * DToR) _
      - 0.00035 * Sin(L(4) - P(4) + 2 * BigPi - 2 * Psi) _
      - 0.00032 * Sin(L(4) + P(4) - 2 * BigPi - g) _
      + 0.0003 * Sin(2 * Gd - 2 * g + 149.15 * DToR) _
      + 0.00029 * Sin(3 * L(3) - 7 * L(4) + 2 * P(3) + 2 * P(4))

    dl = dl + _
      0.00028 * Sin(L(4) - P(4) + 2 * Psi - 2 * BigPi) _
      - 0.00028 * Sin(2 * (L(4) - Om(4))) _
      - 0.00027 * Sin(P(3) - P(4) + Om(3) - Om(4)) _
      - 0.00026 * Sin(5 * Gd - 3 * g + 188.37 * DToR) _
      + 0.00025 * Sin(Om(4) - Om(3)) _
      - 0.00025 * Sin(L(2) - 3 * L(3) + 2 * L(4))

    dl = dl + _
      -0.00023 * Sin(3 * (L(3) - L(4))) _
      + 0.00021 * Sin(2 * L(4) - 2 * BigPi - 3 * g) _
      - 0.00021 * Sin(2 * L(3) - 3 * L(4) + P(4)) _
      + 0.00019 * Sin(L(4) - P(4) - g) _
      - 0.00019 * Sin(2 * L(4) - P(3) - P(4)) _
      - 0.00018 * Sin(L(4) - P(4) + g) _
      - 0.00016 * Sin(L(4) + P(3) - 2 * BigPi - 2 * g)
    dl4 = dl * DToR
End Function

Private Function b4() As Double
Dim tanb As Double
'    tanb = _
      -0.0076579 * Sin(L(4) - Psi) _
      + 0.0044148 * Sin(L(4) - Om(4)) _
      - 0.0005106 * Sin(L(4) - Om(3)) _
      + 0.0000773 * Sin(L(4) + Psi - 2 * BigPi - 2 * g)
'    tanb = tanb _
      + 0.0000104 * Sin(L(4) - Psi + g) _
      - 0.0000102 * Sin(L(4) - Psi - g) _
      + 0.0000088 * Sin(L(4) + Psi - 2 * BigPi - 3 * g) _
      - 0.0000038 * Sin(L(4) + Psi - 2 * BigPi - g)
          tanb = _
      -0.0076579 * Sin(L(4) - Psi) _
      + 0.0044134 * Sin(L(4) - Om(4)) _
      - 0.0005112 * Sin(L(4) - Om(3)) _
      + 0.0000773 * Sin(L(4) + Psi - 2 * BigPi - 2 * g)

    tanb = tanb + _
      0.0000104 * Sin(L(4) - Psi + g) _
      - 0.0000102 * Sin(L(4) - Psi - g) _
      + 0.0000088 * Sin(L(4) + Psi - 2 * BigPi - 3 * g) _
      - 0.0000038 * Sin(L(4) + Psi - 2 * BigPi - g)
    b4 = Atn(tanb)
End Function

Private Function r4() As Double
Dim dr As Double
'     dr = _
      -0.0073391 * Cos(L(4) - P(4)) _
      + 0.000162 * Cos(L(4) - P(3)) _
      + 0.0000974 * Cos(L(3) - L(4)) _
      - 0.0000541 * Cos(L(4) + P(4) - 2 * BigPi - 2 * g) _
      - 0.0000269 * Cos(2 * (L(4) - P(4)))
'    dr = dr + _
     0.0000182 * Cos(L(4) - BigPi) _
     + 0.0000177 * Cos(2 * (L(3) - L(4))) _
     - 0.0000167 * Cos(2 * L(4) - Psi - Om(4)) _
     + 0.0000167 * Cos(Psi - Om(4)) _
     - 0.0000155 * Cos(2 * (L(4) - BigPi - g)) _
     + 0.0000142 * Cos(2 * (L(4) - Psi))
'    dr = dr + _
      0.0000104 * Cos(L(1) - L(4)) _
      + 0.0000092 * Cos(L(2) - L(4)) _
      - 0.0000089 * Cos(L(4) - BigPi - g) _
      - 0.0000062 * Cos(L(4) + P(4) - 2 * BigPi - 3 * g) _
      + 0.0000048 * Cos(2 * (L(4) - Om(4)))
'    r4 = 26.3699 * (1 + dr)
    dr = _
      -0.0073546 * Cos(L(4) - P(4)) _
      + 0.0001621 * Cos(L(4) - P(3)) _
      + 0.0000974 * Cos(L(3) - L(4)) _
      - 0.0000543 * Cos(L(4) + P(4) - 2 * BigPi - 2 * g) _
      - 0.0000271 * Cos(2 * (L(4) - P(4)))

    dr = dr + _
      0.0000182 * Cos(L(4) - BigPi) _
      + 0.0000177 * Cos(2 * (L(3) - L(4))) _
      - 0.0000167 * Cos(2 * L(4) - Psi - Om(4)) _
      + 0.0000167 * Cos(Psi - Om(4)) _
      - 0.0000155 * Cos(2 * (L(4) - BigPi - g)) _
      + 0.0000142 * Cos(2 * (L(4) - Psi))

    dr = dr + _
      0.0000105 * Cos(L(1) - L(4)) _
      + 0.0000092 * Cos(L(2) - L(4)) _
      - 0.0000089 * Cos(L(4) - BigPi - g) _
      - 0.0000062 * Cos(L(4) + P(4) - 2 * BigPi - 3 * g) _
      + 0.0000048 * Cos(2 * (L(4) - Om(4)))

    r4 = 26.36273 * (1 + dr)
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

Dim Psi As Double, Arg As Double, Omega As Double, c As Double, s As Double, r   As Double
Dim vsSatellite As TSVECTOR
  Psi = (316.50043 - 0.075972 * T) * DToR
  If (n = DUMMY_SATELLITE) Then
    v.x = 0
    v.y = 0
    v.Z = 1
    r = 1
  Else
    Select Case n
        Case 1: Call IO(T, vsSatellite)
        Case 2: Call Europa(T, vsSatellite)
        Case 3: Call Ganymede(T, vsSatellite)
        Case 4: Call Callisto(T, vsSatellite)
    End Select
    vsSatellite.L = vsSatellite.L - Psi
    r = vsSatellite.r
    Call SphToRect(vsSatellite, v)
  End If

  Arg = (3.120262 + 0.0006 * (TToJD(T) - 2415020#) / 36525) * DToR
  c = Cos(Arg)
  s = Sin(Arg)
  Call XRot(v, c, s, v)

  Omega = (100.464407 + T * (1.02097745 + T * (0.000403157 + T * 0.000000404))) * DToR
  '{ General precession }
  Arg = (T - TB1950)
  Arg = Arg * (1.3966626 + Arg * 0.0003088) * DToR
  Arg = Psi + Arg - Omega
  c = Cos(Arg)
  s = Sin(Arg)
  Call ZRot(v, c, s, v)

  Arg = (1.303267 + T * (-0.0054965 + T * (0.00000466 - T * 0.000000002))) * DToR
  c = Cos(Arg)
  s = Sin(Arg)
  Call XRot(v, c, s, v)

  Arg = Omega
  c = Cos(Arg)
  s = Sin(Arg)
  Call ZRot(v, c, s, v)
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
          ByRef vDummy As TVECTOR, bScale As Boolean, ByRef W As TVECTOR, bPerspective As Boolean)

Dim Arg As Double, c As Double, s As Double, r As Double, rDummy  As Double
Dim k As Variant ' K(4) As Double

  k = Array(0, 17295#, 21819#, 27558#, 36548#)
  Arg = Pi / 2 - vsOrigin.L
  c = Cos(Arg)
  s = Sin(Arg)
  Call ZRot(v, c, s, W)

  Arg = W.Z
  W.Z = W.y
  W.y = Arg

  r = Sqr(v.x * v.x + v.y * v.y + v.Z * v.Z)
  Arg = vsOrigin.B
  c = Cos(Arg)
  s = Sin(Arg)
  Call XRot(W, c, s, W)

  If (n <> DUMMY_SATELLITE) Then
    '{ 'Rectification' }
    rDummy = Sqr(vDummy.x * vDummy.x + vDummy.y * vDummy.y)
    c = vDummy.y / rDummy
    s = vDummy.x / rDummy
    Call ZRot(W, c, s, W)

    '{ Scaling }
    If bScale Then W.y = W.y * 1.071374

    '{ Light time correction }
    Arg = W.x / r
    W.x = W.x + Abs(W.Z) / k(n) * Sqr(1 - Arg * Arg)

    '{ Perspective effect }
    If bPerspective Then
        Arg = vsOrigin.r / (vsOrigin.r + W.Z / 2095)
        W.x = W.x * Arg
        W.y = W.y * Arg
    End If
  End If
End Sub

