Attribute VB_Name = "modMoonPhys"
'(*****************************************************************************)
'(* Name:    PhysLibrationParams                                              *)
'(* Type:    Support procedure                                                *)
'(* Purpose: Calculate the physical libration parameters rho, sigma and tau.  *)
'(* Arguments:                                                                *)
'(*   T : number of centuries since J2000                                     *)
'(*   F : Moon's mean elongation                                              *)
'(*   Om : longitude of the Moon's ascending node                             *)
'(*   Rho, Sigma, Tau : the required values                                   *)
'(*****************************************************************************)

Sub PhysLibrationParams(t As Double, Om As Double, D As Double, Md As Double, M As Double, F As Double, E As Double, ByRef Rho As Double, ByRef Sigma As Double, ByRef Tau As Double)
Dim k1 As Double, k2 As Double
k1 = (119.75 + 131.849 * t) * DToR
k2 = (72.56 + 20.186 * t) * DToR

Rho = _
  -0.02752 * Cos(Md) _
  - 0.02245 * Sin(F) _
  + 0.00684 * Cos(Md - 2 * F) _
  - 0.00293 * Cos(2 * F) _
  - 0.00085 * Cos(2 * F - 2 * D) _
  - 0.00054 * Cos(Md - 2 * D) _
  - 0.0002 * Sin(Md + F) _
  - 0.0002 * Cos(Md + 2 * F) _
  - 0.0002 * Cos(Md - F) _
  + 0.00014 * Cos(Md + 2 * F - 2 * D)
Rho = Rho * DToR

Sigma = _
  -0.02816 * Sin(Md) _
  + 0.02244 * Cos(F) _
  - 0.00682 * Sin(Md - 2 * F) _
  - 0.00279 * Sin(2 * F) _
  - 0.00083 * Sin(2 * F - 2 * D) _
  + 0.00069 * Sin(Md - 2 * D) _
  + 0.0004 * Cos(Md + F) _
  - 0.00025 * Sin(2 * Md) _
  - 0.00023 * Sin(Md + 2 * F) _
  + 0.0002 * Cos(Md - F) _
  + 0.00019 * Sin(Md - F) _
  + 0.00013 * Sin(Md + 2 * F - 2 * D) _
  - 0.0001 * Cos(Md - 3 * F)
Sigma = Sigma * DToR

Tau = _
  0.0252 * E * Sin(M) _
  + 0.00473 * Sin(2 * Md - 2 * F) _
  - 0.00467 * Sin(Md) _
  + 0.00396 * Sin(k1) _
  + 0.00276 * Sin(2 * Md - 2 * D) _
  + 0.00196 * Sin(Om) _
  - 0.00183 * Cos(Md - F) _
  + 0.00115 * Sin(Md - 2 * D) _
  - 0.00096 * Sin(Md - D) _
  + 0.00046 * Sin(2 * F - 2 * D) _
  - 0.00039 * Sin(Md - F) _
  - 0.00032 * Sin(Md - M - 2 * D) _
  + 0.00027 * Sin(2 * Md - M - 2 * D) _
  + 0.00023 * Sin(k2) _
  - 0.00014 * Sin(2 * D) _
  + 0.00014 * Cos(2 * Md - 2 * F) _
  - 0.00012 * Sin(Md - 2 * F) _
  - 0.00012 * Sin(2 * Md) _
  + 0.00011 * Sin(2 * Md - 2 * M - 2 * D)
Tau = Tau * DToR
End Sub

'(*****************************************************************************)
'(* Name:    MoonPhysEphemeris                                                *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Calculate a physical ephemeris of the Moon.                      *)
'(* Arguments:                                                                *)
'(*   T : number of centuries since J2000                                     *)
'(*   l, b : the Moon'z ecliptical coordinates                                *)
'(*   MoonPhysData : TMOONPHYSDATA record to hold the results                 *)
'(*****************************************************************************)

Sub MoonPhysEphemeris(ByVal t As Double, SGeo As TSVECTOR, SSun As TSVECTOR, ByVal Obl As Double, _
                ByVal NutLon As Double, ByVal NutObl As Double, _
                ByRef MoonPhysData As TMOONPHYSDATA)
    
Dim Om As Double, Om1 As Double, D As Double, M As Double, Md As Double, F As Double, E As Double, a As Double
Dim Rho As Double, Tau As Double, Sigma As Double
Dim RA As Double, Decl As Double, v As Double, x As Double, Y As Double
Dim Psi As Double, RA0 As Double, Decl0 As Double
Dim lH As Double, Bh As Double, l01 As Double, b01 As Double, l011 As Double, b011 As Double

Const i = (1.54242 * DToR)

'{ Calculate Optical libration and auxiliary variables F, Om and A }
Call CalcMoonAngles(t, Om, D, M, Md, F, E)
Call MoonLib_l1b1(SGeo.l, SGeo.B, Om, F, i, MoonPhysData.ld, MoonPhysData.Bd, a)

'{ Calculate Physical libration }
Call PhysLibrationParams(t, Om, D, Md, M, F, E, Rho, Sigma, Tau)
Call MoonLib_l11b11(MoonPhysData.Bd, Tau, Rho, Sigma, a, MoonPhysData.ldd, MoonPhysData.bdd)
Call MoonLib_lb(MoonPhysData.ld, MoonPhysData.Bd, _
           MoonPhysData.ldd, MoonPhysData.bdd, _
           MoonPhysData.l, MoonPhysData.B)

'{ Position angle }
Obl = Obl + NutObl
SGeo.l = SGeo.l + NutLon
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
v = Om + NutLon + Sigma / Sin(i)
x = Sin(i + Rho) * Sin(v)
Y = Sin(i + Rho) * Cos(v) * Cos(Obl) - Cos(i + Rho) * Sin(Obl)
Om1 = atan2(x, Y)
MoonPhysData.P = asin(Sqr(x * x + Y * Y) * Cos(RA - Om1) / Cos(SGeo.B))

'{ Phase & position angle of bright limb }
Psi = acos(Cos(SGeo.B) * Cos(SGeo.l - SSun.l))
MoonPhysData.i = Atn(SSun.r * Sin(Psi) / (SGeo.r / 149597870 - SSun.r * Cos(Psi)))
If MoonPhysData.i < 0 Then MoonPhysData.i = MoonPhysData.i + Pi
MoonPhysData.k = (1 + Cos(MoonPhysData.i)) / 2
Call EclToEqu(SSun.l, SSun.B, Obl, RA0, Decl0)
Y = Cos(Decl0) * Sin(RA0 - RA)
x = Sin(Decl0) * Cos(Decl) - Cos(Decl0) * Sin(Decl) * Cos(RA0 - RA)
MoonPhysData.x = modpi2(atan2(Y, x))

'{Terminator angle}
lH = SSun.l + Pi + SGeo.r / (149597870 * SSun.r) * 57.296 * Cos(SGeo.B) * Sin(SSun.l - SGeo.l) * DToR
Bh = SGeo.r / (149597870 * SSun.r) * SGeo.B

Call MoonLib_l1b1(lH, Bh, Om, F, i, l01, b01, a)
Call MoonLib_l11b11(MoonPhysData.Bd, Tau, Rho, Sigma, a, l011, b011)
Call MoonLib_lb(l01, b01, l011, b011, MoonPhysData.l0, MoonPhysData.b0)
If MoonPhysData.l0 < 0 Then MoonPhysData.l0 = MoonPhysData.l0 + Pi2

MoonPhysData.t = modpi(Pi * 0.5 - MoonPhysData.l0)
If MoonPhysData.t > Pi * 0.5 Then MoonPhysData.t = -(Pi - MoonPhysData.t)
If MoonPhysData.t < -Pi * 0.5 Then MoonPhysData.t = Pi + MoonPhysData.t
End Sub


Private Sub MoonLib_l1b1(l As Double, B As Double, Om As Double, F As Double, i As Double, ByRef l1 As Double, ByRef b1 As Double, ByRef a As Double)
Dim W As Double
      W = l - Om
      a = atan2(Sin(W) * Cos(B) * Cos(i) - Sin(B) * Sin(i), Cos(W) * Cos(B))
      l1 = modpi(a - F)
      b1 = asin(-Sin(W) * Cos(B) * Sin(i) - Sin(B) * Cos(i))
End Sub

Private Sub MoonLib_l11b11(b1 As Double, Tau As Double, Rho As Double, Sigma As Double, a As Double, ByRef l11 As Double, ByRef b11 As Double)
      l11 = -Tau + (Rho * Cos(a) + Sigma * Sin(a)) * tan(b1)
      b11 = Sigma * Cos(a) - Rho * Sin(a)
End Sub

Private Sub MoonLib_lb(l1 As Double, b1 As Double, l11 As Double, b11 As Double, ByRef l As Double, ByRef B As Double)
      l = l1 + l11
      B = b1 + b11
End Sub




'(*****************************************************************************)
'(* Name:    MoonSemiDiameter                                                 *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Calculate a physical ephemeris of the Moon.                      *)
'(* Arguments:                                                                *)
'(*   DistEarth : Distance of the Moon to the Earth in km.                    *)
'(*****************************************************************************)

Function MoonSemiDiameter(DistEarth As Double) As Double
MoonSemiDiameter = 358473400 / DistEarth
End Function
