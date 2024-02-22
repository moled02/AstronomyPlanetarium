Attribute VB_Name = "modEclElmts"
'(*****************************************************************************)
'(* Module: ECLELMTS.PAS                                                      *)
'(* Version 2.0                                                               *)
'(* Last modified: October 1, 1992                                            *)
'(*****************************************************************************)

'(*****************************************************************************)
'(* Name:    ReduceEclElements                                                *)
'(* Type:    sub                                                        *)
'(* Purpose: Reduce ecliptical orbital elements from one epoch to another.    *)
'(* Arguments:                                                                *)
'(*   T0, T1 : number of centuries since J2000 for initial and final epoch    *)
'(*   OrbitEl : TORBITEL record holding the orbital elements                  *)
'(*****************************************************************************)

Sub ReduceEclElements(ByVal T As Double, ByVal T1 As Double, ByRef OrbitEl As TORBITEL)
Dim dT As Double, Eta As Double, BigPi As Double, P As Double, Psi As Double
Dim n As Double, D As Double, i As Double, Om As Double

dT = T1 - T
Eta = dT * (-0.03302 + T * 0.000598 + dT * 0.00006)
Eta = dT * (Eta + (47.0029 + T * (-0.06603 + T * 0.000598))) * SToR
BigPi = dT * (-869.8089 - T * 0.50491 + dT * 0.03536)
BigPi = (BigPi + T * (3289.4789 + T * 0.60622)) * SToR + 174.876384 * DToR
P = dT * (1.11113 - T * 0.000042 - dT * 0.000006)
P = dT * (P + 5029.0966 + T * (2.22226 - T * 0.000042)) * SToR
Psi = BigPi + P
n = Sin(OrbitEl.Incl) * Sin(OrbitEl.LonNode - BigPi)
D = -Sin(Eta) * Cos(OrbitEl.Incl) + Cos(Eta) * Sin(OrbitEl.Incl) * Cos(OrbitEl.LonNode - BigPi)
i = asin(Sqr(n * n + D * D))
Om = atan2(n, D) + Psi

n = Sin(Eta) * Sin(OrbitEl.LonNode - BigPi)
D = Sin(OrbitEl.Incl) * Cos(Eta) - Cos(OrbitEl.Incl) * Sin(Eta) * Cos(OrbitEl.LonNode - BigPi)
OrbitEl.LonPeri = OrbitEl.LonPeri + atan2(n, D)
OrbitEl.Incl = i
OrbitEl.LonNode = modpi2(Om)
End Sub

'(*****************************************************************************)
'(* Name:    QuickReduceEclElements                                           *)
'(* Type:    sub                                                              *)
'(* Purpose: Reduce ecliptical orbital elements from B1950 to J2000.          *)
'(* Arguments:                                                                *)
'(*   OrbitEl : TORBITEL record holding the orbital elements                  *)
'(*****************************************************************************)

Sub QuickReduceEclElements(ByRef OrbitEl As TORBITEL)
Dim W  As Double, A As Double, B As Double
Dim i  As Double, Om As Double, D As Double, n As Double
Const s = 0.0001139788
Const C = 0.9999999935

W = OrbitEl.LonNode - 174.298782 * DToR
A = Sin(OrbitEl.Incl) * Sin(W)
B = C * Sin(OrbitEl.Incl) * Cos(W) - s * Cos(OrbitEl.Incl)
i = asin(Sqr(A * A + B * B))
Om = atan2(A, B) + 174.997194 * DToR
n = -s * Sin(W)
D = C * Sin(OrbitEl.Incl) - s * Cos(OrbitEl.Incl) * Cos(W)
OrbitEl.LonPeri = modpi2(OrbitEl.LonPeri + atan2(n, D))
OrbitEl.Incl = i
OrbitEl.LonNode = Om
End Sub

'(*****************************************************************************)
'(* Name:    FK4FK5ReduceEclElements                                          *)
'(* Type:    sub                                                        *)
'(* Purpose: Reduce ecliptical orbital elements from B1950-FK4 to J2000-FK5.  *)
'(* Arguments:                                                                *)
'(*   OrbitEl : TORBITEL record holding the orbital elements                  *)
'(*****************************************************************************)

Sub FK4FK5ReduceEclElements(ByRef OrbitEl As TORBITEL)
Dim W   As Double, A As Double, B As Double
Dim i  As Double, Om As Double, D As Double, n As Double
Dim sJ As Double, cJ As Double, sW As Double, cW As Double, _
  sI0 As Double, cI0 As Double, sW_W0sI As Double, cW_W0sI As Double, _
  cI  As Double, sLdpOmsI As Double, cLdpOmsI As Double, W_W0 As Double, LdpOm As Double
Const ld = 4.50001688 * DToR
Const l = 5.19856209 * DToR
Const j = 0.00651966 * DToR
  '{ The following are sin J and cos J, respectively }
sJ = Sin(j):           cJ = Cos(j)
W = l + OrbitEl.LonNode
sW = Sin(W):           cW = Cos(W)
sI0 = Sin(OrbitEl.Incl): cI0 = Cos(OrbitEl.Incl)
sW_W0sI = sJ * sW
cW_W0sI = sI0 * cJ + cI0 * sJ * cW
cI = cI0 * cJ - sI0 * sJ * cW
sLdpOmsI = sI0 * sW
cLdpOmsI = cI0 * sJ + sI0 * cJ * cW

W_W0 = atan2(sW_W0sI, cW_W0sI)
W = W_W0 + OrbitEl.LonPeri

LdpOm = atan2(sLdpOmsI, cLdpOmsI)
Om = LdpOm - ld

i = acos(cI)

OrbitEl.Incl = modpi2(i)
OrbitEl.LonNode = modpi2(Om)
OrbitEl.LonPeri = modpi2(W)

'{N = S*sin(W)
D = C * Sin(OrbitEl.Incl) + s * Cos(OrbitEl.Incl) * Cos(W)
A = Sin(OrbitEl.Incl) * Sin(W)
B = s * Cos(OrbitEl.Incl) + C * Sin(OrbitEl.Incl) * Cos(W)

OrbitEl.LonPeri = modpi2(OrbitEl.LonPeri + atan2(n, D))
i = asin(Sqr(A * A + B * B))
Om = atan2(A, B) - ld
If (C * Cos(OrbitEl.Incl) - s * Sin(OrbitEl.Incl) * Cos(W) < 0) Then
  OrbitEl.Incl = PI - i
  OrbitEl.LonNode = modpi2(Om + PI)
Else
  OrbitEl.Incl = i
  OrbitEl.LonNode = Om
End If
End Sub

