Attribute VB_Name = "modKepler"
'(*****************************************************************************)
'(* Module: KEPLER.PAS                                                        *)
'(* Version 2.0                                                               *)
'(* Last modified: October 1, 1992                                            *)
'(*****************************************************************************)

'(*****************************************************************************)
'(* Name:    Kepler                                                           *)
'(* Type:    Function                                                         *)
'(* Purpose: Solve Kepler's equation                                          *)
'(* Arguments:                                                                *)
'(*   M : mean anomaly                                                        *)
'(*   e : orbital eccentricity                                                *)
'(* Return value:                                                             *)
'(*   the corresponding true anomaly                                          *)
'(*****************************************************************************)

Function Kepler(ByVal M As Double, ByVal E As Double) As Double
Dim E0 As Double, DE As Double
Dim A  As Double, B As Double, Z As Double, s0 As Double, s As Double
'{ reduce M to [-pi, pi] }
M = modpi(M)
If (E > 0.97) And (Abs(M) < 0.6) Then
  '{ Danger zone ! Use Mikkola's starting values }
  A = (1 - E) / (4 * E + 0.5)
  B = M / (8 * E + 1)
  Z = Sqr(A * A + B * B)
  If B < 0 Then
    Z = -Exp(Log(-B + Z) / 3)
  Else
    Z = Exp(Log(B + Z) / 3)
  End If
  s0 = Z - A / 2
  s = s0 * (1 - 0.078 * s0 * s0 * s0 * s0 / (1 + E))
  E0 = M + E * s * (3 - 4 * s * s)
Else
  E0 = M
End If
Do
  DE = (M + E * Sin(E0) - E0) / (1 - E * Cos(E0))
  E0 = E0 + DE
  If Abs(DE) < 0.000001 Then
    Exit Do
  End If
Loop
Kepler = 2 * Atn(Sqr((1 + E) / (1 - E)) * tan(E0 / 2))
End Function

'(*****************************************************************************)
'(* Name:   Parabola                                                          *)
'(* Type:   Function                                                          *)
'(* Purpose: Calculate the true anomaly in the case of a parabolic orbit.     *)
'(* Arguments:                                                                *)
'(*   dT : time since perihelion                                              *)
'(*   q : perihelion distance in AU                                           *)
'(* Return value:                                                             *)
'(*   the corresponding true anomaly                                          *)
'(*****************************************************************************)

Function Parabola(ByVal dT As Double, ByVal Q As Double) As Double
Dim W As Double, s As Double, DS As Double
dT = dT * 36525
W = 0.03649116245 * dT / (Q * Sqr(Q))
s = W / 3
Do Until Abs(DS) < 0.000001
  DS = (W - s * (s * s + 3)) / (3 * s * s + 3)
  s = s + DS
Loop
Parabola = 2 * Atn(s)
End Function

'(*****************************************************************************)
'(* Name:   NearParabola                                                      *)
'(* Type:   Function                                                          *)
'(* Purpose: Calculate the true anomaly in the case of a parabolic orbit.     *)
'(* Arguments:                                                                *)
'(*   dT : time since perihelion                                              *)
'(*   q : perihelion distance in AU                                           *)
'(*   e : orbital eccentricity                                                *)
'(* Return value:                                                             *)
'(*   the corresponding true anomaly                                          *)
'(*                                                                           *)
'(* Note:   This function is an implementation of Newton's algorithm for      *)
'(*         formula 34.1.  It is somewhat different from the BASIC program    *)
'(*         on page 230.  It does not, for instance, check for the special    *)
'(*         case of a parabolic orbit.  However, it has a wider convergence   *)
'(*         area of convergence.                                              *)
'(*****************************************************************************)

Function NearParabola(ByVal dT As Double, ByVal Q As Double, ByVal E As Double) As Double
Dim qq   As Double, Gamma As Double, s As Double, s1 As Double, DS As Double, F As Double, SD As Double
Dim nInnerIterations As Long
Dim nIterations As Long

s = tan(Parabola(dT, Q) / 2)
Gamma = (1 - E) / (1 + E)
dT = dT * 36525
qq = GaussConstant * Sqr((1 + E) / Q) / (2 * Q)
qq = qq * dT
nIterations = 1
Do Until (nIterations > 50) Or (Abs(DS) < 0.0001)
  s1 = s
  s = qq
  SD = 0
  F = -s1 * s1
  nInnerIterations = 1
  Do Until (nInnerIterations > 50) Or (Abs(F) < 0.00000001)
    DS = (nInnerIterations - (nInnerIterations + 1) * Gamma)
    s = s + DS * F * s1 / (2 * nInnerIterations + 1)
    SD = SD + DS * F
    F = -F * Gamma * s1 * s1
    nInnerIterations = nInnerIterations + 1
  Loop
  DS = (s - s1) / (SD - 1)
  s = s1 - DS
  nIterations = nIterations + 1
Loop
NearParabola = 2 * Atn(s)
End Function

