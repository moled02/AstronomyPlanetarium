Attribute VB_Name = "modMoonAngles"

'(*****************************************************************************)
'(* Name:    EvalMoonAngles                                                   *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Calculate the angles Om, D, M, Md and F, as well as E.           *)
'(* Arguments:                                                                *)
'(*   T : number of Julian centuries since J2000                              *)
'(*   Om, D, M, Md, F, E : the results                                        *)
'(*                                                                           *)
'(* Note: This procedure is also used by MoonPhysEphemeris and Eclipse        *)
'(*****************************************************************************)

Sub CalcMoonAngles(T As Double, ByRef Om As Double, ByRef D As Double, ByRef M As Double, ByRef Md As Double, ByRef F As Double, ByRef E As Double)
Dim POm As Variant
Dim PD  As Variant
Dim PM  As Variant
Dim PMd As Variant
Dim PF  As Variant
POm = Array(125.044555, -1934.1361849, 0.0020762, 0.000002139449, 0.0000000164973)
PD = Array(297.8502042, 445267.1115168, -0.00163, 0.000001831945, 0.00000884447)
PM = Array(357.5291092, 35999.0502909, -0.0001536, 0.00000004083299, 0)
PMd = Array(134.9634114, 477198.8676313, 0.008997, 0.00001434741, 0.0000000679717)
PF = Array(93.2720993, 483202.0175273, -0.0034029, 0.0000002836075, 0.00000000115833)

Om = modpi2(Eval4Poly(C4P(POm), T) * DToR)
D = modpi2(Eval4Poly(C4P(PD), T) * DToR)
M = modpi2(Eval4Poly(C4P(PM), T) * DToR)
Md = modpi2(Eval4Poly(C4P(PMd), T) * DToR)
F = modpi2(Eval4Poly(C4P(PF), T) * DToR)
E = 1 - T * (0.002516 + T * 0.0000074)
End Sub

