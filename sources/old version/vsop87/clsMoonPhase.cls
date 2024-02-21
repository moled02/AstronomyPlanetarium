VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMoonPhase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'(*****************************************************************************)
'(* Module: MOONPHAS.PAS                                                      *)
'(* Version 2.0                                                               *)
'(* Last modified: October 1, 1992                                            *)
'(*****************************************************************************)

Private MoonPhaseData As Variant


'(*****************************************************************************)
'(* Name:    EvalMoonPhaseData                                                *)
'(* Type:    Support function                                                 *)
'(* Purpose: Evaluate the special 'polynomial' used in lunar phase            *)
'(*          calculations                                                     *)
'(* Arguments:                                                                *)
'(*   n : index into MoonPhaseData                                            *)
'(*   k : number of lunation (since J2000)                                    *)
'(*   T : number of centuries since J2000                                     *)
'(* Return value:                                                             *)
'(*   the function value of the polynomial.                                   *)
'(*****************************************************************************)

Function EvalMoonPhaseData(ByVal n As Long, ByVal k As Double, ByVal T As Double) As Double
Dim sum As Double
sum = MoonPhaseData(n)(0) + k * MoonPhaseData(n)(1)
sum = sum + T * T * (MoonPhaseData(n)(2) + T * (MoonPhaseData(n)(3) + T * MoonPhaseData(n)(4)))
EvalMoonPhaseData = sum
End Function

'(*****************************************************************************)
'(* Name:    MoonPhase                                                        *)
'(* Type:    Function                                                         *)
'(* Purpose: Instant of some moon phase of a given lunation.                  *)
'(* Arguments:                                                                *)
'(*   k : number of lunation (since J2000)                                    *)
'(*   Event : one of NEWMOON, FIRSTQUARTER, FULLMOON or LASTQUARTER           *)
'(* Return value:                                                             *)
'(*   the Julian Day at which the event in question occurs                    *)
'(*****************************************************************************)

Function MoonPhase(ByVal k As Double, ByVal Gebeurtenis As Long) As Double
Dim T   As Double, k1 As Double, JDE0 As Double
Dim M   As Double, Md As Double, F As Double, Om As Double
Dim E(2)   As Double
Dim i As Long, j As Long, l As Long, n As Long, flag As Long, sign    As Long
Dim SinTab(2) As TSINCOSTAB, CosTab(2) As TSINCOSTAB
Dim sum   As Double, SinVal As Double, CosVal As Double, Tmp As Double, A As Double, W As Double

Dim MoonPhaseCoeffTab As Variant
Dim MoonPhaseAngleTab As Variant
Dim MoonPhaseExtra As Variant

'  MoonPhaseCoeffTab : array(1..25, 0..2) of real = (
  MoonPhaseCoeffTab = Array(Array(0, 0, 0), Array(-0.4072, -0.40614, -0.62801), _
    Array(0.17241, 0.17302, 0.17172), _
    Array(0.01608, 0.01614, 0.00862), _
    Array(0.01039, 0.01043, 0.00804), _
    Array(0.00739, 0.00734, 0.00454), _
    Array(-0.00514, -0.00515, -0.01183), _
    Array(0.00208, 0.00209, 0.00204), _
    Array(-0.00111, -0.00111, -0.0018), _
    Array(-0.00057, -0.00057, -0.0007), _
    Array(0.00056, 0.00056, 0.00027), _
    Array(-0.00042, 0.00042, -0.0004), _
    Array(0.00042, 0.00042, 0.00032), _
    Array(0.00038, 0.00038, 0.00032), _
    Array(-0.00024, -0.00024, -0.00034), _
    Array(-0.00007, -0.00007, -0.00028), _
    Array(0.00004, 0.00004, 0.00002), _
    Array(0.00004, 0.00004, 0.00003), _
    Array(0.00003, 0.00003, 0.00003), _
    Array(0.00003, 0.00003, 0.00004), _
    Array(-0.00003, -0.00003, -0.00004), _
    Array(0.00003, 0.00003, 0.00002), _
    Array(-0.00002, -0.00002, -0.00005), _
    Array(-0.00002, -0.00002, -0.00002), _
    Array(0.00002, 0.00002, 0), _
    Array(0, 0, 0.00004))

'  MoonPhaseAngleTab : array(1..25,0..3) of integer =
  MoonPhaseAngleTab = Array(Array(0, 0, 0, 0), _
    Array(0, 1, 0, 0), Array(1, 0, 0, 1), Array(0, 2, 0, 0), Array(0, 0, 2, 0), _
    Array(-1, 1, 0, 1), Array(1, 1, 0, 1), Array(2, 0, 0, 2), Array(0, 1, -2, 0), _
    Array(0, 1, 2, 0), Array(1, 2, 0, 1), Array(0, 3, 0, 0), Array(1, 0, 2, 1), _
    Array(1, 0, -2, 1), Array(-1, 2, 0, 1), Array(2, 1, 0, 0), Array(0, 2, -2, 0), _
    Array(3, 0, 0, 0), Array(1, 1, -2, 0), Array(0, 2, 2, 0), Array(1, 1, 2, 0), _
    Array(-1, 1, 2, 1), Array(-1, 1, -2, 0), Array(1, 3, 0, 0), Array(0, 4, 0, 0), _
    Array(-2, 1, 0, 0))

  '{ Coefficients of A1-A14.  Corrections have }
  '{ the form C3 * sin(C1 + C2*T)              }
'  MoonPhaseExtra : array(1..14, 0..2) of real =
  MoonPhaseExtra = Array(Array(0, 0, 0), _
    Array(299.77, 0.107408, 0.000325), Array(251.88, 0.016321, 0.000165), _
    Array(251.83, 26.651886, 0.000164), Array(349.42, 36.412478, 0.000126), _
    Array(84.66, 18.206239, 0.00011), Array(141.74, 53.303771, 0.000062), _
    Array(207.14, 2.453732, 0.000062), Array(154.84, 7.30686, 0.000056), _
    Array(34.52, 27.261239, 0.000047), Array(207.19, 0.121824, 0.000042), _
    Array(291.34, 1.844379, 0.00004), Array(161.72, 24.198154, 0.000037), _
    Array(239.56, 25.513099, 0.000035), Array(331.55, 3.592518, 0.000023))

k1 = k + Gebeurtenis / 4#
T = k1 / 1236.85
JDE0 = EvalMoonPhaseData(0, k1, T)
M = modpi2(EvalMoonPhaseData(1, k1, T) * DToR)
Md = modpi2(EvalMoonPhaseData(2, k1, T) * DToR)
F = modpi2(EvalMoonPhaseData(3, k1, T) * DToR)
Om = modpi2(EvalMoonPhaseData(4, k1, T) * DToR)
Call CalcSinCosTab(M, 3, SinTab(0), CosTab(0))
Call CalcSinCosTab(Md, 4, SinTab(1), CosTab(1))
Call CalcSinCosTab(F, 2, SinTab(2), CosTab(2))

JDE0 = JDE0 - 0.00017 * Sin(Om)
E(1) = 1 - T * (0.002516 + T * 0.0000074)
E(2) = E(1) * E(1)

If Gebeurtenis = NEWMOON Then
  l = 0
ElseIf Gebeurtenis = FULLMOON Then
  l = 1
Else
  l = 2
End If
sum = 0

For i = 1 To 25
  flag = 0
  For j = 0 To 2
    n = MoonPhaseAngleTab(i)(j)
    If (n <> 0) Then
      If (n < 0) Then
        n = -n
        sign = -1
      Else
        sign = 1
      End If
      If (flag = 0) Then
        flag = 1
        SinVal = SinTab(j).W(n) * sign
        CosVal = CosTab(j).W(n)
      Else
        Tmp = CosVal * CosTab(j).W(n) - SinVal * sign * SinTab(j).W(n)
        SinVal = SinVal * CosTab(j).W(n) + CosVal * sign * SinTab(j).W(n)
        CosVal = Tmp
      End If
    End If
  Next
  n = MoonPhaseAngleTab(i)(3)
  If n <> 0 Then SinVal = SinVal * E(n)
  sum = sum + SinVal * MoonPhaseCoeffTab(i)(l)
Next
JDE0 = JDE0 + sum

If (Gebeurtenis = FIRSTQUARTER) Or (Gebeurtenis = LASTQUARTER) Then
  W = 0.00306 - 0.00038 * E(1) * Cos(M) + 0.00026 * Cos(Md)
  W = W - 0.00002 * Cos(Md - M) + 0.00002 * Cos(Md + M) + 0.00002 * Cos(2 * F)
  If Gebeurtenis = FIRSTQUARTER Then
    JDE0 = JDE0 + W
  Else
    JDE0 = JDE0 - W
  End If
End If
For i = 1 To 14
  A = MoonPhaseExtra(i)(0) + k1 * MoonPhaseExtra(i)(1)
  If i = 1 Then A = A - 0.009173 * T * T
  JDE0 = JDE0 + MoonPhaseExtra(i)(2) * Sin(A * DToR)
Next
MoonPhase = JDE0

End Function
Function MoonDistance(ByVal Par As Double) As Double
MoonDistance = 6378.14 / Sin(Par / 3600 * DToR)
End Function

Private Sub Class_Initialize()
'  MoonPhaseData : array(0..4,0..4) of real = (
  MoonPhaseData = Array( _
    Array(2451550.09765, 29.530588853, 0.0001337, 0.00000015, 0.00000000073), _
    Array(2.5534, 29.10535669, -0.0000218, -0.00000011, 0), _
    Array(201.5643, 385.81693528, 0.0107438, 0.00001239, -0.000000058), _
    Array(160.7108, 390.67050274, -0.0016341, -0.00000227, 0.000000011), _
    Array(124.7746, -1.5637558, 0.0020691, 0.00000215, 0))
End Sub
