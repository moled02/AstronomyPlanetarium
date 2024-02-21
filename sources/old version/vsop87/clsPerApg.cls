VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPerApg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'(*****************************************************************************)
'(* Module: PERAPG.PAS                                                        *)
'(* Version 2.0                                                               *)
'(* Last modified: June 3, 1996                                               *)
'(*****************************************************************************)

Private MeanPerApgData As Variant
Private CPerTab48A  As Variant
Private APerTab48A As Variant
Private CApgTab48A As Variant
Private AApgTab48A As Variant
Private CPerTab48B As Variant
Private CApgTab48B As Variant
Private APerTab48B As Variant
Private AApgTab48B As Variant


'(*****************************************************************************)
'(* Name:    FirstkOfYear                                                     *)
'(* Type:    Function                                                         *)
'(* Purpose: Approximate 'k' corresponding to the beginning of a year.        *)
'(* Arguments:                                                                *)
'(*   Year : the year                                                         *)
'(* Return value:                                                             *)
'(*   the 'k' value corresponding to approximately the first occurance of     *)
'(*   event in Year                                                           *)
'(*****************************************************************************)

Public Function FirstkPerApgOfYear(Year As Long) As Double
FirstkPerApgOfYear = Int((Year - 1999.97) * 13.2555)
End Function

'(*****************************************************************************)
'(* Name:   FindMeanValues                                                    *)
'(* Type:   Support function                                                  *)
'(* Purpose:Calculate the mean values, JDE0 and M corresponding to a          *)
'(*         a given event and k value.                                        *)
'(*****************************************************************************)

Public Sub FindPerApgValues(ByVal k As Double, ByRef JDE As Double, ByRef D As Double, ByRef m As Double, ByRef f As Double)

Dim T As Double
Dim Geg(4) As Double
Dim i As Long
T = k / 1325.55
For i = 1 To 4
    Geg(i) = MeanPerApgData(i)(0)
    Geg(i) = Geg(i) + MeanPerApgData(i)(1) * k
    Geg(i) = Geg(i) + T * T * (MeanPerApgData(i)(2) + _
                             T * (MeanPerApgData(i)(3) + _
                             T * MeanPerApgData(i)(4)))
Next
D = modpi2(Geg(1) * DToR)
f = modpi2(Geg(3) * DToR)
m = modpi2(Geg(2) * DToR)
JDE = Geg(4)
End Sub

'(*****************************************************************************)
'(* Name:    NextPerApg                                                       *)
'(* Type:    Function                                                         *)
'(*   k : index from FirstkOfYear                                             *)
'(* Return value:                                                             *)
'(*   Julian day of instance when the Event occurs                            *)
'(*****************************************************************************)

Public Sub perApg(ByVal k As Double, ByRef JDE As Double, ByRef Par As Double)

Dim D As Double, m As Double, f As Double, T As Double, sum As Double, A As Double
Dim i As Long, j As Long
Dim bPerigeum As Boolean
Dim hAngle   As Double, DeltaJDE As Double, DeltaPar As Double

If Abs(k - Int(k)) < 0.01 Then
  bPerigeum = True
Else
  bPerigeum = False
End If
Call FindPerApgValues(k, JDE, D, m, f)
T = k / 1325.55
If bPerigeum Then
    For i = 0 To 59
        hAngle = APerTab48A(i)(0) * D
        hAngle = hAngle + APerTab48A(i)(1) * f
        hAngle = hAngle + APerTab48A(i)(2) * m
        DeltaJDE = Sin(hAngle) * (CPerTab48A(i)(0) + T * (CPerTab48A(i)(1)))
        JDE = JDE + DeltaJDE
    Next
    Par = 0
    For i = 0 To 31
        hAngle = APerTab48B(i)(0) * D
        hAngle = hAngle + APerTab48B(i)(1) * f
        hAngle = hAngle + APerTab48B(i)(2) * m
        DeltaPar = Cos(hAngle) * (CPerTab48B(i)(0) + T * (CPerTab48B(i)(1)))
        Par = Par + DeltaPar
    Next
Else
    DeltaJDE = 0
    For i = 0 To 31
'        If i Mod 5 = 0 Then Debug.Print "========================================="
'        Debug.Print AApgTab48A(i)(0), AApgTab48A(i)(1), AApgTab48A(i)(2), CApgTab48A(i)(0), CApgTab48A(i)(1)
        hAngle = AApgTab48A(i)(0) * D
        hAngle = hAngle + AApgTab48A(i)(1) * f
        hAngle = hAngle + AApgTab48A(i)(2) * m
        DeltaJDE = DeltaJDE + Sin(hAngle) * (CApgTab48A(i)(0) + T * (CApgTab48A(i)(1)))
    Next
    JDE = JDE + DeltaJDE
    Par = 0
    For i = 0 To 17
        hAngle = AApgTab48B(i)(0) * D
        hAngle = hAngle + AApgTab48B(i)(1) * f
        hAngle = hAngle + AApgTab48B(i)(2) * m
        DeltaPar = Cos(hAngle) * (CApgTab48B(i)(0) + T * (CApgTab48B(i)(1)))
        Par = Par + DeltaPar
    Next
End If
End Sub



Private Sub Class_Initialize()
  '{ Table 48.A, 48.B }
'  MeanPerApgData : array[1..4, 0..4] of real = (
  MeanPerApgData = Array(Array(0, 0, 0, 0, 0), _
    Array(171.9179, 335.9106046, -0.010025, -0.00001156, 0.000000055), _
    Array(347.3477, 27.1577721, -0.0008323, -0.000001, 0), _
    Array(316.6109, 364.5287911, -0.0125131, -0.0000148, 0), _
    Array(2451534.6698, 27.55454988, -0.0006886, -0.000001098, 0.0000000052))
'  CPerTab48A : array[0..59, 0..1] of real = (
  CPerTab48A = Array(Array(-1.6769, 0), Array(0.4589, 0), Array(-0.1856, 0), Array(0.0883, 0), _
    Array(-0.0773, 0.00019), Array(0.0502, -0.00013), Array(-0.046, 0), Array(0.0422, -0.00011), _
    Array(-0.0256, 0), Array(0.0253, 0), Array(0.0237, 0), Array(0.0162, 0), _
    Array(-0.0145, 0), Array(0.0129, 0), Array(-0.0112, 0), Array(-0.0104, 0), _
    Array(0.0086, 0), Array(0.0069, 0), Array(0.0066, 0), Array(-0.0053, 0), _
    Array(-0.0052, 0), Array(-0.0046, 0), Array(-0.0041, 0), Array(0.004, 0), _
    Array(0.0032, 0), Array(-0.0032, 0), Array(0.0031, 0), Array(-0.0029, 0), _
    Array(0.0027, 0), Array(0.0027, 0), Array(-0.0027, 0), Array(0.0024, 0), _
    Array(-0.0021, 0), Array(-0.0021, 0), Array(-0.0021, 0), Array(0.0019, 0), _
    Array(-0.0018, 0), Array(-0.0014, 0), Array(-0.0014, 0), Array(-0.0014, 0), _
    Array(0.0014, 0), Array(-0.0014, 0), Array(0.0013, 0), Array(0.0013, 0), _
    Array(0.0011, 0), Array(-0.0011, 0), Array(-0.001, 0), Array(-0.0009, 0), _
    Array(-0.0008, 0), Array(0.0008, 0), Array(0.0008, 0), Array(0.0007, 0), _
    Array(0.0007, 0), Array(0.0007, 0), Array(-0.0006, 0), Array(-0.0006, 0), _
    Array(0.0006, 0), Array(0.0005, 0), Array(0.0005, 0), Array(-0.0004, 0))
'   APerTab48A : array[0..59, 0..2] of real = (
   APerTab48A = Array(Array(2, 0, 0), Array(4, 0, 0), Array(6, 0, 0), Array(8, 0, 0), _
    Array(2, 0, -1), Array(0, 0, 1), Array(10, 0, 0), Array(4, 0, -1), _
    Array(6, 0, -1), Array(12, 0, 0), Array(1, 0, 0), Array(8, 0, -1), _
    Array(14, 0, 0), Array(0, 0, 2), Array(3, 0, 0), Array(10, 0, -1), _
    Array(16, 0, 0), Array(12, 0, -1), Array(5, 0, 0), Array(2, 2, 0), _
    Array(18, 0, 0), Array(14, 0, -1), Array(7, 0, 0), Array(2, 0, 1), _
    Array(20, 0, 0), Array(1, 0, 1), Array(16, 0, -1), Array(4, 0, 1), _
    Array(9, 0, 0), Array(4, 0, 2), Array(2, 0, -2), Array(4, 0, -2), _
    Array(6, 0, -2), Array(22, 0, 0), Array(18, 0, -1), Array(6, 0, 1), _
    Array(11, 0, 0), Array(8, 0, 1), Array(4, 0, -2), Array(6, 0, 2), _
    Array(3, 0, 1), Array(5, 0, 1), Array(13, 0, 0), Array(20, 0, -1), _
    Array(3, 0, 2), Array(4, 2, -2), Array(1, 0, 2), Array(22, 0, -1), _
    Array(0, 0, 4), Array(6, 0, -2), Array(2, -2, 1), Array(0, 0, 2), _
    Array(0, 2, -1), Array(2, 4, 0), Array(0, 2, -2), Array(2, -2, 2), _
    Array(24, 0, 0), Array(4, -4, 0), Array(2, 0, 2), Array(1, 0, -1))
'   CApgTab48A : array[0..31, 0..1] of real = (
   CApgTab48A = Array(Array(0.4392, 0), Array(0.0684, 0), Array(0.0456, -0.00011), Array(0.0426, -0.00011), _
    Array(0.0212, 0), Array(-0.0189, 0), Array(0.0144, 0), Array(0.0113, 0), _
    Array(0.0047, 0), Array(0.0036, 0), Array(0.0035, 0), Array(0.0034, 0), _
    Array(-0.0034, 0), Array(0.0022, 0), Array(-0.0017, 0), Array(0.0013, 0), _
    Array(0.0011, 0), Array(0.001, 0), Array(0.0009, 0), Array(0.0007, 0), _
    Array(0.0006, 0), Array(0.0005, 0), Array(0.0005, 0), Array(0.0004, 0), _
    Array(0.0004, 0), Array(0.0004, 0), Array(-0.0004, 0), Array(-0.0004, 0), _
    Array(0.0003, 0), Array(0.0003, 0), Array(0.0003, 0), Array(-0.0003, 0))
'   AApgTab48A : array[0..31, 0..2] of real = (
   AApgTab48A = Array(Array(2, 0, 0), Array(4, 0, 0), Array(0, 0, 1), Array(2, 0, -1), _
    Array(0, 2, 0), Array(1, 0, 0), Array(6, 0, 0), Array(4, 0, -1), _
    Array(2, 2, 0), Array(1, 0, 1), Array(8, 0, 0), Array(6, 0, -1), _
    Array(2, -2, 0), Array(2, 0, -2), Array(3, 0, 0), Array(4, 2, 0), _
    Array(8, 0, -1), Array(4, 0, -2), Array(10, 0, 0), Array(3, 0, 1), _
    Array(0, 0, 2), Array(2, 0, 1), Array(2, 0, 2), Array(6, 2, 0), _
    Array(6, 0, -2), Array(10, 0, -1), Array(5, 0, 0), Array(4, -2, 0), _
    Array(0, 2, 1), Array(12, 0, 0), Array(2, 2, -1), Array(1, 0, -1))
'   CPerTab48B : array[0..46, 0..1] of real = (
   CPerTab48B = Array(Array(3629.215, 0), Array(63.224, 0), Array(-6.99, 0), Array(2.834, -0.0071), _
   Array(1.927, 0), Array(-1.263, 0), Array(-0.702, 0), Array(0.696, -0.0017), _
   Array(-0.69, 0), Array(-0.629, 0.0016), Array(-0.392, 0), Array(0.297, 0), _
   Array(0.26, 0), Array(0.201, 0), Array(-0.161, 0), Array(0.157, 0), _
   Array(-0.138, 0), Array(-0.127, 0), Array(0.104, 0), Array(0.104, 0), _
   Array(-0.079, 0), Array(0.068, 0), Array(0.067, 0), Array(0.054, 0), _
   Array(-0.038, 0), Array(-0.038, 0), Array(0.037, 0), Array(-0.037, 0), _
   Array(-0.035, 0), Array(-0.03, 0), Array(0.029, 0), Array(-0.025, 0), _
   Array(0.023, 0), Array(0.023, 0), Array(-0.023, 0), Array(0.022, 0), Array(-0.021, 0), _
   Array(-0.02, 0), Array(0.019, 0), Array(0.017, 0), Array(0.014, 0), _
   Array(-0.014, 0), Array(0.013, 0), Array(0.012, 0), Array(0.011, 0), _
   Array(0.01, 0), Array(-0.01, 0))
'   CApgTab48B : array[0..17,0..1] of real = (
   CApgTab48B = Array(Array(3245.251, 0), Array(-9.147, 0), Array(-0.841, 0), Array(0.697, 0), _
   Array(-0.656, 0.0016), Array(0.355, 0), Array(0.159, 0), Array(0.127, 0), _
   Array(0.065, 0), Array(0.052, 0), Array(0.043, 0), Array(0.031, 0), _
   Array(-0.023, 0), Array(0.022, 0), Array(0.019, 0), Array(-0.016, 0), _
   Array(0.014, 0), Array(0.01, 0))
'   APerTab48B : array[0..46, 0..2] of real = (
   APerTab48B = Array(Array(0, 0, 0), Array(2, 0, 0), Array(4, 0, 0), Array(2, 0, -1), _
   Array(6, 0, 0), Array(1, 0, 0), Array(8, 0, 0), Array(0, 0, 1), _
   Array(0, 2, 0), Array(4, 0, -1), Array(2, -2, 0), Array(10, 0, 0), _
   Array(6, 0, -1), Array(3, 0, 0), Array(2, 0, 1), Array(1, 0, 1), _
   Array(12, 0, 0), Array(8, 0, -1), Array(2, 2, 0), Array(2, 0, -2), _
   Array(5, 0, 0), Array(14, 0, 0), Array(10, 0, -1), Array(4, 0, 1), _
   Array(12, 0, -1), Array(4, 0, -2), Array(7, 0, 0), Array(4, 2, 0), _
   Array(16, 0, 0), Array(3, 0, 1), Array(1, 0, -1), Array(6, 0, 1), _
   Array(0, 0, 2), Array(14, 0, -1), Array(2, 0, 2), Array(6, 0, -2), _
   Array(2, -2, -1), Array(9, 0, 0), Array(18, 0, 0), Array(6, 2, 0), _
   Array(0, 2, -1), Array(16, 0, -1), Array(4, -2, 0), Array(8, 0, 1), _
   Array(11, 0, 0), Array(4, 0, 1), Array(20, 0, 0))
'   AApgTab48B : array[0..17,0..2] of real = (
   AApgTab48B = Array(Array(0, 0, 0), Array(2, 0, 0), Array(1, 0, 0), Array(0, 2, 0), _
   Array(0, 0, 1), Array(4, 0, 0), Array(2, 0, -1), Array(1, 0, 1), _
   Array(4, 0, -1), Array(6, 0, 0), Array(2, 0, 1), Array(2, 2, 0), _
   Array(2, -2, 0), Array(2, 0, -2), Array(2, 0, 2), Array(0, 0, 2), _
   Array(6, 0, -1), Array(8, 0, 0))
End Sub
