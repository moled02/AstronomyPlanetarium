VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMaxDec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'(*****************************************************************************)
'(* Module: MAXDEC.PAS                                                      *)
'(* Version 2.0                                                               *)
'(* Last modified: October 1, 1992                                            *)
'(*****************************************************************************)

Private MeanMaxDecData As Variant
Private CoeffTab50A  As Variant
Private ArgTab50A As Variant
Private CoeffTab50B As Variant
Private ArgTab50B As Variant
  



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

Function FirstkMaxDecOfYear(Year As Long) As Long
FirstkMaxDecOfYear = Int((Year - 2000.03) * 13.3686)
End Function

'(*****************************************************************************)
'(* Name:   FindMeanValues                                                    *)
'(* Type:   Support function                                                  *)
'(* Purpose:Calculate the mean values, JDE0 and M corresponding to a          *)
'(*         a given event and k value.                                        *)
'(*****************************************************************************)

Sub FindMaxDecValues(ByVal NorthDec As Boolean, ByVal k As Long, ByRef JDE As Double, ByRef D As Double, ByRef M As Double, ByRef M1 As Double, ByRef F As Double, ByRef E As Double, ByRef mDec As Double)
Dim T As Double
Dim Geg(6) As Double
Dim i As Long
T = k / 1336.86
For i = 1 To 5
    If NorthDec Then
      Geg(i) = MeanMaxDecData(i)(0)(0)
    Else
      Geg(i) = MeanMaxDecData(i)(1)(0)
    End If
    Geg(i) = Geg(i) + MeanMaxDecData(i)(0)(1) * k
    Geg(i) = Geg(i) + MeanMaxDecData(i)(0)(2) * T * T
    Geg(i) = Geg(i) + MeanMaxDecData(i)(0)(3) * T * T * T
Next
E = 1 - T * (0.002516 + T * 0.0000074)
mDec = MeanMaxDecData(6)(0)(0) + MeanMaxDecData(6)(0)(1) * T
D = modpi2(Geg(1) * DToR)
M = modpi2(Geg(2) * DToR)
M1 = modpi2(Geg(3) * DToR)
F = modpi2(Geg(4) * DToR)
JDE = Geg(5)
End Sub

'(*****************************************************************************)
'(* Name:    ConjunctionOpposition                                            *)
'(* Type:    Function                                                         *)
'(* Purpose: Instant (JD) of conjunction or opposition of a planet.           *)
'(* Arguments:                                                                *)
'(*   Planet : index to planet (Mercury = 1, Neptune = 8                      *)
'(*   Event : one of OPPOSITION, CONJUNCTION, INFCONJ, SUPCONJ                *)
'(*   k : index from FirstkOfYear                                             *)
'(* Return value:                                                             *)
'(*   Julian day of instance when the Event occurs                            *)
'(*****************************************************************************)

Sub MaxDec(ByVal k As Long, ByVal PosNb As Boolean, ByRef JDE As Double, ByRef mDec As Double)
Dim D   As Double, M As Double, M1 As Double, F As Double, E As Double, T As Double, sum As Double, A As Double
Dim i As Long, j As Long
Dim preset As Long
Dim hAngle   As Double, DeltaJDE As Double, DeltaDec As Double

If PosNb Then
  preset = 0
Else
  preset = 1
End If
Call FindMaxDecValues(PosNb, k, JDE, D, M, M1, F, E, mDec)
For i = 0 To 43
    hAngle = ArgTab50A(i)(0) * D
    hAngle = hAngle + ArgTab50A(i)(1) * M
    hAngle = hAngle + ArgTab50A(i)(2) * M1
    hAngle = hAngle + ArgTab50A(i)(3) * F
    If ArgTab50A(i)(4) = 0 Then
      DeltaJDE = Cos(hAngle) * CoeffTab50A(i)(preset)
    Else
      DeltaJDE = Sin(hAngle) * CoeffTab50A(i)(preset)
    End If
    If ArgTab50A(i)(1) <> 0 Then DeltaJDE = DeltaJDE * E
    JDE = JDE + DeltaJDE
Next
For i = 0 To 36
    hAngle = ArgTab50B(i)(0) * D
    hAngle = hAngle + ArgTab50B(i)(1) * M
    hAngle = hAngle + ArgTab50B(i)(2) * M1
    hAngle = hAngle + ArgTab50B(i)(3) * F
    If ArgTab50B(i)(4) = 0 Then
      DeltaDec = Cos(hAngle) * CoeffTab50B(i)(preset)
    Else
      DeltaDec = Sin(hAngle) * CoeffTab50B(i)(preset)
    End If
    If ArgTab50B(i)(1) <> 0 Then DeltaDec = DeltaDec * E
    mDec = mDec + DeltaDec
Next
mDec = mDec
If Not PosNb Then mDec = -mDec
End Sub

Private Sub Class_Initialize()
  '{ Table 50.A, 50.B }
'  MeanMaxDecData : array(1..6, 0..1, 0..3) of real = (
  MeanMaxDecData = Array(Array(Array(0, 0), Array(0, 0, 0, 0)), _
    Array(Array(152.2029, 333.0705546, -0.0004025, 0.00000011), Array(345.6676, 0, 0, 0)), _
    Array(Array(14.8591, 26.9281592, -0.0000544, -0.0000001), Array(1.3951, 0, 0, 0)), _
    Array(Array(4.6881, 356.9562795, 0.0103126, 0.00001251), Array(186.21, 0, 0, 0)), _
    Array(Array(325.8867, 1.4467806, -0.0020708, -0.00000215), Array(145.1633, 0, 0, 0)), _
    Array(Array(2451562.5897, 27.321582241, 0.000100695, -0.000000141), Array(2451548.9289, 0, 0, 0)), _
    Array(Array(23.6961, -0.013004, 0, 0), Array(0, 0, 0, 0)))
'  CoeffTab50A : array(0..43, 0..1) of real = (
  CoeffTab50A = Array( _
    Array(0.8975, -0.8975), Array(-0.4726, -0.4726), Array(-0.103, -0.103), _
    Array(-0.0976, -0.0976), Array(-0.0462, 0.0541), Array(-0.0461, 0.0516), _
    Array(-0.0438, -0.0438), Array(0.0162, 0.0112), Array(-0.0157, 0.0157), _
    Array(0.0145, 0.0023), Array(0.0136, -0.0136), Array(-0.0095, 0.011), _
    Array(-0.0091, 0.0091), Array(-0.0089, 0.0089), Array(0.0075, 0.0075), _
    Array(-0.0068, -0.003), Array(0.0061, -0.0061), Array(-0.0047, -0.0047), _
    Array(-0.0043, -0.0043), Array(-0.004, 0.004), Array(-0.0037, -0.0037), _
    Array(0.0031, -0.0031), Array(0.003, 0.003), Array(-0.0029, 0.0029), _
    Array(-0.0029, -0.0029), Array(-0.0027, -0.0027), Array(0.0024, 0.0024), _
    Array(-0.0021, -0.0021), Array(0.0019, -0.0019), Array(0.0018, -0.0006), _
    Array(0.0018, -0.0018), Array(0.0017, -0.0017), Array(0.0017, 0.0017), _
    Array(-0.0014, 0.0014), Array(0.0013, -0.0013), Array(0.0013, -0.0013), _
    Array(0.0012, 0.0012), Array(0.0011, 0.0011), Array(-0.0011, 0.0011), _
    Array(0.001, 0.001), Array(0.001, 0.001), Array(-0.0009, -0.0009), _
    Array(0.0007, -0.0007), Array(-0.0007, -0.0007))
  ArgTab50A = Array( _
    Array(0, 0, 0, 1, 0), Array(0, 0, 1, 0, 1), _
    Array(0, 0, 0, 2, 1), Array(2, 0, -1, 0, 1), _
    Array(0, 0, 1, -1, 0), Array(0, 0, 1, 1, 0), _
    Array(2, 0, 0, 0, 1), Array(0, 1, 0, 0, 1), _
    Array(0, 0, 0, 3, 0), Array(0, 0, 1, 2, 1), _
    Array(2, 0, 0, -1, 0), Array(2, 0, -1, -1, 0), _
    Array(2, 0, -1, 1, 0), Array(2, 0, 0, 1, 0), _
    Array(0, 0, 2, 0, 1), Array(0, 0, 1, -2, 1), _
    Array(0, 0, 2, -1, 0), Array(0, 0, 1, 3, 1), _
    Array(2, -1, -1, 0, 1), Array(0, 0, 1, -2, 0), _
    Array(2, 0, -2, 0, 1), Array(0, 0, 0, 1, 1), _
    Array(2, 0, 1, 0, 1), Array(0, 0, 1, 2, 0), _
    Array(2, -1, 0, 0, 1), Array(0, 0, 1, 1, 1), _
    Array(0, 1, -1, 0, 1), Array(0, 0, 1, -3, 1), _
    Array(0, 0, 2, 1, 1), Array(2, 0, -2, -1, 0), _
    Array(0, 0, 0, 3, 1), Array(0, 0, 1, 3, 0), _
    Array(0, 0, 2, 0, 0), Array(2, 0, -1, 0, 0), _
    Array(2, 0, 1, 1, 0), Array(0, 0, 1, 0, 0), _
    Array(0, 0, 3, 1, 1), Array(2, 0, -1, 1, 1), _
    Array(2, 0, -2, 0, 0), Array(1, 0, 0, 1, 0), _
    Array(0, 1, 1, 0, 1), Array(2, 0, 0, -2, 1), _
    Array(0, 0, 2, 1, 0), Array(0, 0, 3, 1, 0))
'  CoeffTab50B : array(0..36, 0..1) of real = (
  CoeffTab50B = Array( _
    Array(5.1093, -5.1093), Array(0.2658, 0.2658), Array(0.1448, -0.1448), _
    Array(-0.0322, 0.0322), Array(0.0133, 0.0133), Array(0.0125, 0.0125), _
    Array(-0.0124, -0.0015), Array(-0.0101, 0.0101), Array(0.0097, -0.0097), _
    Array(-0.0087, 0.0087), Array(0.0074, 0.0074), Array(0.0067, 0.0067), _
    Array(0.0063, -0.0063), Array(0.006, -0.006), Array(-0.0057, 0.0057), _
    Array(-0.0056, -0.0056), Array(0.0052, -0.0052), Array(0.0041, -0.0041), _
    Array(-0.004, -0.004), Array(0.0038, -0.0038), Array(-0.0034, 0.0034), _
    Array(-0.0029, -0.0029), Array(0.0029, 0.0029), Array(-0.0028, 0.0028), _
    Array(-0.0028, -0.0028), Array(-0.0023, 0.0023), Array(-0.0021, 0.0021), _
    Array(0.0019, 0.0019), Array(0.0018, 0.0018), Array(0.0017, -0.0017), _
    Array(0.0015, 0.0015), Array(0.0014, 0.0014), Array(-0.0012, 0.0012), _
    Array(-0.0012, -0.0012), Array(-0.001, 0.001), Array(-0.001, -0.001), _
    Array(0.0006, 0.0037))
'    ArgTab50B : array(0..36, 0..4) of real = (
    ArgTab50B = Array( _
    Array(0, 0, 0, 1, 1), Array(0, 0, 0, 2, 0), _
    Array(2, 0, 0, -1, 1), Array(0, 0, 0, 3, 1), _
    Array(2, 0, 0, -2, 0), Array(2, 0, 0, 0, 0), _
    Array(0, 0, 1, -1, 1), Array(0, 0, 1, 2, 1), _
    Array(0, 0, 0, 1, 0), Array(2, 1, 0, -1, 1), _
    Array(0, 0, 1, 3, 1), Array(1, 0, 0, 1, 1), _
    Array(0, 0, 1, -2, 1), Array(2, -1, 0, -1, 1), _
    Array(2, 0, -1, -1, 1), Array(0, 0, 1, 1, 0), _
    Array(0, 0, 1, 2, 0), Array(0, 0, 2, 1, 0), _
    Array(0, 0, 1, -3, 0), Array(0, 0, 2, -1, 0), _
    Array(0, 0, 1, -2, 0), Array(0, 0, 2, 0, 1), _
    Array(0, 0, 3, 1, 1), Array(2, 1, 0, -1, 0), _
    Array(0, 0, 1, -1, 0), Array(0, 0, 0, 3, 0), _
    Array(2, 0, 0, 1, 1), Array(0, 0, 1, 3, 0), _
    Array(1, 0, 0, 1, 0), Array(0, 0, 2, -1, 1), _
    Array(0, 0, 3, 1, 0), Array(2, 0, 2, 1, 0), _
    Array(2, 0, -2, -1, 1), Array(0, 0, 2, 0, 0), _
    Array(0, 0, 1, 0, 0), Array(0, 0, 0, 2, 1), _
    Array(0, 0, 1, 1, 1))
End Sub
