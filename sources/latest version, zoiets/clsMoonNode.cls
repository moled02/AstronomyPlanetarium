VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMoonNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'(*****************************************************************************)
'(* Module: MOONNODE.PAS                                                      *)
'(* Version 2.0                                                               *)
'(* Last modified: June 6, 1996                                               *)
'(*****************************************************************************)

Private MeanMoonNodeData As Variant

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

Function FirstkMoonNodeOfYear(Year As Long) As Double
FirstkMoonNodeOfYear = Int((Year - 2000.05) * 13.4223)
End Function

'(*****************************************************************************)
'(* Name:   FindMeanValues                                                    *)
'(* Type:   Support function                                                  *)
'(* Purpose:Calculate the mean values, JDE0 and M corresponding to a          *)
'(*         a given event and k value.                                        *)
'(*****************************************************************************)

Sub FindMoonNodeValues(ByVal k As Double, ByRef D As Double, ByRef m As Double, ByRef M1 As Double, ByRef Om As Double, ByRef v As Double, ByRef P As Double, ByRef E As Double)
Dim T As Double
Dim Geg(7) As Double
Dim i As Long
T = k / 1342.23
For i = 1 To 4
    Geg(i) = MeanMoonNodeData(i)(0)
    Geg(i) = Geg(i) + MeanMoonNodeData(i)(1) * k
    Geg(i) = Geg(i) + T * T * (MeanMoonNodeData(i)(2) _
                           + T * (MeanMoonNodeData(i)(3) _
                                 + T * MeanMoonNodeData(i)(4)))
Next
For i = 5 To 7
    Geg(i) = MeanMoonNodeData(i)(0)
    Geg(i) = Geg(i) + T * MeanMoonNodeData(i)(1)
    Geg(i) = Geg(i) + T * T * (MeanMoonNodeData(i)(2) _
                           + T * (MeanMoonNodeData(i)(3) _
                                 + T * MeanMoonNodeData(i)(4)))
Next
D = modpi2(Geg(1) * DToR)
m = modpi2(Geg(2) * DToR)
M1 = modpi2(Geg(3) * DToR)
Om = modpi2(Geg(4) * DToR)
v = modpi2(Geg(5) * DToR)
P = modpi2(Geg(6) * DToR + Om)
E = Geg(7)
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

Sub moonnode(ByVal k As Double, ByRef JDE As Double)

Dim D As Double, m As Double, M1 As Double, Om As Double, v As Double, P As Double, E As Double, T As Double, sum, A As Double
Dim i As Long, j As Long
Dim preset As Long
Dim hAngle As Double, DeltaJDE As Double, DeltaDec As Double

T = k / 1342.23
Call FindMoonNodeValues(k, D, m, M1, Om, v, P, E)

JDE = 2451565.1619 + 27.212220817 * k + 0.0002572 * T * T _
                   + 0.000000021 * T * T * T _
                   - 0.000000000088 * T * T * T * T _
                   - 0.4721 * Sin(M1) _
                   - 0.1649 * Sin(2 * D) _
                   - 0.0868 * Sin(2 * D - M1) _
                   + 0.0084 * Sin(2 * D + M1) _
                   - 0.0083 * Sin(2 * D - m) * E _
                   - 0.0039 * Sin(2 * D - m - M1) * E _
                   + 0.0034 * Sin(2 * M1) _
                   - 0.0031 * Sin(2 * D - 2 * M1) _
                   + 0.003 * Sin(2 * D + m) * E _
                   + 0.0028 * Sin(m - M1) * E _
                   + 0.0026 * Sin(m) * E _
                   + 0.0025 * Sin(4 * D) _
                   + 0.0024 * Sin(D) _
                   + 0.0022 * Sin(m + M1) * E _
                   + 0.0017 * Sin(Om) _
                   + 0.0014 * Sin(4 * D - M1) _
                   + 0.0005 * Sin(2 * D + m - M1) * E _
                   + 0.0004 * Sin(2 * D - m + M1) * E _
                   - 0.0003 * Sin(2 * D - 2 * m) * E _
                   + 0.0003 * Sin(4 * D - m) _
                   + 0.0003 * Sin(v) _
                   + 0.0003 * Sin(P)
End Sub

Private Sub Class_Initialize()
  '{ Table 49 }
'  MeanMoonNodeData : array(1..7, 0..4) of real = (
  MeanMoonNodeData = Array(Array(0, 0, 0, 0, 0), _
    Array(183.638, 331.73735691, 0.0015057, 0.00000209, -0.00000001), _
    Array(17.4006, 26.8203725, 0.0000999, 0.00000006, 0), _
    Array(38.3776, 355.52747322, 0.0123577, 0.000014628, -0.000000069), _
    Array(123.9767, -1.44098949, 0.0020625, 0.00000214, -0.000000016), _
    Array(299.75, 132.85, -0.009173, 0, 0), _
    Array(272.75, -2.3, 0, 0, 0), _
    Array(1, -0.002516, 0.0000074, 0, 0))
End Sub
