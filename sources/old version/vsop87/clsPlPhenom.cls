VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlPhenom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'(*****************************************************************************)
'(* Module: PLPHENOM.PAS                                                      *)
'(* Version 2.0                                                               *)
'(* Last modified: October 1, 1992                                            *)
'(*****************************************************************************)

  '{ Table 35.A }
  '{ The outline is somewhat different from Table 35.A, since the     }
  '{ coefficients of k are the same for oppositions and conjunctions. }
  Private MeanPhenomData As Variant
  Private CoeffTable As Variant
  Private ExtraAngles As Variant
  Private ExtraCoeffs As Variant
  Private Index As Variant
  Private ElongCoeff As Variant
  Private ElongIndex As Variant

'(*****************************************************************************)
'(* Name:    FirstkOfYear                                                     *)
'(* Type:    Function                                                         *)
'(* Purpose: Approximate 'k' corresponding to the beginning of a year.        *)
'(* Arguments:                                                                *)
'(*   Year : the year                                                         *)
'(*   Planet : index to planet (Mercury = 1, Neptune = 8                      *)
'(*   Event : one of OPPOSITION, CONJUNCTION, INFCONJ, SUPCONJ                *)
'(* Return value:                                                             *)
'(*   the 'k' value corresponding to approximately the first occurance of     *)
'(*   event in Year                                                           *)
'(*****************************************************************************)

Function FirstkOfYear(Year As Long, ByVal Planet As Long, lEvent As Long) As Long
Dim n As Double, D As Double
If Planet > 2 Then Planet = Planet - 1
n = 365.2425 * Year + 1721060# - MeanPhenomData(Planet)(lEvent)
D = MeanPhenomData(Planet)(2)
FirstkOfYear = Int(n / D)
End Function

'(*****************************************************************************)
'(* Name:   FindMeanValues                                                    *)
'(* Type:   Support function                                                  *)
'(* Purpose:Calculate the mean values, JDE0 and M corresponding to a          *)
'(*         a given event and k value.                                        *)
'(*****************************************************************************)

Sub FindMeanValues(Planet As Long, lEvent As Long, k As Long, ByRef JDE0 As Double, ByRef m As Double)
JDE0 = MeanPhenomData(Planet)(lEvent) + k * MeanPhenomData(Planet)(2)
m = (MeanPhenomData(Planet)(3 + lEvent) + k * MeanPhenomData(Planet)(5)) * DToR
End Sub

'(*****************************************************************************)
'(* Name:    ConjunctionOpposition                                            *)
'(* Type:    Function                                                         *)
'(* Purpose: Instant (JD) of conjunction or opposition of a planet.           *)
'(* Arguments:                                                                *)
'(*   Planet : index to planet (Mercury = 1, Neptune = 8                      *)
'(*   Event : one of OPPOSITION, zero of CONJUNCTION, INFCONJ, SUPCONJ        *)
'(*   k : index from FirstkOfYear                                             *)
'(* Return value:                                                             *)
'(*   Julian day of instance when the Event occurs                            *)
'(*****************************************************************************)

Function ConjunctionOpposition(ByVal Planet As Long, ByVal lEvent As Long, ByVal k As Long) As Double
Dim JDE0   As Double, m As Double, T As Double, sum As Double, A As Double
Dim i As Long, j As Long
Dim MTab(1) As TSINCOSTAB

'{ Earth is left out, so decrement Planet if greater than 3. }
If Planet > 2 Then Planet = Planet - 1
Call FindMeanValues(Planet, lEvent, k, JDE0, m)
T = JDToT(JDE0)
Call CalcSinCosTab(m, Index(Planet)(0), MTab(0), MTab(1))
i = Index(Planet)(1)
sum = Eval2Poly(C2P(CoeffTable(i)(lEvent)), T)
For j = 1 To Index(Planet)(0)
  For k = 0 To 1
    sum = sum + MTab(k).W(j) * Eval2Poly(C2P(CoeffTable(i + 2 * j - 1 + k)(lEvent)), T)
  Next
Next
Select Case Planet
Case 4    '{ Jupiter }
    A = (ExtraAngles(1)(0) + T * ExtraAngles(1)(1)) * DToR
    sum = sum + Eval2Poly(C2P(ExtraCoeffs(0)), T) * Sin(A)
    sum = sum + Eval2Poly(C2P(ExtraCoeffs(1)), T) * Cos(A)
Case 5    '{ Saturn }
    For i = 1 To 4
      A = (ExtraAngles(i)(0) + T * ExtraAngles(i)(1)) * DToR
      sum = sum + Eval2Poly(C2P(ExtraCoeffs(2 * i)), T) * Sin(A)
      sum = sum + Eval2Poly(C2P(ExtraCoeffs(2 * i + 1)), T) * Cos(A)
    Next
Case 6    '{ Uranus }
    A = (ExtraAngles(5)(0) + T * ExtraAngles(5)(1)) * DToR
    sum = sum + 0.885 * Cos(A)
    A = (ExtraAngles(6)(0) + T * ExtraAngles(6)(1)) * DToR
    sum = sum + 0.2153 * Cos(A)
Case 7    '{ Neptune }
    A = (ExtraAngles(5)(0) + T * ExtraAngles(5)(1)) * DToR
    sum = sum - 0.5964 * Cos(A)
    A = (ExtraAngles(7)(0) + T * ExtraAngles(7)(1)) * DToR
    sum = sum + 0.0728 * Cos(A)
End Select
ConjunctionOpposition = JDE0 + sum
End Function

'(*****************************************************************************)
'(* Name:    GreatestElongation                                               *)
'(* Type:    Function                                                         *)
'(* Purpose: Instant (JD) of greatest elongation of Mercury or Venus.         *)
'(* Arguments:                                                                *)
'(*   Planet : index to planet (Mercury = 1, Neptune = 8                      *)
'(*   Event : one of WESTELONGATION and EASTELONGATION                        *)
'(*   k : index from FirstkOfYear                                             *)
'(* Return value:                                                             *)
'(*   Julian day of instance when the Event occurs                            *)
'(*****************************************************************************)

Function GreatestElongation(Planet As Long, lEvent As Long, k As Long, ByRef Elongation As Double) As Double
Dim JDE0   As Double, T As Double, m As Double, sum As Double
Dim i As Long, j As Long, l As Long
Dim MTab(1) As TSINCOSTAB

Call FindMeanValues(Planet, INFCONJ, k, JDE0, m)
T = JDToT(JDE0)
Call CalcSinCosTab(m, ElongIndex(Planet)(lEvent)(0), MTab(0), MTab(1))
i = ElongIndex(Planet)(lEvent)(1)
sum = Eval2Poly(C2P(ElongCoeff(i)(0)), T)
Elongation = Eval2Poly(C2P(ElongCoeff(i)(1)), T)
For j = 1 To ElongIndex(Planet)(lEvent)(0)
  For l = 0 To 1
    sum = sum + MTab(l).W(j) * Eval2Poly(C2P(ElongCoeff(i + 2 * j - 1 + l)(0)), T)
    Elongation = Elongation + MTab(l).W(j) * Eval2Poly(C2P(ElongCoeff(i + 2 * j - 1 + l)(1)), T)
  Next
Next
GreatestElongation = JDE0 + sum
End Function

Private Sub Class_Initialize()
  '{ Table 35.A }
'  MeanPhenomData : array(1..7, 0..5)  as double = (
  MeanPhenomData = Array( _
    Array(0, 0, 0, 0, 0, 0), _
    Array(2451612.023, 2451554.084, 115.8774771, 63.5867, 6.4822, 114.2088742), _
    Array(2451996.706, 2451704.746, 583.921361, 82.7311, 154.9745, 215.513058), _
    Array(2452097.382, 2451707.414, 779.936104, 181.9573, 157.6047, 48.705244), _
    Array(2451870.628, 2451671.186, 398.884046, 318.4681, 121.898, 33.140229), _
    Array(2451870.17, 2451681.124, 378.091904, 318.0172, 131.6934, 12.647487), _
    Array(2451764.317, 2451579.489, 369.656035, 213.6884, 31.5219, 4.333093), _
    Array(2451753.122, 2451569.379, 367.486703, 202.6544, 21.5569, 2.194998))

  '{ Table 35.B }
'  CoeffTable : array(0..54,0..1) of T2POLY  = (
  CoeffTable = Array( _
    Array(Array(0.0545, 0.0002, 0), Array(-0.0548, -0.0002, 0)), Array(Array(-6.2008, 0.0074, 0.00003), Array(7.3894, -0.01, -0.00003)), Array(Array(-3.275, -0.0197, 0.00001), Array(3.22, 0.0197, -0.00001)), Array(Array(0.4737, -0.0052, -0.00001), Array(0.8383, -0.0064, -0.00001)), _
    Array(Array(0.8111, 0.0033, -0.00002), Array(0.9666, 0.0039, -0.00003)), Array(Array(0.0037, 0.0018, 0), Array(0.077, -0.0026, 0)), Array(Array(-0.1768, 0, 0.00001), Array(0.2758, 0.0002, -0.00002)), Array(Array(-0.0211, -0.0004, 0), Array(-0.0128, -0.0008, 0)), _
    Array(Array(0.0326, -0.0003, 0), Array(0.0734, -0.0004, -0.00001)), Array(Array(0.0083, 0.0001, 0), Array(-0.0122, -0.0002, 0)), Array(Array(-0.004, 0.0001, 0), Array(0.0173, -0.0002, 0)), Array(Array(-0.0096, 0.0002, -0.00001), Array(0.0099, -0.0002, -0.00001)), _
    Array(Array(2.0009, -0.0033, -0.00001), Array(4.1991, -0.0121, -0.00003)), Array(Array(0.598, -0.0104, 0.00001), Array(-0.6095, 0.0102, -0.00002)), Array(Array(0.0967, -0.0018, -0.00003), Array(0.25, -0.0028, -0.00003)), Array(Array(0.0913, 0.0009, -0.00002), Array(0.0063, 0.0025, -0.00002)), _
    Array(Array(0.0046, -0.0002, 0), Array(0.0232, -0.0005, -0.00001)), Array(Array(0.0079, 0.0001, 0), Array(0.0031, 0.0004, 0)), Array(Array(-0.3088, 0, 0.00002), Array(0.3102, -0.0001, 0.00001)), Array(Array(-17.6965, 0.0363, 0.00005), Array(9.7273, -0.0156, 0.00001)), _
    Array(Array(18.3131, 0.0467, -0.00006), Array(-18.3195, -0.0467, 0.00009)), Array(Array(-0.2162, -0.0198, -0.00001), Array(-1.6488, -0.0133, 0.00001)), Array(Array(-4.5028, -0.0019, 0.00007), Array(-2.6117, -0.002, 0.00004)), Array(Array(0.8987, 0.0058, -0.00002), Array(-0.6827, -0.0026, 0.00001)), _
    Array(Array(0.7666, -0.005, -0.00003), Array(0.0281, 0.0035, 0.00001)), Array(Array(-0.3636, -0.0001, 0.00002), Array(-0.0823, 0.0006, 0.00001)), Array(Array(0.0402, 0.0032, 0), Array(0.1584, 0.0013, 0)), Array(Array(0.0737, -0.0008, 0), Array(0.027, 0.0005, 0)), _
    Array(Array(-0.098, -0.0011, 0), Array(0.0433, 0, 0)), Array(Array(-0.1029, 0, -0.00009), Array(0.1027, 0.0002, -0.00009)), Array(Array(-1.9658, -0.0056, 0.00007), Array(-2.2637, 0.0163, -0.00003)), Array(Array(6.1537, 0.021, -0.00006), Array(-6.154, -0.021, 0.00008)), _
    Array(Array(-0.2081, -0.0013, 0), Array(-0.2021, -0.0017, 0.00001)), Array(Array(-0.1116, -0.001, 0), Array(0.131, -0.0008, 0)), Array(Array(0.0074, 0.0001, 0), Array(0.0086, 0, 0)), Array(Array(-0.0097, -0.0001, 0), Array(0.0087, 0.0002, 0)), _
    Array(Array(-0.0209, 0.0006, 0.00023), Array(0.0172, -0.0006, 0.00023)), Array(Array(4.5795, -0.0312, -0.00017), Array(-8.5885, 0.0411, 0.0002)), Array(Array(1.1462, -0.0351, 0.00011), Array(-1.147, 0.0352, -0.00011)), Array(Array(0.0985, -0.0015, 0), Array(0.3331, -0.0034, -0.00001)), _
    Array(Array(0.0733, -0.0031, 0.00001), Array(0.1145, -0.0045, 0.00002)), Array(Array(0.0025, -0.0001, 0), Array(-0.0169, 0.0002, 0)), Array(Array(0.005, -0.0002, 0), Array(-0.0109, 0.0004, 0)), _
    Array(Array(0.0844, -0.0006, 0), Array(-0.0859, 0.0003, 0)), Array(Array(-0.1048, 0.0246, 0), Array(-3.8179, -0.0148, 0.00003)), _
    Array(Array(-5.1221, 0.0104, 0.00003), Array(5.1228, -0.0105, -0.00002)), Array(Array(-0.1428, 0.0005, 0), Array(-0.0803, 0.0011, 0)), _
    Array(Array(-0.0148, -0.0013, 0), Array(-0.1905, -0.0006, 0)), Array(Array(0, 0, 0), Array(0.0088, 0.0001, 0)), _
    Array(Array(0.0055, 0, 0), Array(0, 0, 0)), Array(Array(-0.014, 0, 0.00001), Array(0.0168, 0, 0)), _
    Array(Array(-1.3486, 0.001, 0.00001), Array(-2.5606, 0.0088, 0.00002)), Array(Array(0.8597, 0.0037, 0), Array(-0.8611, -0.0037, 0.00002)), _
    Array(Array(-0.0082, -0.0002, 0.00001), Array(0.0118, -0.0004, 0.00001)), Array(Array(0.0037, -0.0003, 0), Array(0.0307, -0.0003, 0)))

  '{ Auxiliary angles a, b, ..., g }
'  ExtraAngles : array(1..7,0..1)  as double = (
  ExtraAngles = Array( _
    Array(0, 0), _
    Array(82.74, 40.76), _
    Array(29.86, 1181.36), _
    Array(14.13, 590.68), _
    Array(220.02, 1262.87), _
    Array(207.83, 8.51), _
    Array(108.84, 419.96), _
    Array(276.74, 209.98))

  '{ Coefficient of cos a, sin a, ..., cos d for Saturn and Jupiter }
'  ExtraCoeffs : array(0..9) of T2POLY = (
  ExtraCoeffs = Array( _
    Array(0, 0.0144, -0.00008), _
    Array(0.3642, -0.0019, -0.00029), _
    Array(0, -0.0337, 0.00018), _
    Array(-0.851, 0.0044, 0.00068), _
    Array(0, -0.0064, 0.00004), _
    Array(0.2397, -0.0012, -0.00008), _
    Array(0, -0.001, 0), _
    Array(0.1245, 0.0006, 0), _
    Array(0, 0.0024, -0.00003), _
    Array(0.0477, -0.0005, -0.00006))

  '{ Index into CoeffTable :                         }
  '{   1st field = N umber of terms                  }
  '{   2nd field = index of first term in CoeffTable }
'  Index : array(1..7, 0..1) of long = (
  Index = Array(Array(0, 0), _
    Array(5, 0), Array(3, 11), Array(5, 18), Array(3, 29), _
    Array(3, 36), Array(3, 43), Array(2, 50))

  '{ Table 35.C }
'  ElongCoeff : array(0..35,0..1) of T2POLY = (
  ElongCoeff = Array( _
    Array(Array(-21.6101, 0.0002, 0), Array(22.4697, 0, 0)), Array(Array(-1.9803, -0.006, 0.00001), Array(-4.2666, 0.0054, 0.00002)), _
    Array(Array(1.4151, -0.0072, -0.00001), Array(-1.8537, -0.0137, 0)), Array(Array(0.5528, -0.0005, -0.00001), Array(0.3598, 0.0008, -0.00001)), _
    Array(Array(0.2905, 0.0034, 0.00001), Array(-0.068, 0.0026, 0)), Array(Array(-0.1121, -0.0001, 0.00001), Array(-0.0524, -0.0003, 0)), _
    Array(Array(-0.0098, -0.0015, 0), Array(0.0052, -0.0006, 0)), Array(Array(0.0192, 0, 0), Array(0.0107, 0.0001, 0)), _
    Array(Array(0.0111, 0.0004, 0), Array(-0.0013, 0.0001, 0)), Array(Array(-0.0061, 0, 0), Array(-0.0021, 0, 0)), _
    Array(Array(-0.0032, -0.0001, 0), Array(0.0003, 0, 0)), Array(Array(21.6249, -0.0002, 0), Array(22.4143, -0.0001, 0)), _
    Array(Array(0.1306, 0.0065, 0), Array(4.3651, -0.0048, -0.00002)), Array(Array(-2.7661, -0.0011, 0.00001), Array(2.3787, 0.0121, -0.00001)), _
    Array(Array(0.2438, -0.0024, -0.00001), Array(0.2674, 0.0022, 0)), Array(Array(0.5767, 0.0023, 0), Array(-0.3873, 0.0008, 0.00001)), _
    Array(Array(0.1041, 0, 0), Array(-0.0369, -0.0001, 0)), Array(Array(-0.0184, 0.0007, 0), Array(0.0017, -0.0001, 0)), _
    Array(Array(-0.0051, -0.0001, 0), Array(0.0059, 0, 0)), Array(Array(0.0048, 0.0001, 0), Array(0.0061, 0.0001, 0)), _
    Array(Array(0.0026, 0, 0), Array(0.0007, 0, 0)), Array(Array(0.0037, 0, 0), Array(-0.0011, 0, 0)), _
    Array(Array(-70.76, 0.0002, -0.00001), Array(46.3173, 0.0001, 0)), Array(Array(1.0282, -0.001, -0.00001), Array(0.6916, -0.0024, 0)), _
    Array(Array(0.2761, -0.006, 0), Array(0.6676, -0.0045, 0)), Array(Array(-0.0438, -0.0023, 0.00002), Array(0.0309, -0.0002, 0)), _
    Array(Array(0.166, -0.0037, -0.00004), Array(0.0036, -0.0001, 0)), Array(Array(0.0036, 0.0001, 0), Array(0, 0, 0)), _
    Array(Array(-0.0011, 0, 0.00001), Array(0, 0, 0)), Array(Array(70.7462, 0, -0.00001), Array(46.3245, 0, 0)), _
    Array(Array(1.1218, -0.0025, -0.00001), Array(-0.5366, -0.0003, 0.00001)), Array(Array(0.4538, -0.0066, 0), Array(0.3097, 0.0016, -0.00001)), _
    Array(Array(0.132, 0.002, -0.00003), Array(-0.0163, 0, 0)), Array(Array(-0.0702, 0.0022, 0.00004), Array(-0.0075, 0.0001, 0)), _
    Array(Array(0.0062, -0.0001, 0), Array(0, 0, 0)), Array(Array(0.0015, 0, -0.00001), Array(0, 0, 0)))

  '{ Index into ElongCoeffTable.                        }
  '{   1st index : planet number                        }
  '{   2nd index : event (Western or Eastern elongation }
  '{   3rd index : 'field number'                       }
  '{     1st field : number of terms                    }
  '{     2nd field : index into ElongCoeffTable         }
'  ElongIndex : array(1..2, 0..1, 0..1) of long =
  ElongIndex = Array(Array(Array(0, 0), Array(0, 0)), Array(Array(5, 0), Array(5, 11)), Array(Array(3, 22), Array(3, 29)))
End Sub
