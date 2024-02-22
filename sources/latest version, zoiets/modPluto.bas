Attribute VB_Name = "modPluto"
'(*****************************************************************************)
'(* Name:    PlutoPos                                                         *)
'(* Type:    Procedure                                                        *)
'(* Purpose: calculate Pluto's heliocentric ecliptical coordinates.           *)
'(* Arguments:                                                                *)
'(*   T : number of centuries since J2000                                     *)
'(*   S : TSVECTOR record to hold the coordinates                             *)
'(*****************************************************************************)

Sub PlutoPos(T As Double, ByRef s As TSVECTOR)

Dim Angle(3) As Double
Dim SinTab(3) As TSINCOSTAB, CosTab(3) As TSINCOSTAB
Dim SinVal As Double, CosVal As Double, Tmp As Double
Dim sum(3) As Double
Dim i As Long, j As Long, k As Long, sign As Long, flag   As Long

Dim PlutoAngleTab As Variant
Dim PlutoCoeffTab As Variant

  
  '{ Table 36.A }
  '{ PlutoAngleTab contains the coefficients of the angles J, S and P. }
  '{ PlutoCoeffTab contains the coefficients of the sin and cos terms  }
  PlutoAngleTab = Array(Array(0, 0, 0, 0), _
    Array(0, 0, 0, 1), Array(0, 0, 0, 2), Array(0, 0, 0, 3), Array(0, 0, 0, 4), Array(0, 0, 0, 5), Array(0, 0, 0, 6), _
    Array(0, 0, 1, -1), Array(0, 0, 1, 0), Array(0, 0, 1, 1), Array(0, 0, 1, 2), Array(0, 0, 1, 3), Array(0, 0, 2, -2), _
    Array(0, 0, 2, -1), Array(0, 0, 2, 0), Array(0, 1, -1, 0), Array(0, 1, -1, 1), Array(0, 1, 0, -3), Array(0, 1, 0, -2), _
    Array(0, 1, 0, -1), Array(0, 1, 0, 0), Array(0, 1, 0, 1), Array(0, 1, 0, 2), Array(0, 1, 0, 3), Array(0, 1, 0, 4), _
    Array(0, 1, 1, -3), Array(0, 1, 1, -2), Array(0, 1, 1, -1), Array(0, 1, 1, 0), Array(0, 1, 1, 1), Array(0, 1, 1, 3), _
    Array(0, 2, 0, -6), Array(0, 2, 0, -5), Array(0, 2, 0, -4), Array(0, 2, 0, -3), Array(0, 2, 0, -2), Array(0, 2, 0, -1), _
    Array(0, 2, 0, 0), Array(0, 2, 0, 1), Array(0, 2, 0, 2), Array(0, 2, 0, 3), Array(0, 3, 0, -2), Array(0, 3, 0, -1), _
    Array(0, 3, 0, 0))

  PlutoCoeffTab = Array(Array(0, 0, 0, 0, 0, 0, 0), _
    Array(0, -19798886, 19848454, -5453098, -14974876, 66867334, 68955876), Array(0, 897499, -4955707, 3527363, 1672673, -11826086, -333765), _
    Array(0, 610820, 1210521, -1050939, 327763, 1593657, -1439953), Array(0, -341639, -189719, 178691, -291925, -18948, 482443), _
    Array(0, 129027, -34863, 18763, 100448, -66634, -85576), Array(0, -38215, 31061, -30594, -25838, 30841, -5765), _
    Array(0, 20349, -9886, 4965, 11263, -6140, 22254), Array(0, -4045, -4904, 310, -132, 4434, 4443), _
    Array(0, -5885, -3238, 2036, -947, -1518, 641), Array(0, -3812, 3011, -2, -674, -5, 792), _
    Array(0, -601, 3468, -329, -563, 518, 518), Array(0, 1237, 463, -64, 39, -13, -221), _
    Array(0, 1086, -911, -94, 210, 837, -494), Array(0, 595, -1229, -8, -160, -281, 616), _
    Array(0, 2484, -485, -177, 259, 260, -395), Array(0, 839, -1414, 17, 234, -191, -396), _
    Array(0, -964, 1059, 582, -285, -3218, 370), Array(0, -2303, -1038, -298, 692, 8019, -7869), _
    Array(0, 7049, 747, 157, 201, 105, 45637), Array(0, 1179, -358, 304, 825, 8623, 8444), _
    Array(0, 393, -63, -124, -29, -896, -801), Array(0, 111, -268, 15, 8, 208, -122), _
    Array(0, -52, -154, 7, 15, -133, 65), Array(0, -78, -30, 2, 2, -16, 1), _
    Array(0, -34, -26, 4, 2, -22, 7), Array(0, -43, 1, 3, 0, -8, 16), _
    Array(0, -15, 21, 1, -1, 2, 9), Array(0, -1, 15, 0, -2, 12, 5), _
    Array(0, 4, 7, 1, 0, 1, -3), Array(0, 1, 5, 1, -1, 1, 0), _
    Array(0, 8, 3, -2, -3, 9, 5), Array(0, -3, 6, 1, 2, 2, -1), _
    Array(0, 6, -13, -8, 2, 14, 10), Array(0, 10, 22, 10, -7, -65, 12), _
    Array(0, -57, -32, 0, 21, 126, -233), Array(0, 157, -46, 8, 5, 270, 1068), _
    Array(0, 12, -18, 13, 16, 254, 155), Array(0, -4, 8, -2, -3, -26, -2), _
    Array(0, -5, 0, 0, 0, 7, 0), Array(0, 3, 4, 0, 1, -11, 4), _
    Array(0, -1, -1, 0, 1, 4, -14), Array(0, 6, -3, 0, 0, 18, 35), Array(0, -1, -2, 0, 1, 13, 3))

Angle(1) = (34.35 + 3034.9057 * T) * DToR
Angle(2) = (50.08 + 1222.1138 * T) * DToR
Angle(3) = (238.96 + 144.96 * T) * DToR
Call CalcSinCosTab(Angle(1), 3, SinTab(1), CosTab(1))
Call CalcSinCosTab(Angle(2), 2, SinTab(2), CosTab(2))
Call CalcSinCosTab(Angle(3), 6, SinTab(3), CosTab(3))
For i = 1 To 3
  sum(i) = 0
Next
For i = 1 To 43
  flag = 0
  For j = 1 To 3
    k = PlutoAngleTab(i)(j)
    If (k <> 0) Then
      If (k < 0) Then
        k = -k
        sign = -1
      Else
        sign = 1
      End If
      If (flag = 0) Then
        flag = 1
        SinVal = SinTab(j).W(k) * sign
        CosVal = CosTab(j).W(k)
      Else
        Tmp = CosVal * CosTab(j).W(k) - SinVal * sign * SinTab(j).W(k)
        SinVal = SinVal * CosTab(j).W(k) + CosVal * sign * SinTab(j).W(k)
        CosVal = Tmp
      End If
    End If
  Next
  For j = 1 To 3
    sum(j) = sum(j) + SinVal * PlutoCoeffTab(i)(2 * j - 1) _
                     + CosVal * PlutoCoeffTab(i)(2 * j)
  Next
Next
sum(1) = (0.000001 * sum(1) + 238.956785 + 144.96 * T) * DToR
sum(2) = (0.000001 * sum(2) + -3.908202) * DToR
sum(3) = 0.0000001 * sum(3) + 40.7247248
s.l = modpi2(sum(1))
s.B = sum(2)
s.r = sum(3)
End Sub

