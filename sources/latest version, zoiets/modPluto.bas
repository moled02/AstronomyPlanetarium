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

Dim angle(3) As Double
Dim SinTab(3) As TSINCOSTAB, CosTab(3) As TSINCOSTAB
Dim SinVal As Double, CosVal As Double, tmp As Double
Dim sum(3) As Double
Dim I As Long, j As Long, k As Long, sign As Long, Flag   As Long

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
    Array(0, -19799805#, 19850055#, -5452852#, -14974862, 66865439#, 68951812#), Array(0, 897144#, -4954829#, 3527812#, 1672790#, -11827535#, -332538#), _
    Array(0, 611149#, 1211027#, -1050748#, 327647#, 1593179#, -1438890#), Array(0, -341243#, -189585#, 178690#, -292153#, -18444#, 483220#), _
    Array(0, 129287#, -34992#, 18650#, 100340#, -65977#, -85431#), Array(0, -38164#, 30893#, -30697#, -25823#, 31174#, -6032#), _
    Array(0, 20442#, -9987#, 4878#, 11248#, -5794#, 22161#), Array(0, -4063#, -5071#, 226#, -64, 4601#, 4032#), _
    Array(0, -6016#, -3336#, 2030#, -836#, -1729#, 234#), Array(0, -3956#, 3039#, 69#, -604#, -415#, 702#), _
    Array(0, -667#, 3572#, -247#, -567#, 239#, 723#), Array(0, 1276#, 501#, -57#, 1#, 67, -67), _
    Array(0, 1152#, -917#, -122#, 175#, 1034#, -451#), Array(0, 630#, -1277#, -49#, -164#, -129#, 504#), _
    Array(0, 2571#, -459#, -197#, 199#, 480#, -231#), Array(0, 899#, -1449#, -25#, 217#, 2#, -441#), _
    Array(0, -1016#, 1043#, 589#, -248#, -3359#, 265#), Array(0, -2343#, -1012#, -269#, 711#, 7856#, -7832#), _
    Array(0, 7042#, 788#, 185#, 193#, 36#, 45763#), Array(0, 1199#, -338#, 315#, 807#, 8663#, 8547#), _
    Array(0, 418#, -67, -130#, -43#, -809#, -769#), Array(0, 120#, -274#, 5#, 3#, 263#, -144#), _
    Array(0, -60#, -159#, 2#, 17#, -126#, 32#), Array(0, -82#, -29#, 2, 5, -35#, -16#), _
    Array(0, -36#, -29#, 2, 3, -19#, -4#), Array(0, -40#, 7#, 3, 1, -15#, 8#), _
    Array(0, -14, 22, 2, -1, -4#, 12), Array(0, 4, 13, 1, -1, 5, 6), _
    Array(0, 5, 2, 0, -1, 3, 1), Array(0, -1, 0, 0, 0, 6, -2), _
    Array(0, 2, 0, 0, -2, 2, 2), Array(0, -4, 5, 2, 2, -2, -2), _
    Array(0, 4, -7, -7, 0, 14, 13), Array(0, 14, 24, 10, -8, -63, 13), _
    Array(0, -49, -34, -3, 20, 136, -236), Array(0, 163, -48, 6, 5, 273, 1065), _
    Array(0, 9, -24, 14, 17, 251, 149), Array(0, -4, 1, -2, 0, -25, -9), _
    Array(0, -3, 1, 0, 0, 9, -2), Array(0, 1, 3, 0, 0, -8, 7), _
    Array(0, -3, -1, 0, 1, 2, -10), Array(0, 5, -3, 0, 0, 19, 35), Array(0, 0, 0, 1, 0, 10, 3))

angle(1) = (34.35 + 3034.9057 * T) * DToR
angle(2) = (50.08 + 1222.1138 * T) * DToR
angle(3) = (238.96 + 144.96 * T) * DToR
Call CalcSinCosTab(angle(1), 3, SinTab(1), CosTab(1))
Call CalcSinCosTab(angle(2), 2, SinTab(2), CosTab(2))
Call CalcSinCosTab(angle(3), 6, SinTab(3), CosTab(3))
For I = 1 To 3
  sum(I) = 0
Next
For I = 1 To 43
  Flag = 0
  For j = 1 To 3
    k = PlutoAngleTab(I)(j)
    If (k <> 0) Then
      If (k < 0) Then
        k = -k
        sign = -1
      Else
        sign = 1
      End If
      If (Flag = 0) Then
        Flag = 1
        SinVal = SinTab(j).W(k) * sign
        CosVal = CosTab(j).W(k)
      Else
        tmp = CosVal * CosTab(j).W(k) - SinVal * sign * SinTab(j).W(k)
        SinVal = SinVal * CosTab(j).W(k) + CosVal * sign * SinTab(j).W(k)
        CosVal = tmp
      End If
    End If
  Next
  For j = 1 To 3
    sum(j) = sum(j) + SinVal * PlutoCoeffTab(I)(2 * j - 1) _
                     + CosVal * PlutoCoeffTab(I)(2 * j)
  Next
Next
sum(1) = (0.000001 * sum(1) + 238.958116 + 144.96 * T) * DToR
sum(2) = (0.000001 * sum(2) + -3.908239) * DToR
sum(3) = 0.0000001 * sum(3) + 40.7241346
s.L = modpi2(sum(1))
s.B = sum(2)
s.r = sum(3)
End Sub

