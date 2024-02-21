Attribute VB_Name = "ModDeltat"
'(*****************************************************************************)
'(* Module: DELTAT.PAS                                                        *)
'(* Version 2.0                                                               *)
'(* Last modified: October 1, 1992                                            *)
'(*****************************************************************************)

Private DeltaTable As Variant

'(*****************************************************************************)
'(* Name:    ApproxDeltaT                                                     *)
'(* Type:    Function                                                         *)
'(* Purpose: Calculate an approximate value for DeltaT (= TD - UT) in seconds *)
'(* Arguments:                                                                *)
'(*   T : number of Julian centuries since J2000                              *)
'(* Return value:                                                             *)''
'(*   an approximate value for DeltaT, if possible interpolated from the      *)
'(*   table above                                                             *)
'(*****************************************************************************)

Function ApproxDeltaT(ByVal T As Double) As Double

Dim Y As Double
Dim Index As Long

  
'  { Table 9.A }
DeltaTTable = Array _
 (1240, 1150, 1060, 980, 910, 850, 790, 740, 700, 650, _
   620, 580, 550, 530, 500, 480, 460, 440, 420, 400, _
   370, 350, 330, 310, 280, 260, 240, 220, 200, 180, _
   160, 140, 130, 120, 110, 100, 90, 90, 90, 90, _
    90, 90, 90, 90, 100, 100, 100, 100, 100, 110, _
   110, 110, 110, 110, 110, 110, 110, 120, 120, 120, _
   120, 120, 130, 130, 130, 130, 140, 140, 140, 150, _
   150, 150, 150, 160, 160, 160, 160, 160, 170, 170, _
   170, 170, 170, 170, 170, 170, 160, 160, 150, 140, _
   137, 131, 127, 125, 125, 125, 125, 125, 125, 123, _
   120, 114, 106, 96, 86, 75, 66, 60, 57, 56, _
    57, 59, 62, 65, 68, 71, 73, 75, 77, 78, _
    79, 75, 64, 54, 29, 16, -10, -27, -36, -47 _
   - 54, -52, -55, -56, -58, -59, -62, -64, -61, -47, _
   -27, 0, 26, 54, 77, 105, 134, 160, 182, 202, _
   212, 224, 235, 239, 243, 240, 239, 239, 237, 240, _
   243, 253, 262, 273, 282, 291, 300, 307, 314, 322, _
   331, 340, 350, 365, 383, 402, 422, 445, 465, 485, _
   505, 522, 538, 549, 558, 569, 580)

Y = 2000 + T * 100
If Int(Y) >= 2005 Then
    Y = Y - 2000
    ApproxDeltaT = 62.92 + 0.23217 * Y + 0.005589 * Y ^ 2
ElseIf Y >= 1986 Then
    Y = Y - 2000
    ApproxDeltaT = 63.86 + 0.3345 * Y - 0.060374 * Y ^ 2 + 0.0017275 * Y ^ 3 + 0.000651814 * Y ^ 4 + 0.00002373599 * Y ^ 5
ElseIf Y >= 1961 Then
    Y = Y - 1975
    ApproxDeltaT = 45.45 + 1.067 * Y - Y ^ 2 / 260 - Y ^ 3 / 718
Else
  If (Y < 1620) Then
    If (Y < 948) Then
      ApproxDeltaT = 2715.6 + T * (573.36 + T * 46.5)
    Else
      ApproxDeltaT = 50.6 + T * (67.5 + T * 22.5)
    End If
  Else '{ Interpolate from the above table }
    Index = Int((Y - 1620) / 2)
    If Index > 184 Then Index = 184
    Y = Y / 2 - Index - 810
    ApproxDeltaT = (DeltaTTable(Index) + (DeltaTTable(Index + 1) - DeltaTTable(Index)) * Y) / 10
  End If
End If
If 2000 + T * 100 >= 1961 Then ApproxDeltaT = ApproxDeltaT - (0.000012932 * (2000 + T * 100 - 1955) ^ 2)
End Function

