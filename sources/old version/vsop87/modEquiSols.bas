Attribute VB_Name = "modEquiSols"
'(*****************************************************************************)
'(* Module: EQUISOLS.PAS                                                      *)
'(* Version 2.0                                                               *)
'(* Last modified: October 1, 1992                                            *)
'(*****************************************************************************)

'(*****************************************************************************)
'(* Name:    EquinoxSolstice                                                  *)
'(* Type:    Function                                                         *)
'(* Purpose: Instant of Equinox or Solstice for a given year.                 *)
'(*                                                                           *)
'(* Arguments:                                                                *)
'(*   Year : the year for which the calculation is to be performed            *)
'(*   Event : one of SPRINGEQUINOX, SUMMERSOLSTICE, FALLEQUINOX or            *)
'(*           WINTERSOLSTICE                                                  *)
'(* Return value:                                                             *)
'(*   the Julian Day at which the event in question occurs                    *)
'(*****************************************************************************)

Function EquinoxSolstice(Year As Long, lEvent As Long) As Double
Dim Y As Double, T As Double, JDE0 As Double, W As Double, dl As Double, Angle As Double, s   As Double
Dim i As Long, Table As Long

Dim EquiSolstPoly As Variant
Dim EquiSolstTable As Variant
  
  '{ Table 26.A and 26.B }
'  EquiSolstPoly : array(0..1,0..3) of T4POLY = (
  EquiSolstPoly = Array( _
   Array(Array(1721139.29189, 365242.1374, 0.06134, 0.00111, -0.00071), _
    Array(1721233.25401, 365241.72562, -0.05323, 0.00907, 0.00025), _
    Array(1721325.70455, 365242.49558, -0.11677, -0.00297, 0.00074), _
    Array(1721414.39987, 365242.88257, -0.00769, -0.00933, -0.00006)), _
   Array(Array(2451623.80984, 365242.37404, 0.05169, -0.00411, -0.00057), _
    Array(2451716.56767, 365241.62603, 0.00325, 0.00888, -0.0003), _
    Array(2451810.21715, 365242.01767, -0.11575, 0.00337, 0.00078), _
    Array(2451900.05952, 365242.74049, -0.06223, -0.00823, 0.00032)))

  '{ Table 26.C }
'  EquiSolstTable : array(1..24,0..2) of double = (
  EquiSolstTable = Array(Array(0, 0, 0), _
    Array(485, 324.96, 1934.136), Array(203, 337.23, 32964.467), _
    Array(199, 342.08, 20.186), Array(182, 27.85, 445267.112), _
    Array(156, 73.14, 45036.886), Array(136, 171.52, 22518.443), _
    Array(77, 222.54, 65928.934), Array(74, 296.72, 3034.906), _
    Array(70, 243.58, 9037.513), Array(58, 119.81, 33718.147), _
    Array(52, 297.17, 150.678), Array(50, 21.02, 2281.226), _
    Array(45, 247.54, 29929.562), Array(44, 325.15, 31555.956), _
    Array(29, 60.93, 4443.417), Array(18, 155.12, 67555.328), _
    Array(17, 288.79, 4562.452), Array(16, 198.04, 62894.029), _
    Array(14, 199.76, 31436.921), Array(12, 95.39, 14577.848), _
    Array(12, 287.11, 31931.756), Array(12, 320.81, 34777.259), _
    Array(9, 227.73, 1222.114), Array(8, 15.45, 16859.074))
                           
If (Year < 1000) Then
  Y = Year / 1000#
  Table = 0
Else
  Y = (Year - 2000) / 1000#
  Table = 1
End If
JDE0 = Eval4Poly(C4P(EquiSolstPoly(Table)(lEvent)), Y)
T = JDToT(JDE0)
W = (35999.373 * T - 2.47) * DToR
dl = 1 + 0.0334 * Cos(W) + 0.0007 * Cos(2 * W)
s = 0
For i = 1 To 24
  Angle = (EquiSolstTable(i)(1) + EquiSolstTable(i)(2) * T) * DToR
  s = s + EquiSolstTable(i)(0) * Cos(Angle)
Next
EquinoxSolstice = JDE0 + 0.00001 * s / dl
End Function

