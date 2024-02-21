Attribute VB_Name = "modObliquity"
'(*****************************************************************************)
'(* Name:    Nutation                                                         *)
'(* Type:    Procedure                                                        *)
'(* Purpose: calculate the obliquity of the ecliptic                          *)
'(* Arguments:                                                                *)
'(*   T : number of centuries since J2000                                     *)
'(* Return value:                                                             *)
'(*   the mean obliquity of the ecliptic at the given instant                 *)
'(*****************************************************************************)

Function Obliquity(T As Double) As Double
Dim I As Double
Dim u As Double, tmp As Double
Dim OblPoly As Variant

OblPoly = Array(84381.448, -4680.93, -1.55, 1999.25, -51.38, -249.67, -39.05, 7.12, 27.87, 5.79, 2.45)
u = T / 100
tmp = OblPoly(10)
For I = 9 To 0 Step -1
    tmp = tmp * u + OblPoly(I)
Next
Obliquity = tmp * SToR
End Function

