Attribute VB_Name = "modConstants"
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
Dim i As Double
Dim u As Double, tmp As Double
Dim OblPoly(10) As Double

OblPoly(0) = 84381.448
OblPoly(1) = -4680.93
OblPoly(2) = -1.55
OblPoly(3) = 1999.25
OblPoly(4) = -51.38
OblPoly(5) = -249.67
OblPoly(6) = -39.05
OblPoly(7) = 7.12
OblPoly(8) = 27.87
OblPoly(9) = 5.79
OblPoly(10) = 2.45
u = T / 100
tmp = OblPoly(10)
For i = 9 To 0 Step -1
    tmp = tmp * u + OblPoly(i)
Next
Obliquity = tmp * SToR
End Function

