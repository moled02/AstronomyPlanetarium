Attribute VB_Name = "modEquTime"
'(*****************************************************************************)
'(* Name:    EquationOfTime                                                   *)
'(* Type:    Function                                                         *)
'(* Purpose: Calculate the equation of time for the given instant.            *)
'(*                                                                           *)
'(* Arguments:                                                                *)
'(*   T : number of Julian centuries since J2000.0                            *)
'(*   RA : right ascension of the Sun                                         *)
'(*   Obl : obliquity of the ecliptic                                         *)
'(*   NutLon : nutation in longitude                                          *)
'(* Return value:                                                             *)
'(*   the equation of time in radians                                         *)
'(*****************************************************************************)

Function EquationOfTime(ByVal T As Double, ByVal RA As Double, ByVal Obl As Double, ByVal NutLon As Double) As Double
Dim L0 As Double
T = T / 10
L0 = T * (1 / 49931 - T * (1 / 15299 + T / 1988000))
L0 = (280.4664567 + T * (360007.6982779 + T * (0.03032028 + L0))) * DToR
EquationOfTime = modpi(L0 - 0.0057183 * DToR - RA + NutLon * Cos(Obl))
End Function

