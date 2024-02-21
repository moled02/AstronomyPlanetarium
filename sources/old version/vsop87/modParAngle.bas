Attribute VB_Name = "modParAngle"
'(*****************************************************************************)
'(* Name:    ParallacticAngle                                                 *)
'(* Type:    Function                                                         *)
'(* Purpose: Parallactic angle for any position.                              *)
'(* Arguments:                                                                *)
'(*   RA, Decl : coordinates of the object                                    *)
'(*   ObsLat : observer's latitude                                            *)
'(*   LAST : Local (Apparent) Sidereal Time                                   *)
'(* Return value:                                                             *)
'(*   the parallactic angle                                                   *)
'(*****************************************************************************)

Function ParallacticAngle(ByVal RA As Double, ByVal Decl As Double, ByVal ObsLat As Double, ByVal LAST As Double) As Double
Dim H As Double, D As Double, n As Double
H = LAST - RA
D = Sin(H)
n = tan(ObsLat) * Cos(Decl) - Sin(Decl) * Cos(H)
ParallacticAngle = atan2(D, n)
End Function

