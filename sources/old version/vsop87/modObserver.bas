Attribute VB_Name = "modObserver"
'(*****************************************************************************)
'(* Name:    ObserverCoord                                                    *)
'(* Type:    Procedure                                                        *)
'(* Purpose: calculate the geocentric rectangular coordinates of an observer. *)
'(* Arguments:                                                                *)
'(*   Phi : observer's geographic latitude                                    *)
'(*   Height : observer's height above sea level in meters                    *)
'(*   RhoCosPhi,RhoSinPhi : the observer's geocentric rectangular coordinates *)
'(*****************************************************************************)

Sub ObserverCoord(ByVal Phi As Double, ByVal Height As Double, ByRef RhoCosPhi As Double, ByRef RhoSinPhi As Double)

Const F = 1 / 298.257
Dim u As Double
Height = Height / 6378140#
u = Atn((1 - F) * tan(Phi))
RhoSinPhi = (1 - F) * Sin(u) + Height * Sin(Phi)
RhoCosPhi = Cos(u) + Height * Cos(Phi)
End Sub
