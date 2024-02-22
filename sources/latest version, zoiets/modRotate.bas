Attribute VB_Name = "modRotate"

'(*****************************************************************************)
'(* Name:    XRot                                                             *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Rotates a vector about the X-axis given the sine and cosine of   *)
'(*          the angle.                                                       *)
'(* Arguments:                                                                *)
'(*   v : TVECTOR holding the vector to rotate                                *)
'(*   c, s : cosine and sine of the angle to rotate through                   *)
'(*   w : TVECTOR to hold the rotated vector                                  *)
'(*****************************************************************************)

Sub XRot(ByRef v As TVECTOR, ByVal C As Double, ByVal s As Double, ByRef W As TVECTOR)
Dim Tmp As Double

W.x = v.x
Tmp = C * v.Y - s * v.Z
W.Z = s * v.Y + C * v.Z
W.Y = Tmp
End Sub

'(*****************************************************************************)
'(* Name:    ZRot                                                             *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Rotates a vector about the Z-axis given the sine and cosine of   *)
'(*          the angle.                                                       *)
'(* Arguments:                                                                *)
'(*   v : TVECTOR holding the vector to rotate                                *)
'(*   c, s : cosine and sine of the angle to rotate through                   *)
'(*   w : TVECTOR to hold the rotated vector                                  *)
'(*****************************************************************************)

Sub ZRot(ByRef v As TVECTOR, ByVal C As Double, ByVal s As Double, ByRef W As TVECTOR)
Dim Tmp As Double
    Tmp = C * v.x - s * v.Y
    W.Y = s * v.x + C * v.Y
    W.Z = v.Z
    W.x = Tmp
End Sub



