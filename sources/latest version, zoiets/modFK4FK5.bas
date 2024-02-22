Attribute VB_Name = "modFK4FK5"
'(*****************************************************************************)
'(* Name:    FK5ToFK4                                                         *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Converts FK5/J2000.0 coordinates to the corresponding            *)
'(*          FK4/B1950.0 coordinates.                                         *)
'(* Arguments:                                                                *)
'(*   RA, Decl : coordinates to be converted                                  *)
'(*   dRA, dDecl : proper motion in the FK5 system                            *)
'(* Note:                                                                     *)
'(*   Proper motions are NOT converted.                                       *)
'(*****************************************************************************)

Sub FK5ToFK4(ByRef RA As Double, ByRef Decl As Double, ByVal dRA As Double, ByVal dDecl As Double)
  RA = RA - EquinoxCorrection(TB1950)
  RA = RA + TB1950 * dRA
  Decl = Decl + TB1950 * dDecl
  Call PrecessFK5(TJ2000, TB1950, RA, Decl)
  Call eTermsOfAberration(FK5System, RA, Decl)
End Sub

'(*****************************************************************************)
'(* Name:    FK4ToFK5                                                         *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Converts FK4/B1950.0 coordinates to the corresponding            *)
'(*          FK5/J2000.0 coordinates.                                         *)
'(* Arguments:                                                                *)
'(*   RA, Decl : coordinates to be converted                                  *)
'(*   dRA, dDecl : proper motion in the FK5 system                            *)
'(* Note:                                                                     *)
'(*   Proper motions are NOT converted.                                       *)
'(*****************************************************************************)

Sub FK4ToFK5(ByRef RA As Double, ByRef Decl As Double, ByVal dRA As Double, ByVal dDecl As Double)
  Call eTermsOfAberration(FK4System, RA, Decl)
  RA = RA - TB1950 * dRA
  Decl = Decl - TB1950 * dDecl
  Call PrecessFK4(TB1950, TJ2000, RA, Decl)
  RA = RA + EquinoxCorrection(TB1950)
End Sub

'(*****************************************************************************)
'(* Name:    eTermsOfAberration                                               *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Adds or subtracts the e-terms of aberration from equatorial      *)
'(*          coordinates.                                                     *)
'(* Arguments:                                                                *)
'(*   System : ORIGINAL system (FK4System/FK5System)                          *)
'(*   RA, Decl : coordinates to be corrected                                  *)
'(* Note:                                                                     *)
'(*   These formulae come from the Astronomical Almanac of 1984.              *)
'(*****************************************************************************)

Sub eTermsOfAberration(ByVal System As Long, ByRef RA As Double, ByRef Decl As Double)
Dim dRA As Double, dDecl As Double
  dRA = (0.0227 * HToR / 3600) * Sin(RA + 11.25 * HToR) / Cos(Decl)
  dDecl = (0.341 * SToR) * Cos(RA + 11.25 * HToR) * Sin(Decl) + 0.029 * SToR * Cos(Decl)
  If (System = FK4System) Then
      '{ Add the corrections to get FK5 position }
      RA = RA + dRA
      Decl = Decl + dDecl
  Else
      '{ Subtract the corrections to get FK4 position }
      RA = RA - dRA
      Decl = Decl - dDecl
  End If
End Sub

Sub ConvertVSOP_FK5(ByVal T As Double, l As Double, ByRef B As Double)
Dim l1 As Double, dl As Double, db As Double
l1 = l - (1.397 * T - 0.00031 * T * T) * PI / 180
dl = -0.09033 + 0.03916 * (Cos(l1) + Sin(l1)) * tan(B)
db = 0.03916 * (Cos(l1) - Sin(l1))
l = l + dl * PI / 180 / 3600
B = B + db * PI / 180 / 3600
End Sub
