Attribute VB_Name = "modObjPos"
'****************************************************************************)
'                                                                           *)
'                  Copyright (c) 1991-1992 by Jeffrey Sax                   *)
'                            All rights reserved                            *)
'                        Published and Distributed by                       *)
'                           Willmann-Bell, Inc.                             *)
'                             P.O. Box 35025                                *)
'                        Richmond, Virginia 23235                           *)
'                Voice (804) 320-7016 FAX (804) 272-5920                    *)
'                                                                           *)
'                                                                           *)
'                NOTICE TO COMMERCIAL SOFTWARE DEVELOPERS                   *)
'                                                                           *)
'        Prior to distributing software incorporating this code             *)
'        you MUST write Willmann-Bell, Inc. at the above address            *)
'        for validation of your book's (Astronomical Algorithms             *)
'        by Jean Meeus) and software Serial Numbers.  No additional         *)
'        fees will be required BUT you MUST have the following              *)
'        notice at the start of your program(s):                            *)
'                                                                           *)
'                    This program contains code                             *)
'              Copyright (c) 1991-1992 by Jeffrey Sax                       *)
'              and Distributed by Willmann-Bell, Inc.                       *)
'                         Serial #######                                    *)
'                                                                           *)
'****************************************************************************)
' Module: OBJPOS.PAS                                                        *)
' Version 2.0                                                               *)
' Last modified: October 1, 1992                                            *)
'****************************************************************************)

'****************************************************************************)
' Name:    CalcOrbitCon                                                     *)
' Type:    Procedure                                                        *)
' Purpose: Calculate the 6 orbital constants from the orhital elements.     *)
' Arguments:                                                                *)
'   OrbitEl : the orbital elements of the object                            *)
'   Obl : the obliquity of the ecliptic                                     *)
'   OrbitCon : TORBITCON record to hold the results                         *)
'****************************************************************************)

Sub CalcOrbitCon(OrbitEl As TORBITEL, Obl As Double, ByRef OrbitCon As TORBITCON)

Dim cOm As Double, sOm As Double, cI As Double, si As Double, cObl As Double, sObl As Double
Dim F As Double, g As Double, h As Double, P As Double, Q As Double, R As Double

cOm = Cos(OrbitEl.LonNode)
sOm = Sin(OrbitEl.LonNode)
cI = Cos(OrbitEl.incl)
si = Sin(OrbitEl.incl)
cObl = Cos(Obl)
sObl = Sin(Obl)

F = cOm
g = sOm * cObl
h = sOm * sObl
P = -sOm * cI
Q = cOm * cI * cObl - si * sObl
R = cOm * cI * sObl + si * cObl

OrbitCon.A = atan2(F, P)
OrbitCon.B = atan2(g, Q)
OrbitCon.C = atan2(h, R)
OrbitCon.aa = Sqr(F * F + P * P)
OrbitCon.bb = Sqr(g * g + Q * Q)
OrbitCon.cc = Sqr(h * h + R * R)
End Sub

'****************************************************************************)
' Name:    PosRectCo                                                        *)
' Type:    Procedure                                                        *)
' Purpose: Calculate rectangular equatorial coordinates of an object from   *)
'          its orbital elements and -constants                              *)
' Arguments:                                                                *)
'   T : number of Julian centuries since J2000.0                            *)
'   OrbitEl : the orbital elements of the object                            *)
'   OrbitCon : orbital constants associated with the orbital elements       *)
'   R : TVECTOR array to hold the equatorial coordinates                    *)
'****************************************************************************)

Sub PosRectCo(ByVal T As Double, OrbitEl As TORBITEL, OrbitCon As TORBITCON, ByRef R As TVECTOR)

Dim MeanAn As Double, TrueAn As Double, rv As Double, Theta As Double
If OrbitEl.E < 1 Then
  MeanAn = OrbitEl.M0 + (T - OrbitEl.t0) * OrbitEl.n * 36525
  TrueAn = Kepler(MeanAn, OrbitEl.E)
  rv = OrbitEl.A * (1 - OrbitEl.E * OrbitEl.E) / (1 + OrbitEl.E * Cos(TrueAn))
ElseIf OrbitEl.E = 1 Then
  TrueAn = Parabola(T - OrbitEl.t0, OrbitEl.Q)
  rv = OrbitEl.Q * (1 + OrbitEl.E) / (1 + OrbitEl.E * Cos(TrueAn))
Else
  TrueAn = NearParabola(T - OrbitEl.t0, OrbitEl.Q, OrbitEl.E)
  rv = OrbitEl.Q * (1 + OrbitEl.E) / (1 + OrbitEl.E * Cos(TrueAn))
End If
R.x = rv * OrbitCon.aa * Sin(OrbitCon.A + OrbitEl.LonPeri + TrueAn)
R.Y = rv * OrbitCon.bb * Sin(OrbitCon.B + OrbitEl.LonPeri + TrueAn)
R.Z = rv * OrbitCon.cc * Sin(OrbitCon.C + OrbitEl.LonPeri + TrueAn)
End Sub

