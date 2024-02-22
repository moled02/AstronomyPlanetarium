Attribute VB_Name = "modNutation"
'(*****************************************************************************)
'(* Name:    Nutation                                                         *)
'(* Type:    Procedure                                                        *)
'(* Purpose: correct right ascension and declination for nutation             *)
'(* Arguments:                                                                *)
'(*   NutLon, NutObl : nutation in longitude and obliquity                    *)
'(*   Obl : obliquity of the ecliptic                                         *)
'(*   RA, Decl : coordinates to be corrected                                  *)
'(*****************************************************************************)

Sub Nutation(NutLon As Double, NutObl As Double, Obl As Double, ByRef RA As Double, ByRef Decl As Double)

Dim CosObl As Double, SinObl As Double, CosRA As Double, SinRA As Double

CosObl = Cos(Obl): SinObl = Sin(Obl)
CosRA = Cos(RA): SinRA = Sin(RA)
If (Abs(Decl) < 1.565) Then
  RA = RA + (CosObl + SinObl * SinRA * tan(Decl)) * NutLon
  RA = RA - (CosRA * tan(Decl)) * NutObl
  Decl = Decl + (SinObl * CosRA) * NutLon + (SinRA) * NutObl
Else
  Call EclToEqu(RA, Decl, -Obl, RA, Decl)
  Call EclToEqu(RA + NutLon, Decl, Obl + NutObl, RA, Decl)
End If

End Sub

