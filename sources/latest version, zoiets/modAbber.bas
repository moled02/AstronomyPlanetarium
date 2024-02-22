Attribute VB_Name = "modAbber"
Const AbConst = (20.49552 * SToR)

'(*****************************************************************************)
'(* Name:    SunLonAndEcc                                                     *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Calculate the Sun's longitude and orbital eccentricity (low acc).*)
'(* Arguments:                                                                *)
'(*   T : Julian centuries since J2000                                        *)
'(*   Lon : longitude of the Sun                                              *)
'(*   e : eccentricity of the Earth's orbit                                   *)
'(*****************************************************************************)

Sub SunLonAndEcc(T As Double, ByRef Lon As Double, ByRef E As Double)
Dim l As Double, M As Double, C As Double, e2 As Double
l = (280.46646 + T * (36000.76983 + T * 0.0003032)) * DToR
M = (357.5291 + T * (35999.05028 - T * 0.0001561)) * DToR
E = 0.016708617 - T * (0.00004204 + T * 0.0000001236)
e2 = E * E '{ Equation of the center, in terms of e and M }
C = E * (2 - 0.25 * e2) * Sin(M) + 1.25 * e2 * Sin(2 * M) + 1.0833333333 * E * e2 * Sin(3 * M)
Lon = l + C
End Sub


'(*****************************************************************************)
'(* Name:    Aberration                                                       *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Corrects equatorial coordinates for the effect of aberration.    *)
'(* Arguments:                                                                *)
'(*   T : Julian centuries since J2000                                        *)
'(*   Obl : obliquity of the ecliptic                                         *)
'(*   System : either FK4System or FK5System                                  *)
'(*   RA, Decl : coordinates to be corrected                                  *)
'(*****************************************************************************)

Sub Aberration(T As Double, Obl As Double, System As Long, ByRef RA As Double, ByRef Decl As Double)

Dim CosObl As Double, SinObl As Double
Dim Lon     As Double, E As Double, PI As Double, Tmp As Double
Dim CosRA     As Double, SinRA As Double
Dim cosDecl     As Double, sinDecl As Double
Dim cosLon     As Double, sinLon As Double
CosObl = Cos(Obl)
SinObl = Sin(Obl)
CosRA = Cos(RA)
SinRA = Sin(RA)
cosDecl = Cos(Decl)
sinDecl = Sin(Decl)
Call SunLonAndEcc(T, Lon, E)
cosLon = Cos(Lon)
sinLon = Sin(Lon)

If (System = FK5System) Then '{ FK5 - include the e-terms }
    PI = (102.93735 + T * (1.71954 + T * 0.00046)) * DToR
    cosLon = cosLon - E * Cos(PI)
    sinLon = sinLon - E * Sin(PI)
End If

RA = RA - AbConst * (CosRA * cosLon * CosObl + SinRA * sinLon) / cosDecl
Tmp = cosLon * (SinObl * cosDecl - SinRA * sinDecl * CosObl) + CosRA * sinDecl * sinLon
Decl = Decl - AbConst * Tmp
End Sub

'(*****************************************************************************)
'(* Name:    EclAberration                                                    *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Corrects ecliptical coordinates for the effect of aberration.    *)
'(* Arguments:                                                                *)
'(*   T : Julian centuries since J2000                                        *)
'(*   System : either FK4System or FK5System                                  *)
'(*   l, b : coordinates to be corrected                                      *)
'(*****************************************************************************)

Sub EclAberration(T As Double, System As Long, ByRef l As Double, B As Double)

Dim Lon As Double, E As Double, PI As Double
Dim cosLon As Double, sinLon As Double
  Call SunLonAndEcc(T, Lon, E)
  cosLon = Cos(Lon)
  sinLon = Sin(Lon)
  If (System = FK5System) Then '{ FK5 - include the e-terms }
    PI = (102.93735 + T * (1.71954 + T * 0.00046)) * DToR
     cosLon = cosLon - E * Cos(PI)
     sinLon = sinLon - E * Sin(PI)
  End If
  B = B + AbConst * Sin(B) * (Sin(l) * cosLon - Cos(l) * sinLon)
  l = l - AbConst * Cos(l) * cosLon / Cos(B)
End Sub

