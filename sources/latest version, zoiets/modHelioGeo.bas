Attribute VB_Name = "modHelioGeo"
'(*****************************************************************************)
'(* Name:    HelioToGeo                                                       *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Convert heliocentric to geocentric coordinates.                  *)
'(* Arguments:                                                                *)
'(*   SHelio, SEarth : TSVECTOR record holding the object's and the Earth's   *)
'(*                    heliocentric ecliptical coordinates                    *)
'(*   SGeo : TSVECTOR to hold the geocentric ecliptical coordinates           *)
'(*****************************************************************************)

Sub HelioToGeo(SHelio As TSVECTOR, SEarth As TSVECTOR, ByRef SGeo As TSVECTOR)

Dim R As TVECTOR, REarth As TVECTOR
Dim I As Integer
Call SphToRect(SHelio, R)
Call SphToRect(SEarth, REarth)
R.x = R.x - REarth.x
R.Y = R.Y - REarth.Y
R.Z = R.Z - REarth.Z
Call RectToSph(R, SGeo)
End Sub

'(*****************************************************************************)
'(* Name:    RectHelioToGeo                                                   *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Convert heliocentric rectangular coordinates to geocentric       *)
'(*          spherical coordinates.                                           *)
'(* Arguments:                                                                *)
'(*   RHelio, RSun : TVECTOR arrays holding the object's and the Sun's        *)
'(*                  heliocentric equatorial coordinates                      *)
'(*   SGeo : TSVECTOR to hold the geocentric equatorial coordinates           *)
'(*****************************************************************************)

Sub RectHelioToGeo(RHelio As TVECTOR, RSun As TVECTOR, ByRef SGeo As TSVECTOR)

Dim R As TVECTOR
Dim tmp As Double
tmp = 0

R.x = RHelio.x + RSun.x: R.Y = RHelio.Y + RSun.Y: R.Z = RHelio.Z + RSun.Z
tmp = tmp + R.x * R.x + R.Y * R.Y + R.Z * R.Z

SGeo.R = Sqr(tmp)
SGeo.l = modpi2(atan2(R.Y, R.x))
SGeo.B = asin(R.Z / SGeo.R)
End Sub

