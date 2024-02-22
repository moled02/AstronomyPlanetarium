Attribute VB_Name = "modParallax"
'(*****************************************************************************)
'(* Name:    ParallaxHi                                                       *)
'(* Type:    Procedure                                                        *)
'(* Purpose: correct right ascension and declination for the effect of        *)
'(*          parallax using the rigorous formulae.                            *)
'(* Arguments:                                                                *)
'(*   sinParallax : sine of the object's parallax                             *)
'(*   LST : the Local Sidereal Time                                           *)
'(*   RhoCosPhi,RhoSinPhi : the observer's geocentric rectangular coordinates *)
'(*   RA, Decl : the coordinates to be corrected                              *)
'(*****************************************************************************)

Sub ParallaxHi(ByVal sinParallax As Double, ByVal LST As Double, ByVal RhoCosPhi As Double, ByVal RhoSinPhi As Double, ByRef RA As Double, ByRef Decl As Double)
Dim H As Double
Dim D As Double, n As Double, dRA As Double
H = LST - RA
D = -RhoCosPhi * sinParallax * Sin(H)
n = Cos(Decl) - RhoCosPhi * sinParallax * Cos(H)
dRA = Atn(D / n)
RA = RA + dRA
D = (Sin(Decl) - RhoSinPhi * sinParallax) * Cos(dRA)
Decl = Atn(D / n)
End Sub

'(*****************************************************************************)
'(* Name:    Parallax                                                         *)
'(* Type:    Procedure                                                        *)
'(* Purpose: correct right ascension and declination for the effect of        *)
'(*          parallax using the simplified formulae.                          *)
'(* Arguments:                                                                *)
'(*   sinParallax : sine of the object's parallax                             *)
'(*   LST : the Local Sidereal Time                                           *)
'(*   RhoCosPhi,RhoSinPhi : the observer's geocentric rectangular coordinates *)
'(*   RA, Decl : the coordinates to be corrected                              *)
'(*****************************************************************************)

Sub Parallax(ByVal Parallax As Double, ByVal LST As Double, ByVal RhoCosPhi As Double, ByVal RhoSinPhi As Double, ByRef RA As Double, ByRef Decl As Double)
Dim H As Double
H = LST - RA
RA = RA - Parallax * RhoCosPhi * Sin(H) / Cos(Decl)
Decl = Decl - Parallax * (RhoSinPhi * Cos(Decl) - RhoCosPhi * Cos(H) * Sin(Decl))
End Sub

'(*****************************************************************************)
'(* Name:    ParallaxHiAlt                                                    *)
'(* Type:    Procedure                                                        *)
'(* Purpose: correct right ascension, declination and semi-diameter for       *)
'(*          parallax using the alternative method.                           *)
'(* Arguments:                                                                *)
'(*   sinParallax : sine of the object's parallax                             *)
'(*   LST : the Local Sidereal Time                                           *)
'(*   RhoCosPhi,RhoSinPhi : the observer's geocentric rectangular coordinates *)
'(*   RA, Decl, SD : the coordinates and semi-diameter to be corrected        *)
'(*****************************************************************************)

Sub ParallaxHiAlt(sinParallax As Double, LST As Double, RhoCosPhi As Double, RhoSinPhi As Double, ByRef RA As Double, ByRef Decl As Double, ByRef SD As Double)
Dim H As Double
Dim A   As Double, B As Double, C As Double, Q As Double
H = LST - RA
A = Cos(Decl) * Sin(H)
B = Cos(Decl) * Cos(H) - RhoCosPhi * sinParallax
C = Sin(Decl) - RhoSinPhi * sinParallax
Q = Sqr(A * A + B * B + C * C)

RA = modpi2(LST - atan2(A, B))
Decl = asin(C / Q)
SD = SD / Q
End Sub

