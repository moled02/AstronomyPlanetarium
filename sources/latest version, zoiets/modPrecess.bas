Attribute VB_Name = "modPrecess"
'(*****************************************************************************)
'(* Name:    PrecessFK5                                                       *)
'(* Type:    Procedure                                                        *)
'(* Purpose: precess equatorial coordinates from one FK5 epoch to another.    *)
'(* Arguments:                                                                *)
'(*   T0, T1 : initial and final epochs in centuries since J2000              *)
'(*   RA, Decl : coordinates to be converted                                  *)
'(*****************************************************************************)

Sub PrecessFK5(T0 As Double, t1 As Double, ByRef RA As Double, ByRef Decl As Double)
Dim T As Double, zeta As Double, Z As Double, Theta As Double
Dim A As Double, B As Double, C As Double

T = t1 - T0
Z = 2306.2181 + T0 * (1.39656 - T0 * 0.000139)
zeta = T * (Z + T * ((0.30188 - T0 * 0.000344) + T * 0.017998)) * SToR
Z = T * (Z + T * ((1.09468 + T0 * 0.000066) + T * 0.018203)) * SToR
Theta = (2004.3109 - T0 * (0.8533 + T0 * 0.000217))
Theta = T * (Theta - T * ((0.42665 + T0 * 0.000217) + T * 0.041833)) * SToR

A = Cos(Decl) * Sin(RA + zeta)
B = Cos(Theta) * Cos(Decl) * Cos(RA + zeta) - Sin(Theta) * Sin(Decl)
C = Sin(Theta) * Cos(Decl) * Cos(RA + zeta) + Cos(Theta) * Sin(Decl)

RA = atan2(A, B) + Z
If RA < 0 Then RA = RA + Pi2
Decl = asin(C)
End Sub

'(*****************************************************************************)
'(* Name:    PrecessFK4                                                       *)
'(* Type:    Procedure                                                        *)
'(* Purpose: precess equatorial coordinates from one FK4 epoch to another.    *)
'(* Arguments:                                                                *)
'(*   T0, T1 : initial and final epochs in centuries since J2000              *)
'(*   RA, Decl : coordinates to be converted                                  *)
'(*****************************************************************************)

Sub PrecessFK4(T0 As Double, t1 As Double, ByRef RA As Double, ByRef Decl As Double)

Dim T As Double, zeta As Double, Z As Double, Theta As Double
Dim A     As Double, B As Double, C As Double

Const TB1900 = (2415020.3135 - 2451545#) / 36525
Const JulianToBessel = (36525# / 36524.2199)

'{ Convert T values to units of tropical centuries since B1900.0 }
T = (t1 - T0) * JulianToBessel
T0 = (T0 - TB1900) * JulianToBessel

zeta = T * (2304.25 + T0 * 1.396 + T * (0.302 + T * 0.018)) * SToR
Z = zeta + T * T * (0.791 + T * 0.001) * SToR
Theta = T * (2004.682 - T0 * 0.853 - T * (0.426 + T * 0.042)) * SToR

A = Cos(Decl) * Sin(RA + zeta)
B = Cos(Theta) * Cos(Decl) * Cos(RA + zeta) - Sin(Theta) * Sin(Decl)
C = Sin(Theta) * Cos(Decl) * Cos(RA + zeta) + Cos(Theta) * Sin(Decl)

RA = atan2(A, B) + Z
If RA < 0 Then RA = RA + Pi2
Decl = asin(C)
End Sub

'(*****************************************************************************)
'(* Name:    EquinoxCorrection                                                *)
'(* Type:    Function                                                         *)
'(* Purpose: calculate the equinox correction from FK4 to FK5 system.         *)
'(* Arguments:                                                                *)
'(*   T : number of Julian centuries since J2000.0                            *)
'(* Return value:                                                             *)
'(*   the equinox correction in radians                                       *)
'(*****************************************************************************)

Function EquinoxCorrection(T As Double) As Double
EquinoxCorrection = (0.0775 + 0.085 * T) * 15 * SToR
End Function

Sub PrecessEcliptic(T0 As Double, t1 As Double, ByRef lambda As Double, ByRef beta As Double)
Dim T As Double
Dim tt As Double
Dim P As Double, pp As Double, a1 As Double, b1 As Double, C1 As Double, n As Double
Const TB1900 = (2415020.3135 - 2451545#) / 36525
Const JulianToBessel = (36525# / 36524.2199)

T = (TToJD(T0) - 2451545) / 36525#
tt = (TToJD(t1) - TToJD(T0)) / 36525

n = (47.0029 - 0.06603 * T + 0.000598 * T * T) * tt _
     + (-0.03302 + 0.000598 * T) * tt * tt + 0.00006 * tt * tt * tt
n = n / 3600
pp = 174.876384 + 3289.4789 * T + 0.60622 * T * T _
     - (869.8089 + 0.50491 * T) * tt + 0.03536 * tt * tt
P = (5029.0966 + 2.22226 * T - 0.000042 * T * T) * tt _
   + (1.11113 - 0.000042 * T) * tt * tt - 0.000006 * tt * tt * tt
P = P / 3600
a1 = cosd(n) * Cos(beta) * sind(pp - lambda / p11) - sind(n) * Sin(beta)
b1 = Cos(beta) * cosd(pp - lambda / p11)
C1 = cosd(n) * Sin(beta) + sind(n) * Cos(beta) * sind(pp - lambda / p11)
lambda = (P + pp) * p11 - atan2(a1, b1)
beta = asin(C1)
End Sub
