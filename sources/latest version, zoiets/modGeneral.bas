Attribute VB_Name = "modGeneral"

'*****************************************************************************)
'* Name:    EclToEqu                                                         *)
'* Type:    Procedure                                                        *)
'* Purpose: convert ecliptic coordinates to equatorial coordinates           *)
'* Arguments:                                                                *)
'*   l, b : ecliptical coordinates to be converted                           *)
'*   Obl : obliquity of the ecliptic                                         *)
'*   RA, Decl : the converted equatorial coordinates                         *)
'*****************************************************************************)

Sub EclToEqu(L As Double, B As Double, Obl As Double, ByRef RA As Double, ByRef Decl As Double)
Dim SinObl As Double, CosObl As Double
Dim sinl   As Double, cosl   As Double
Dim sinb   As Double, cosb   As Double

SinObl = Sin(Obl): CosObl = Cos(Obl)
sinl = Sin(L): cosl = Cos(L)
sinb = Sin(B): cosb = Cos(B)
RA = atan2(cosb * sinl * CosObl - sinb * SinObl, cosb * cosl)
If RA < 0 Then
    RA = RA + Pi2
End If
Decl = asin(sinb * CosObl + cosb * SinObl * sinl)
End Sub

'*****************************************************************************)
'* Name:    EquToEcl                                                         *)
'* Type:    Procedure                                                        *)
'* Purpose: convert equatorial coordinates to ecliptic coordinates           *)
'* Arguments:                                                                *)
'*   l, b : ecliptical coordinates to be converted                           *)
'*   Obl : obliquity of the ecliptic                                         *)
'*   RA, Decl : the converted equatorial coordinates                         *)
'*****************************************************************************)

Sub EquToEcl(RA As Double, Decl As Double, Obl As Double, ByRef L As Double, ByRef B As Double)
Dim SinObl As Double, CosObl As Double
Dim sinl   As Double, cosl   As Double
Dim sinb   As Double, cosb   As Double

SinObl = Sin(Obl): CosObl = Cos(Obl)
L = atan2(Sin(RA) * CosObl + tan(Decl) * SinObl, Cos(RA))
B = asin(Sin(Decl) * CosObl - Cos(Decl) * SinObl * Sin(RA))
If L < 0 Then
    L = L + Pi2
End If
End Sub

'(*****************************************************************************)
'(* Name:    EquToHor                                                         *)
'(* Type:    Procedure                                                        *)
'(* Purpose: convert equatorial coordinates to horizontal coordinates         *)
'(* Arguments:                                                                *)
'(*   RA, Decl : the equatorial coordinates to be converted                   *)
'(*   LST : Local Sidereal Time                                               *)
'(*   ObsLat : Observer's latitude                                            *)
'(*   Az, Alt : the converted horizontal coordinates                          *)
'(*****************************************************************************)

Sub EquToHor(RA As Double, Decl As Double, LST As Double, ObsLat As Double, ByRef Az As Double, ByRef Alt As Double)
Dim H As Double
H = LST - RA
Az = atan2(Sin(H), Cos(H) * Sin(ObsLat) - tan(Decl) * Cos(ObsLat))
Alt = asin(Sin(ObsLat) * Sin(Decl) + Cos(ObsLat) * Cos(Decl) * Cos(H))
End Sub

'(*****************************************************************************)
'(* Name:    AtmRefraction                                                    *)
'(* Type:    Function                                                         *)
'(* Purpose: calculate the atmospheric refraction at a given altitude         *)
'(* Arguments:                                                                *)
'(*   Altitude : the true altitude                                            *)
'(* Return value:                                                             *)
'(*   the apparent change in altitude due to atmospheric refraction           *)
'(*****************************************************************************)

Function AtmRefraction(Altitude As Double) As Double
Dim r As Double
r = 1.02 / tan(Altitude + 10.3 * DToR / (Altitude * RToD + 5.11))
AtmRefraction = (r + 0.0019279) * DToR / 60
End Function

'(*****************************************************************************)
'(* Name:    EquToGal                                                         *)
'(* Type:    Procedure                                                        *)
'(* Purpose: convert equatorial coordinates to galactic coordinates.          *)
'(* Arguments:                                                                *)
'(*   RA, Decl : the equatorial coordinates to be converted                   *)
'(*   l, b : the converted galactic coordinates                               *)
'(*****************************************************************************)

Sub EquToGal(RA As Double, Decl As Double, ByRef L As Double, ByRef B As Double)
RA = 192.25 * DToR - RA
L = 303 * DToR - atan2(Sin(RA), Cos(RA) * 0.4601 - tan(Decl) * 0.8878)
B = asin(0.4601 * Sin(Decl) + 0.8878 * Cos(Decl) * Cos(RA))
End Sub


'(*****************************************************************************)
'(* Name:    tan                                                              *)
'(* Type:    Function                                                         *)
'(* Purpose: trigonometric tangent function.                                  *)
'(*****************************************************************************)

Function tan(X As Double) As Double
tan = Sin(X) / Cos(X)
End Function

'(*****************************************************************************)
'(* Name:    asin                                                             *)
'(* Type:    Function                                                         *)
'(* Purpose: trigonometric inverse sine function.                             *)
'(*****************************************************************************)

Function asin(X As Double) As Double
If X = 1 Then
  asin = Pi / 2
ElseIf X = -1 Then
  asin = -Pi / 2
Else
  asin = Atn(X / Sqr(1 - X * X))
End If
End Function

'(*****************************************************************************)
'(* Name:    acos                                                             *)
'(* Type:    Function                                                         *)
'(* Purpose: trigonometric inverse cosine function.                           *)
'(*****************************************************************************)

Function acos(X As Double) As Double
acos = Pi / 2 - Atn(X / Sqr(1 - X * X))
End Function

'(*****************************************************************************)
'(* Name:    atan2                                                            *)
'(* Type:    Function                                                         *)
'(* Purpose: trigonometric 'complete' arctangent function.                    *)
'(* Return value:                                                             *)
'(*   angle between -pi and pi.                                               *)
'(*                                                                           *)
'(* Note: y is FIRST argument, x is SECOND argument.                          *)
'(*****************************************************************************)

Function atan2(Y As Double, X As Double) As Double
Dim Tmp As Double

If (X = 0) And (Y = 0) Then
  Tmp = 0
ElseIf (Abs(X) > Abs(Y)) Then
  Tmp = Atn(Y / X)
  If (X < 0) Then
    If (Tmp < 0) Then
      Tmp = Tmp + Pi
    Else
      Tmp = Tmp - Pi
    End If
  End If
Else
  Tmp = Pi / 2 - Atn(X / Y)
  If (Y < 0) Then
    Tmp = Tmp - Pi
  End If
End If
atan2 = Tmp
End Function

'(*****************************************************************************)
'(* Name:    log                                                              *)
'(* Type:    Function                                                         *)
'(* Purpose: logarithm to base 10.                                            *)
'(*****************************************************************************)

Function log10(X As Double) As Double
log10 = Log(X) / 2.302585093
End Function

'(*****************************************************************************)
'(* Name:    modpi2                                                           *)
'(* Type:    Function                                                         *)
'(* Purpose: reduce an angle to the interval (0, 2pi).                        *)
'(*****************************************************************************)

Function modpi2(X As Double) As Double
X = X - Int(X / Pi2) * Pi2
If (X < 0) Then
  X = X + Pi2
End If
modpi2 = X
End Function

'(*****************************************************************************)
'(* Name:    modpi2                                                           *)
'(* Type:    Function                                                         *)
'(* Purpose: reduce an angle to the interval (-pi, pi).                       *)
'(*****************************************************************************)

Function modpi(X As Double) As Double
X = X - Int(X / Pi2) * Pi2
If (X < -Pi) Then
  X = X + Pi2
ElseIf (X > Pi) Then
  X = X - Pi2
End If
modpi = X
End Function
'(*****************************************************************************)
'(* Name:    floor                                                            *)
'(* Type:    Function                                                         *)
'(* Purpose: largest integer not greater than a real number                   *)
'(*****************************************************************************)

Function floor(X As Double) As Double
Dim Tmp As Double
Tmp = Fix(X)
If (X < 0) Then
  If (X = Tmp) Then
    floor = Tmp
  Else
    floor = Tmp - 1
  End If
Else
  floor = Tmp
End If
End Function

'(*****************************************************************************)
'(* Name:    EclToRect                                                        *)
'(* Type:    Procedure                                                        *)
'(* Purpose: convert from ecliptical to rectangular coordinates               *)
'(* Arguments:                                                                *)
'(*   S : TSVECTOR record holding the spherical coordinates                   *)
'(*   obl: Obliquity                                                          *)
'(*   R : TVECTOR array to hold the rectangular coordinates                   *)
'(*****************************************************************************)

Sub EclToRect(s As TSVECTOR, Obl As Double, ByRef p As TVECTOR)
With s
  p.X = s.r * Cos(s.L) * Cos(s.B)
  p.Y = s.r * (Sin(s.L) * Cos(s.B) * Cos(Obl) - Sin(s.B) * Sin(Obl))
  p.Z = s.r * (Sin(s.L) * Cos(s.B) * Sin(Obl) + Sin(s.B) * Cos(Obl))
End With
End Sub

'(*****************************************************************************)
'(* Name:    SphToRect                                                        *)
'(* Type:    Procedure                                                        *)
'(* Purpose: convert from spherical to rectangular coordinates                *)
'(* Arguments:                                                                *)
'(*   S : TSVECTOR record holding the spherical coordinates                   *)
'(*   R : TVECTOR array to hold the rectangular coordinates                   *)
'(*****************************************************************************)

Sub SphToRect(s As TSVECTOR, ByRef p As TVECTOR)
With s
  p.X = s.r * Cos(s.L) * Cos(s.B)
  p.Y = s.r * Sin(s.L) * Cos(s.B)
  p.Z = s.r * Sin(s.B)
End With
End Sub

'(*****************************************************************************)
'(* Name:    RectToSph                                                        *)
'(* Type:    Procedure                                                        *)
'(* Purpose: convert from rectangular to spherical coordinates                *)
'(* Arguments:                                                                *)
'(*   R : TVECTOR array holding the rectangular coordinates                   *)
'(*   S : TSVECTOR record to hold the spherical coordinates                   *)
'(*****************************************************************************)

Sub RectToSph(p As TVECTOR, ByRef s As TSVECTOR)
s.L = atan2(p.Y, p.X)
If s.L < 0 Then s.L = s.L + Pi2
s.r = Sqr(p.X * p.X + p.Y * p.Y + p.Z * p.Z)
s.B = asin(p.Z / s.r)
End Sub


Function Interpol3(y1 As Double, y2 As Double, y3 As Double, n As Double) As Double
Dim A As Double, B As Double, C As Double
A = y2 - y1
B = y3 - y2
C = B - A
Interpol3 = y2 + n / 2 * (A + B + n * C)
End Function

Function Inv_interpol3(y1 As Double, y2 As Double, y3 As Double) As Double
Dim A As Double, B As Double, C As Double, n As Double, n1 As Double

n = 0
A = y2 - y1
B = y3 - y2
C = B - A
n1 = -2 * y2 / (A + B + n * C)
While Abs(n1 - n) > 0.00001
    n = n1
    n1 = -2 * y2 / (A + B + n * C)
Wend
Inv_interpol3 = n1
End Function

Function Nulpunt3(ByVal y1 As Double, ByVal y2 As Double, ByVal y3 As Double, ByRef n As Double, ByVal nStap As Double)
Dim yx As Double

If Abs(y2) < 0.00001 Then
    Exit Function
End If
If sign(y1) <> sign(y2) Then
    n = n - nStap / 2
    yx = Interpol3(y1, y2, y3, -0.5)
    Call Nulpunt(y1, yx, y2, n, nStap / 2)
Else
    n = n + nStap / 2
    yx = Interpol3(y1, y2, y3, 0.5)
    Call Nulpunt(y2, yx, y3, n, nStap / 2)
End If
End Function

Function Interpol5(y1 As Double, y2 As Double, y3 As Double, y4 As Double, y5 As Double, n As Double) As Double
Dim A As Double, B As Double, C As Double, D As Double, E As Double, F As Double, g As Double, H As Double, j As Double, k As Double
A = y2 - y1: B = y3 - y2: C = y4 - y3: D = y5 - y4
E = B - A: F = C - B: g = D - C
H = F - E: j = g - F
k = j - H
Interpol5 = y3 + n / 2 * (B + C) + n * n / 2 * F + n * (n * n - 1) / 12 * (H + j) + n * n * (n * n - 1) / 24 * k
End Function

Function Extreem5(y1 As Double, y2 As Double, y3 As Double, y4 As Double, y5 As Double) As Double
Dim A As Double, B As Double, C As Double, D As Double, E As Double, F As Double, g As Double, H As Double, j As Double, k As Double
Dim n As Double, n1 As Double
A = y2 - y1: B = y3 - y2: C = y4 - y3: D = y5 - y4
E = B - A: F = C - B: g = D - C
H = F - E: j = g - F
k = j - H
n1 = (63 + 6 * C - H - j) / (k - 12 * F)
While Abs(n1 - n) > 0.00001
    n = (63 + 6 * C - H - j + 3 * n1 + 2 * n1 * n1 * n1 * k) / (k - 12 * F)
Wend
Extreem5 = n1
End Function

Function Inv_interpol5(y1 As Double, y2 As Double, y3 As Double, y4 As Double, y5 As Double) As Double
Dim A As Double, B As Double, C As Double, D As Double, E As Double, F As Double, g As Double, H As Double, j As Double, k As Double
Dim n As Double, n1 As Double
A = y2 - y1: B = y3 - y2: C = y4 - y3: D = y5 - y4
E = B - A: F = C - B: g = D - C
H = F - E: j = g - F
k = j - H
n1 = -24 * y3 / (2 * (63 + 6 * C - H - j))
While Abs(n1 - n) > 0.00001
    n = (-24 * y3 + n1 * n1 * (k - 12 * F) - 2 * n1 * n1 * n1 * (H + j) - n1 * n1 * n1 * n1 * k) / (2 * (63 + 6 * C - H - j))
Wend
Inv_interpol5 = n1
End Function
Function Nulpunt5(ByVal y1 As Double, ByVal y2 As Double, ByVal y3 As Double, ByVal y4 As Double, ByVal y5 As Double)
Dim A As Double, B As Double, C As Double, D As Double, E As Double, F As Double, g As Double, H As Double, j As Double, k As Double
Dim M As Double, n As Double, p As Double, Q As Double
Dim n0 As Double
Dim dn0 As Double
A = y2 - y1: B = y3 - y2: C = y4 - y3: D = y5 - y4
E = B - A: F = C - B: g = D - C
H = F - E: j = g - F
k = j - H
M = k / 24: n = (H + j) / 12: p = F / 2 - M: Q = (B + C) / 2 - n
n0 = 0
dn0 = -(M * n0 * n0 * n0 * n0 + n * n0 * n0 * n0 + p * n0 * n0 + Q * n0 + y3) / (4 * M * n0 * n0 * n0 + 3 * M * n0 * n0 + 2 * p * n0 + Q)
While Abs(dn0) > 0.00001
    n0 = n0 + dn0
    dn0 = -(M * n0 * n0 * n0 * n0 + n * n0 * n0 * n0 + p * n0 * n0 + Q * n0 + y3) / (4 * M * n0 * n0 * n0 + 3 * M * n0 * n0 + 2 * p * n0 + Q)
Wend
Nulpunt5 = n0
End Function

Function Eval2Poly(p As T2POLY, X As Double) As Double
Eval2Poly = p.p(0) + X * (p.p(1) + X * p.p(2))
End Function

Function Eval3Poly(p As T3POLY, X As Double) As Double
Eval3Poly = p.p(0) + X * (p.p(1) + X * (p.p(2) + X * p.p(3)))
End Function

Function Eval4Poly(p As T4POLY, X As Double) As Double
Dim i As Integer
Dim r As Double
r = p.p(4)
For i = 3 To 0 Step -1
  r = r * X + p.p(i)
Next
Eval4Poly = r
End Function

'(*****************************************************************************)
'(* Name:    CalcSinCosTab                                                    *)
'(* Type:    Procedure                                                        *)
'(* Purpose: calculate a table of sines and cosines of multiples of an anlge. *)
'(* Arguments:                                                                *)
'(*   X : the base angle                                                      *)
'(*   Degree : the highest multiple for which to calculate sine and cosine    *)
'(*   SinTab, CosTab : TSINCOSTAB arrays to hold the return values            *)
'(*                                                                           *)
'(* Note: Degree must be less than or equal to 6.                             *)
'(*****************************************************************************)

Public Sub CalcSinCosTab(ByVal X As Double, ByVal Degree As Long, ByRef SinTab As TSINCOSTAB, ByRef CosTab As TSINCOSTAB)

Dim SinVal As Double, CosVal As Double
Dim i As Long

SinVal = Sin(X)
SinTab.W(1) = SinVal
CosVal = Cos(X)
CosTab.W(1) = CosVal
For i = 2 To Degree
    SinTab.W(i) = SinTab.W(i - 1) * CosVal + CosTab.W(i - 1) * SinVal
    CosTab.W(i) = CosTab.W(i - 1) * CosVal - SinTab.W(i - 1) * SinVal
Next
End Sub

Sub Reduction2000(T0 As Double, ByRef RA As Double, ByRef Decl As Double)

Dim X As Double, Y As Double, Z As Double, t1 As Double, t2 As Double, A As Double
Dim xx As Double, yx As Double, xy As Double, zx As Double, xz As Double, yy As Double, yz As Double, zy As Double, zz As Double
Dim x0 As Double, y0 As Double, z0 As Double, JD     As Double

'{t2 = startperiode,
' t1 = eindperiode : Vanaf nu naar 2000}

    X = Cos(RA) * Cos(Decl)
    Y = Sin(RA) * Cos(Decl)
    Z = Sin(Decl)
    JD = TToJD(T0)
    t1 = (JD - 2451545#) / 36524.21988
    t2 = (JD - 2433282.4235) / 36524.21988
    A = 0.00000001
    xx = A * (-(29696 + 26 * t2) * t1 * t1 - 13 * t1 * t1 * t1) + 1
    yx = A * (-(2234941 + 1355 * t2) * t1 - 676 * t1 * t1 + 221 * t1 * t1 * t1)
    xy = -yx
    zx = A * (-(971690 - 414 * t2) * t1 + 207 * t1 * t1 + 96 * t1 * t1 * t1)
    xz = -zx
    yy = A * (-(24975 + 30 * t2) * t1 * t1 - 15 * t1 * t1 * t1) + 1
    yz = A * (-(10858 + 2 * t2) * t1 * t1)
    zy = yz
    zz = A * (-(4721 - 4 * t2) * t1 * t1) + 1
    x0 = xx * X + xy * Y + xz * Z
    y0 = yx * X + yy * Y + yz * Z
    z0 = zx * X + zy * Y + zz * Z
    RA = atan2(y0, x0)
    If RA < 0 Then RA = RA + 2 * Pi
    Decl = Atn(z0 / (Sqr(x0 * x0 + y0 * y0)))
End Sub

Sub EclVSOP2000_equFK52000(ByRef X As Double, ByRef Y As Double, ByRef Z As Double)
Dim x0 As Double
Dim y0 As Double
Dim z0 As Double

x0 = X + 0.00000044036 * Y - 0.000000190919 * Z
y0 = -0.000000479966 * X + 0.917482138087 * Y - 0.397776982902 * Z
z0 = 0.397776982902 * Y + 0.917482137087 * Z
X = x0: Y = y0: Z = z0
End Sub
Function C2P(p As Variant) As T2POLY
Dim i As Long
Dim aPoly2 As T2POLY

For i = 0 To 2
    aPoly2.p(i) = p(i)
Next
C2P = aPoly2
End Function

Function C3P(p As Variant) As T3POLY
Dim i As Long
Dim aPoly3 As T3POLY

For i = 0 To 3
    aPoly3.p(i) = p(i)
Next
C3P = aPoly3
End Function

Function C4P(p As Variant) As T4POLY
Dim i As Long
Dim aPoly4 As T4POLY

For i = 0 To 4
    aPoly4.p(i) = p(i)
Next
C4P = aPoly4
End Function

Function sind(X As Double) As Double
sind = Sin(X * p11)
End Function
Function cosd(X As Double) As Double
cosd = Cos(X * p11)
End Function
Function atan2d(Y As Double, X As Double) As Double
    atan2d = atan2(Y, X) / p11
End Function
Function asind(X As Double) As Double
    asind = asin(X) / p11
End Function
Function acosd(X As Double) As Double
    acosd = acos(X) / p11
End Function
Function tand(X As Double) As Double
    tand = tan(X * p11)
End Function
