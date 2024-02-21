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

Sub EclToEqu(l As Double, B As Double, Obl As Double, ByRef RA As Double, ByRef Decl As Double)
Dim SinObl As Double, CosObl As Double
Dim sinl   As Double, cosl   As Double
Dim sinb   As Double, cosb   As Double

SinObl = Sin(Obl): CosObl = Cos(Obl)
sinl = Sin(l): cosl = Cos(l)
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

Sub EquToEcl(RA As Double, Decl As Double, Obl As Double, ByRef l As Double, ByRef B As Double)
Dim SinObl As Double, CosObl As Double
Dim sinl   As Double, cosl   As Double
Dim sinb   As Double, cosb   As Double

SinObl = Sin(Obl): CosObl = Cos(Obl)
l = atan2(Sin(RA) * CosObl + tan(Decl) * SinObl, Cos(RA))
B = asin(Sin(Decl) * CosObl - Cos(Decl) * SinObl * Sin(RA))
If l < 0 Then
    l = l + Pi2
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

Sub EquToGal(RA As Double, Decl As Double, ByRef l As Double, ByRef B As Double)
RA = 192.25 * DToR - RA
l = 303 * DToR - atan2(Sin(RA), Cos(RA) * 0.4601 - tan(Decl) * 0.8878)
B = asin(0.4601 * Sin(Decl) + 0.8878 * Cos(Decl) * Cos(RA))
End Sub


'(*****************************************************************************)
'(* Name:    tan                                                              *)
'(* Type:    Function                                                         *)
'(* Purpose: trigonometric tangent function.                                  *)
'(*****************************************************************************)

Function tan(x As Double) As Double
tan = Sin(x) / Cos(x)
End Function

'(*****************************************************************************)
'(* Name:    asin                                                             *)
'(* Type:    Function                                                         *)
'(* Purpose: trigonometric inverse sine function.                             *)
'(*****************************************************************************)

Function asin(x As Double) As Double
If x = 1 Then
  asin = PI / 2
ElseIf x = -1 Then
  asin = -PI / 2
Else
  asin = Atn(x / Sqr(1 - x * x))
End If
End Function

'(*****************************************************************************)
'(* Name:    acos                                                             *)
'(* Type:    Function                                                         *)
'(* Purpose: trigonometric inverse cosine function.                           *)
'(*****************************************************************************)

Function acos(x As Double) As Double
acos = PI / 2 - Atn(x / Sqr(1 - x * x))
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

Function atan2(Y As Double, x As Double) As Double
Dim Tmp As Double

If (x = 0) And (Y = 0) Then
  Tmp = 0
ElseIf (Abs(x) > Abs(Y)) Then
  Tmp = Atn(Y / x)
  If (x < 0) Then
    If (Tmp < 0) Then
      Tmp = Tmp + PI
    Else
      Tmp = Tmp - PI
    End If
  End If
Else
  Tmp = PI / 2 - Atn(x / Y)
  If (Y < 0) Then
    Tmp = Tmp - PI
  End If
End If
atan2 = Tmp
End Function

'(*****************************************************************************)
'(* Name:    log                                                              *)
'(* Type:    Function                                                         *)
'(* Purpose: logarithm to base 10.                                            *)
'(*****************************************************************************)

Function log10(x As Double) As Double
log10 = Log(x) / 2.302585093
End Function

'(*****************************************************************************)
'(* Name:    modpi2                                                           *)
'(* Type:    Function                                                         *)
'(* Purpose: reduce an angle to the interval (0, 2pi).                        *)
'(*****************************************************************************)

Function modpi2(x As Double) As Double
x = x - Int(x / Pi2) * Pi2
If (x < 0) Then
  x = x + Pi2
End If
modpi2 = x
End Function

'(*****************************************************************************)
'(* Name:    modpi2                                                           *)
'(* Type:    Function                                                         *)
'(* Purpose: reduce an angle to the interval (-pi, pi).                       *)
'(*****************************************************************************)

Function modpi(x As Double) As Double
x = x - Int(x / Pi2) * Pi2
If (x < -PI) Then
  x = x + Pi2
ElseIf (x > PI) Then
  x = x - Pi2
End If
modpi = x
End Function
'(*****************************************************************************)
'(* Name:    floor                                                            *)
'(* Type:    Function                                                         *)
'(* Purpose: largest integer not greater than a real number                   *)
'(*****************************************************************************)

Function floor(x As Double) As Double
Dim Tmp As Double
Tmp = Fix(x)
If (x < 0) Then
  If (x = Tmp) Then
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

Sub EclToRect(s As TSVECTOR, Obl As Double, ByRef P As TVECTOR)
With s
  P.x = s.r * Cos(s.l) * Cos(s.B)
  P.Y = s.r * (Sin(s.l) * Cos(s.B) * Cos(Obl) - Sin(s.B) * Sin(Obl))
  P.Z = s.r * (Sin(s.l) * Cos(s.B) * Sin(Obl) + Sin(s.B) * Cos(Obl))
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

Sub SphToRect(s As TSVECTOR, ByRef P As TVECTOR)
With s
  P.x = s.r * Cos(s.l) * Cos(s.B)
  P.Y = s.r * Sin(s.l) * Cos(s.B)
  P.Z = s.r * Sin(s.B)
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

Sub RectToSph(P As TVECTOR, ByRef s As TSVECTOR)
s.l = atan2(P.Y, P.x)
If s.l < 0 Then s.l = s.l + Pi2
s.r = Sqr(P.x * P.x + P.Y * P.Y + P.Z * P.Z)
s.B = asin(P.Z / s.r)
End Sub


Function Interpol3(y1 As Double, y2 As Double, y3 As Double, N As Double) As Double
Dim A As Double, B As Double, C As Double
A = y2 - y1
B = y3 - y2
C = B - A
Interpol3 = y2 + N / 2 * (A + B + N * C)
End Function

Function Inv_interpol3(y1 As Double, y2 As Double, y3 As Double) As Double
Dim A As Double, B As Double, C As Double, N As Double, n1 As Double

N = 0
A = y2 - y1
B = y3 - y2
C = B - A
n1 = -2 * y2 / (A + B + N * C)
While Abs(n1 - N) > 0.00001
    N = n1
    n1 = -2 * y2 / (A + B + N * C)
Wend
Inv_interpol3 = n1
End Function

Function Nulpunt3(ByVal y1 As Double, ByVal y2 As Double, ByVal y3 As Double, ByRef N As Double, ByVal nStap As Double)
Dim yx As Double

If Abs(y2) < 0.00001 Then
    Exit Function
End If
If sign(y1) <> sign(y2) Then
    N = N - nStap / 2
    yx = Interpol3(y1, y2, y3, -0.5)
    Call Nulpunt(y1, yx, y2, N, nStap / 2)
Else
    N = N + nStap / 2
    yx = Interpol3(y1, y2, y3, 0.5)
    Call Nulpunt(y2, yx, y3, N, nStap / 2)
End If
End Function

Function Interpol5(y1 As Double, y2 As Double, y3 As Double, y4 As Double, y5 As Double, N As Double) As Double
Dim A As Double, B As Double, C As Double, D As Double, E As Double, F As Double, G As Double, H As Double, J As Double, K As Double
A = y2 - y1: B = y3 - y2: C = y4 - y3: D = y5 - y4
E = B - A: F = C - B: G = D - C
H = F - E: J = G - F
K = J - H
Interpol5 = y3 + N / 2 * (B + C) + N * N / 2 * F + N * (N * N - 1) / 12 * (H + J) + N * N * (N * N - 1) / 24 * K
End Function

Function Extreem5(y1 As Double, y2 As Double, y3 As Double, y4 As Double, y5 As Double) As Double
Dim A As Double, B As Double, C As Double, D As Double, E As Double, F As Double, G As Double, H As Double, J As Double, K As Double
Dim N As Double, n1 As Double
A = y2 - y1: B = y3 - y2: C = y4 - y3: D = y5 - y4
E = B - A: F = C - B: G = D - C
H = F - E: J = G - F
K = J - H
n1 = (63 + 6 * C - H - J) / (K - 12 * F)
While Abs(n1 - N) > 0.00001
    N = (63 + 6 * C - H - J + 3 * n1 + 2 * n1 * n1 * n1 * K) / (K - 12 * F)
Wend
Extreem5 = n1
End Function

Function Inv_interpol5(y1 As Double, y2 As Double, y3 As Double, y4 As Double, y5 As Double) As Double
Dim A As Double, B As Double, C As Double, D As Double, E As Double, F As Double, G As Double, H As Double, J As Double, K As Double
Dim N As Double, n1 As Double
A = y2 - y1: B = y3 - y2: C = y4 - y3: D = y5 - y4
E = B - A: F = C - B: G = D - C
H = F - E: J = G - F
K = J - H
n1 = -24 * y3 / (2 * (63 + 6 * C - H - J))
While Abs(n1 - N) > 0.00001
    N = (-24 * y3 + n1 * n1 * (K - 12 * F) - 2 * n1 * n1 * n1 * (H + J) - n1 * n1 * n1 * n1 * K) / (2 * (63 + 6 * C - H - J))
Wend
Inv_interpol5 = n1
End Function
Function Nulpunt5(ByVal y1 As Double, ByVal y2 As Double, ByVal y3 As Double, ByVal y4 As Double, ByVal y5 As Double)
Dim A As Double, B As Double, C As Double, D As Double, E As Double, F As Double, G As Double, H As Double, J As Double, K As Double
Dim M As Double, N As Double, P As Double, Q As Double
Dim n0 As Double
Dim dn0 As Double
A = y2 - y1: B = y3 - y2: C = y4 - y3: D = y5 - y4
E = B - A: F = C - B: G = D - C
H = F - E: J = G - F
K = J - H
M = K / 24: N = (H + J) / 12: P = F / 2 - M: Q = (B + C) / 2 - N
n0 = 0
dn0 = -(M * n0 * n0 * n0 * n0 + N * n0 * n0 * n0 + P * n0 * n0 + Q * n0 + y3) / (4 * M * n0 * n0 * n0 + 3 * M * n0 * n0 + 2 * P * n0 + Q)
While Abs(dn0) > 0.00001
    n0 = n0 + dn0
    dn0 = -(M * n0 * n0 * n0 * n0 + N * n0 * n0 * n0 + P * n0 * n0 + Q * n0 + y3) / (4 * M * n0 * n0 * n0 + 3 * M * n0 * n0 + 2 * P * n0 + Q)
Wend
Nulpunt5 = n0
End Function

Function Eval2Poly(P As T2POLY, x As Double) As Double
Eval2Poly = P.P(0) + x * (P.P(1) + x * P.P(2))
End Function

Function Eval3Poly(P As T3POLY, x As Double) As Double
Eval3Poly = P.P(0) + x * (P.P(1) + x * (P.P(2) + x * P.P(3)))
End Function

Function Eval4Poly(P As T4POLY, x As Double) As Double
Dim i As Integer
Dim r As Double
r = P.P(4)
For i = 3 To 0 Step -1
  r = r * x + P.P(i)
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

Public Sub CalcSinCosTab(ByVal x As Double, ByVal Degree As Long, ByRef SinTab As TSINCOSTAB, ByRef CosTab As TSINCOSTAB)

Dim SinVal As Double, CosVal As Double
Dim i As Long

SinVal = Sin(x)
SinTab.W(1) = SinVal
CosVal = Cos(x)
CosTab.W(1) = CosVal
For i = 2 To Degree
    SinTab.W(i) = SinTab.W(i - 1) * CosVal + CosTab.W(i - 1) * SinVal
    CosTab.W(i) = CosTab.W(i - 1) * CosVal - SinTab.W(i - 1) * SinVal
Next
End Sub

Sub Reduction2000(T0 As Double, ByRef RA As Double, ByRef Decl As Double)

Dim x As Double, Y As Double, Z As Double, T1 As Double, t2 As Double, A As Double
Dim xx As Double, yx As Double, xy As Double, zx As Double, xz As Double, yy As Double, yz As Double, zy As Double, zz As Double
Dim x0 As Double, y0 As Double, z0 As Double, JD     As Double

'{t2 = startperiode,
' t1 = eindperiode : Vanaf nu naar 2000}

    x = Cos(RA) * Cos(Decl)
    Y = Sin(RA) * Cos(Decl)
    Z = Sin(Decl)
    JD = TToJD(T0)
    T1 = (JD - 2451545#) / 36524.21988
    t2 = (JD - 2433282.4235) / 36524.21988
    A = 0.00000001
    xx = A * (-(29696 + 26 * t2) * T1 * T1 - 13 * T1 * T1 * T1) + 1
    yx = A * (-(2234941 + 1355 * t2) * T1 - 676 * T1 * T1 + 221 * T1 * T1 * T1)
    xy = -yx
    zx = A * (-(971690 - 414 * t2) * T1 + 207 * T1 * T1 + 96 * T1 * T1 * T1)
    xz = -zx
    yy = A * (-(24975 + 30 * t2) * T1 * T1 - 15 * T1 * T1 * T1) + 1
    yz = A * (-(10858 + 2 * t2) * T1 * T1)
    zy = yz
    zz = A * (-(4721 - 4 * t2) * T1 * T1) + 1
    x0 = xx * x + xy * Y + xz * Z
    y0 = yx * x + yy * Y + yz * Z
    z0 = zx * x + zy * Y + zz * Z
    RA = atan2(y0, x0)
    If RA < 0 Then RA = RA + 2 * PI
    Decl = Atn(z0 / (Sqr(x0 * x0 + y0 * y0)))
End Sub

Sub EclVSOP2000_equFK52000(ByRef x As Double, ByRef Y As Double, ByRef Z As Double)
Dim x0 As Double
Dim y0 As Double
Dim z0 As Double

x0 = x + 0.00000044036 * Y - 0.000000190919 * Z
y0 = -0.000000479966 * x + 0.917482138087 * Y - 0.397776982902 * Z
z0 = 0.397776982902 * Y + 0.917482137087 * Z
x = x0: Y = y0: Z = z0
End Sub
Function C2P(P As Variant) As T2POLY
Dim i As Long
Dim aPoly2 As T2POLY

For i = 0 To 2
    aPoly2.P(i) = P(i)
Next
C2P = aPoly2
End Function

Function C3P(P As Variant) As T3POLY
Dim i As Long
Dim aPoly3 As T3POLY

For i = 0 To 3
    aPoly3.P(i) = P(i)
Next
C3P = aPoly3
End Function

Function C4P(P As Variant) As T4POLY
Dim i As Long
Dim aPoly4 As T4POLY

For i = 0 To 4
    aPoly4.P(i) = P(i)
Next
C4P = aPoly4
End Function

