Attribute VB_Name = "modPhys"
Type TMOONPHYSDATA
    l As Double
    B As Double '          { Total libration                       }
    ld As Double
    Bd As Double '        { Optical libration                     }
    ldd As Double
    bdd As Double '      { Physical libration                    }
    P  As Double '                      { Position angle                        }
    i As Double
    k As Double '          { Phase angle and illuminated fraction  }
    x As Double '                       { Position angle of the bright limb     }
    l0 As Double '                      { Selenographical longitude             }
    b0  As Double '                     { Selenographical latitude              }
    T As Double '                       { Terminator angle                      }
    End Type
Type TSUNPHYSDATA
    P  As Double '                      { Position angle                      }
    b0 As Double
    l0 As Double '        { Heliographic longitude and latitude }
    End Type
Type TMARSPHYSDATA
    DS As Double
    DE As Double  '       { Planetocentric declination of Sun and Earth }
    Om As Double '                      { Longitude of the central meridian           }
    P As Double '                       { Position angle of the axis                  }
    Q As Double
    qq   As Double '       { Position angle and -defect of illumination  }
    D As Double '                       { Angular diameter in seconds of arc          }
    End Type
Type TJUPITERPHYSDATA
    DS As Double
    DE As Double  '       { Planetocentric declination of Sun and Earth   }
    Om1 As Double
    Om2 As Double '      { Longitude of the central meridians (I and II) }
    C As Double '                       { Correction for phase to the above             }
    P As Double '                       { Position angle of the axis                    }
    End Type
Type TSATURNRINGDATA
    B As Double
    Bd As Double '         { Planetocentric declination of Earth and Sun }
    P As Double '                       { Position angle of the axis                  }
    DeltaU  As Double '                 { Needed for magnitude calculation            }
    aAxis As Double
    bAxis As Double      '{ Length of axes of (outer edge of outer ring)/rings in arcseconds  }
    ioaAxis As Double
    iobAxis As Double  '{ Length of axes of inner edge of outer ring in arcseconds          }
    oiaAxis As Double
    oibAxis As Double  '{ Length of axes of outer edge of inner ring in arcseconds          }
    iiaAxis As Double
    iibAxis As Double  '{ Length of axes of inner edge of inner ring in arcseconds          }
    idaAxis As Double
    idbAxis As Double  '{ Length of axes of inner edge of dusky ring in arcseconds          }
    End Type
Type TALTSATURNRINGDATA
    u As Double
    U1 As Double  '       { geoc/helioc longitude Saturn                }
    B As Double
    b1 As Double  '       { Saturnicentrische latitude of Earth/Sun     }
    P As Double
    P1 As Double  '       { geoc./helioc position angle of the axis     }
    j As Double
    n As Double
    W As Double   '     { Incl, RK stijg. knoop, arg. stijg. knoop    }
  End Type

'(*****************************************************************************)
'(* Name:    CalcPhase                                                        *)
'(* Type:    Function                                                         *)
'(* Purpose: Calculate phase of a planet from interplanetary distances.       *)
'(* Arguments:                                                                *)
'(*   rPS, rES, rPE : distance Planet-Sun, Earth-Sun and Planet-Earth,        *)
'(*                   respectively                                            *)
'(* Return value:                                                             *)
'(*   the phase of the planet                                                 *)
'(*****************************************************************************)

Function CalcPhase(rPS As Double, rES As Double, rPE As Double) As Double
CalcPhase = (1 + (rPS * rPS + rPE * rPE - rES * rES) / (2 * rPS * rPE)) / 2
End Function

'(*****************************************************************************)
'(* Name:    CalcPhaseAngle                                                   *)
'(* Type:    Function                                                         *)
'(* Purpose: Calculate phase angle from interplanetary distances.             *)
'(* Arguments:                                                                *)
'(*   rPS, rES, rPE : distance Planet-Sun, Earth-Sun and Planet-Earth,        *)
'(*                   respectively                                            *)
'(* Return value:                                                             *)
'(*   the phase angle                                                         *)
'(*****************************************************************************)

Function CalcPhaseAngle(rPS As Double, rES As Double, rPE As Double) As Double
CalcPhaseAngle = acos((rPS * rPS + rPE * rPE - rES * rES) / (2 * rPS * rPE))
End Function

'(*****************************************************************************)
'(* Name:    CalcElongation                                                   *)
'(* Type:    Function                                                         *)
'(* Purpose: Calculate elongation from interplanetary distances.              *)
'(* Arguments:                                                                *)
'(*   rPS, rES, rPE : distance Planet-Sun, Earth-Sun and Planet-Earth,        *)
'(*                   respectively                                            *)
'(* Return value:                                                             *)
'(*   the elongation                                                          *)
'(*****************************************************************************)

Function CalcElongation(rPS As Double, rES As Double, rPE As Double) As Double
    CalcElongation = acos((rES * rES + rPE * rPE - rPS * rPS) / (2 * rES * rPE))
End Function
                    
'(*****************************************************************************)
'(* Name:    PlanetMagnitude                                                  *)
'(* Type:    Function                                                         *)
'(* Purpose: Calculate the approximate visual magnitude of a planet.          *)
'(* Arguments:                                                                *)
'(*   Planet : planet number (1 = Mercury, 9 = Pluto)                         *)
'(*   rPS, rPE : distance Planet-Sun and Planet-Earth, respectively           *)
'(*   i : phase angle                                                         *)
'(*   dU, B : the DeltaU and B fields from TSATURNRINGDATA.  Only relevant    *)
'(*           for Saturn                                                      *)
'(* Return value:                                                             *)
'(*   the approximate magnitude, as in the Astronomical Almanac since 1984    *)
'(*****************************************************************************)

Function PlanetMagnitude(Planet As Long, rPS As Double, rPE As Double, i As Double, dU As Double, B As Double) As Double
Dim Mag As Double
Dim AbsPlanetMag As Variant
  
AbsPlanetMag = Array(0, -0.42, -4.4, 0, -1.52, -9.4, -8.88, -7.19, -6.87, -1#)
Mag = AbsPlanetMag(Planet) + 5 * log10(rPS * rPE)
i = i * RToD
Select Case Planet
Case 1:    Mag = Mag + i * (0.038 + i * (-0.000273 + i * 0.000002))
Case 2:    Mag = Mag + i * (0.0009 + i * (0.000239 - i * 0.00000065))
Case 4:    Mag = Mag + i * 0.016
Case 5:    Mag = Mag + i * 0.0005
Case 6:    Mag = Mag + 0.044 * dU * RToD + Sin(B) * (-2.6 + 1.25 * Sin(B))
End Select
PlanetMagnitude = Mag
End Function

'(*****************************************************************************)
'(* Name:    CometMagnitude                                                   *)
'(* Type:    Function                                                         *)
'(* Purpose: Calculate the approximate visual magnitude of a comet.           *)
'(* Arguments:                                                                *)
'(*   rPS, rPE : distance Comet-Sun and Comet-Earth, respectively             *)
'(*   g, k : the comet's magnitude parameters                                 *)
'(* Return value:                                                             *)
'(*   the approximate magnitude                                               *)
'(*****************************************************************************)

Function CometMagnitude(rCS As Double, rCE As Double, G As Double, k As Double) As Double
CometMagnitude = G + 5 * Log(rCS) + k * Log(rCE)
End Function

'(*****************************************************************************)
'(* Name:    AsteroidMagnitude                                                *)
'(* Type:    Function                                                         *)
'(* Purpose: Calculate the approximate visual magnitude of an asteroid.       *)
'(* Arguments:                                                                *)
'(*   rPS, rPE : distance Asteroid-Sun and Asteroid-Earth, respectively       *)
'(*   PhaseAngle : the asteroid's phase angle                                 *)
'(*   H, G : the asteroid's magnitude parameters                              *)
'(* Return value:                                                             *)
'(*   the approximate magnitude                                               *)
'(*****************************************************************************)

Function AsteroidMagnitude(rAS As Double, rAE As Double, PhaseAngle As Double, H As Double, G As Double) As Double
Dim Phi1 As Double, Phi2 As Double
Phi1 = Exp(-3.33 * Exp(0.63 * Log(tan(PhaseAngle / 2))))
Phi2 = Exp(-1.87 * Exp(1.22 * Log(tan(PhaseAngle / 2))))
AsteroidMagnitude = H + 5 * log10(rAS * rAE) - 2.5 * log10((1 - G) * Phi1 + G * Phi2)
End Function

'(*****************************************************************************)
'(* Name:    PlanetSemiDiameter                                               *)
'(* Type:    Function                                                         *)
'(* Purpose: Calculate the apparent semi-diameter of a planet.                *)
'(* Arguments:                                                                *)
'(*   Planet : planet number (Sun = 0, Pluto = 9)                             *)
'(*   DistEarth : distance of planet/Sun from Earth                           *)
'(*   PolarDiam : to hold the polar semi-diameter if available;               *)
'(* Return value:                                                             *)
'(*   the apparent semi-diameter                                              *)
'(*****************************************************************************)

Function PlanetSemiDiameter(Planet As Long, DistEarth As Double, ByRef PolarDiam As Double) As Double
Dim PlanetSemiDiamTab As Variant

PlanetSemiDiamTab = Array(959.63, 3.36, 8.34, 0, 4.68, 98.44, 82.73, 35.02, 33.5, 2.07)
Select Case Planet
  Case 5: PolarDiam = 92.06 / DistEarth
  Case 6: PolarDiam = 73.82 / DistEarth
End Select
PlanetSemiDiameter = PlanetSemiDiamTab(Planet) / DistEarth
End Function

'(*****************************************************************************)
'(* Name:    CorrectSaturnSemiDiameter                                        *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Correct Saturn's polar semi-diameter for the planet's 'tilt'     *)
'(* Arguments:                                                                *)
'(*   B : the Saturnicentric altitude of the Earth                            *)
'(*   PolarSemiDiameter : the (on return apparent) polar semi-diameter        *)
'(*****************************************************************************)

Sub CorrectSaturnSemiDiameter(B As Double, ByRef PolarSemiDiameter As Double)
PolarSemiDiameter = PolarSemiDiameter * Sqr(1 - 0.2038 * Cos(B) * Cos(B))
End Sub


'(*****************************************************************************)
'(* Name:    SunPhysEphemeris                                                 *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Calculate a physical ephemeris of the Sun (L0, B0 and P)         *)
'(* Arguments:                                                                *)
'(*   T : number of centuries since J2000                                     *)
'(*   l : Sun's mean longitude                                                *)
'(*   Obl : obliquity of the ecliptic                                         *)
'(*   NutLon : nutation in longitude                                          *)
'(*   SunPhysData : TSUNPHYSDATA record to hold the return values             *)
'(*****************************************************************************)

Sub SunPhysEphemeris(T As Double, l As Double, Obl As Double, NutLon As Double, ByRef SunPhysData As TSUNPHYSDATA)

Dim JD As Double, Theta As Double, i As Double, k As Double
Dim C As Double, s As Double
Dim x As Double, Y As Double, Eta As Double
JD = (T * 36525#) + 2451545#
Theta = (JD - 2398220) * Pi2 / 25.38
i = 7.25 * DToR
k = (73.6667 + 1.3958333 * (JD - 2396758#) / 36525#) * DToR
C = Cos(l - k)
s = Sin(l - k)
x = Atn(-Cos(l + NutLon) * tan(Obl))
Y = Atn(-C * tan(i))
SunPhysData.P = x + Y
SunPhysData.b0 = asin(s * Sin(i))
Eta = atan2(-s * Cos(i), -C)
SunPhysData.l0 = modpi2(Eta - Theta)
End Sub

'(*****************************************************************************)
'(* Name:    JDOfCarringtonRotation                                           *)
'(* Type:    Function                                                         *)
'(* Purpose: Calculate the beginning of a Carrington rotation.                *)
'(* Arguments:                                                                *)
'(*   C : Carrington rotation number                                          *)
'(* Return value:                                                             *)
'(*   the Julian Day corresponding to the beginning of the C'th Carrington    *)
'(*   rotation                                                                *)
'(*****************************************************************************)

Function JDOfCarringtonRotation(C As Long) As Double
Dim JD As Double, M As Double
JD = 2398140.227 + 27.2752316 * C
M = (281.96 + 26.882476 * C) * DToR
JDOfCarringtonRotation = JD + 0.1454 * Sin(M) - 0.0085 * Sin(2 * M) - 0.0141 * Cos(2 * M)
End Function

'(*****************************************************************************)
'(* Name:    CarringtonRotation                                               *)
'(* Type:    Function                                                         *)
'(* Purpose: Calculate the current Carrington rotation number at a given      *)
'(*          instant                                                          *)
'(* Arguments:                                                                *)
'(*   JD : Julian Day                                                         *)
'(* Return value:                                                             *)
'(*   the Carrington number of the current synodic rotation of the Sun.       *)
'(*****************************************************************************)

Function CarringtonRotation(JD As Double) As Integer

Dim C As Double, M As Double
C = floor((JD - 2398140.227) / 27.2752316)
M = (281.96 + 26.882476 * C) * DToR
JD = JD - 0.1454 * Sin(M) + 0.0085 * Sin(2 * M) + 0.0141 * Cos(2 * M)
C = floor((JD - 2398140.227) / 27.2752316)
CarringtonRotation = Round(C)
End Function

'(*****************************************************************************)
'(* Name:    MarsPhysEphemeris                                                *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Calculate a physical ephemeris of Mars.                          *)
'(* Arguments:                                                                *)
'(*   T : Julian centuries since J2000                                        *)
'(*   SHelio, SGeo, SEarth : TSVECTOR records holding the ecliptical          *)
'(*                          coordinates of Mars (heliocentric and            *)
'(*                          geocentric) and the Earth                        *)
'(*   Obl : mean obliquity of the ecliptic                                    *)
'(*   NutLon, NutObl : nutation in longitude and obliquity                    *)
'(*   MarsPhysData : TMARSPHYSDATA record to hold the results                 *)
'(*****************************************************************************)

Sub MarsPhysEphemeris(T As Double, SHelio As TSVECTOR, SEarth As TSVECTOR, SGeo As TSVECTOR, _
                            Obl As Double, NutLon As Double, NutObl As Double, _
                            ByRef MarsPhysData As TMARSPHYSDATA)

Dim l0 As Double, b0 As Double
Dim n As Double, JDE As Double, W As Double
Dim RA0 As Double, Decl0 As Double
Dim r As TVECTOR
Dim u As Double, v As Double, RA As Double, Decl As Double, P As Double, k As Double
'{ 1. }
l0 = (352.9065 + 1.1733 * T) * DToR
b0 = (63.2818 - 0.00394 * T) * DToR

'{ 2. }
'{ 3. }
'{ 4. }
'{ 5. }
'{ Already done }

'{ 6. }
MarsPhysData.DE = -asin(Sin(b0) * Sin(SGeo.B) + Cos(b0) * Cos(SGeo.B) * Cos(l0 - SGeo.l))

'{ 7. }
n = (49.5581 + 0.7721 * T) * DToR
SHelio.l = SHelio.l - 0.00697 * DToR / SHelio.r
SHelio.B = SHelio.B - 0.000225 * DToR * Cos(SHelio.l - n) / SHelio.r

'{ 8. }
MarsPhysData.DS = -asin(Sin(b0) * Sin(SHelio.B) + Cos(b0) * Cos(SHelio.B) * Cos(l0 - SHelio.l))

'{ 9. }
JDE = (T - SGeo.r * LightTimeConst) * 36525 + 2451545#
W = (11.504 + 350.89200025 * (JDE - 2433282.5)) * DToR

'{ 10. }
Call EclToEqu(l0, b0, Obl, RA0, Decl0)

'{ 11. }
Call SphToRect(SGeo, r)
u = r.Y * Cos(Obl) - r.Z * Sin(Obl)
v = r.Y * Sin(Obl) + r.Z * Cos(Obl)
RA = atan2(u, r.x)
Decl = Atn(v / Sqr(r.x * r.x + u * u))
v = Sin(Decl0) * Cos(Decl) * Cos(RA0 - RA) - Sin(Decl) * Cos(Decl0)
u = Cos(Decl) * Sin(RA0 - RA)
P = atan2(v, u)

'{ 12. }
MarsPhysData.Om = modpi2(W - P)

'{ 13. }
'{ Already done }

'{ 14. }
SGeo.l = SGeo.l + 0.005693 * DToR * Cos(SEarth.l - SGeo.l) / Cos(SGeo.B)
SGeo.B = SGeo.B + 0.005693 * DToR * Sin(SEarth.l - SGeo.l) * Sin(SGeo.B)

'{ 15. }
l0 = l0 + NutLon
SGeo.l = SGeo.l + NutLon
Obl = Obl + NutObl

'{ 16. }
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
Call EclToEqu(l0, b0, Obl, RA0, Decl0)

'{ 17. }
v = Cos(Decl0) * Sin(RA0 - RA)
u = Sin(Decl0) * Cos(Decl) - Cos(Decl0) * Sin(Decl) * Cos(RA0 - RA)
MarsPhysData.P = modpi2(atan2(v, u))

'{ 18. }
Call EclToEqu(SEarth.l + PI, SEarth.B, Obl, RA0, Decl0)
v = Cos(Decl0) * Sin(RA0 - RA)
u = Sin(Decl0) * Cos(Decl) - Cos(Decl0) * Sin(Decl) * Cos(RA0 - RA)
MarsPhysData.Q = atan2(v, u) + PI

'{ 19. }
MarsPhysData.D = 9.36 / SGeo.r
k = CalcPhase(SHelio.r, SEarth.r, SGeo.r)
MarsPhysData.qq = (1 - k) * MarsPhysData.D
End Sub


'(*****************************************************************************)
'(* Name:    JupiterPhysEphemeris                                             *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Calculate a physical ephemeris of Jupiter.                       *)
'(* Arguments:                                                                *)
'(*   T : Julian centuries since J2000                                        *)
'(*   SHelio, SGeo, SEarth : TSVECTOR records holding the ecliptical          *)
'(*                          coordinates of Jupiter (heliocentric and         *)
'(*                          geocentric) and the Earth                        *)
'(*   Obl : mean obliquity of the ecliptic                                    *)
'(*   NutLon, NutObl : nutation in longitude and obliquity                    *)
'(*   JupiterPhysData : TJUPITERPHYSDATA record to hold the results           *)
'(*****************************************************************************)

Sub JupiterPhysEphemeris(T As Double, SHelio As TSVECTOR, SEarth As TSVECTOR, SGeo As TSVECTOR, _
                               Obl As Double, NutLon As Double, NutObl As Double, _
                               ByRef JupiterPhysData As TJUPITERPHYSDATA)

                               
Dim JDE As Double, D As Double, T1 As Double
Dim RA0 As Double, Decl0 As Double
Dim W1 As Double, W2 As Double
Dim rAS As Double, DeclS As Double
Dim u As Double, v As Double, RA As Double, Decl As Double, P As Double
Dim r As TVECTOR
'{ 1. }
JDE = T * 36525 + 2451545#
D = JDE - 2433282.5
T1 = D / 36525
RA0 = (268# + 0.1061 * T1) * DToR
Decl0 = (64.5 - 0.0164 * T1) * DToR

'{ 2. }
W1 = modpi2((17.71 + 877.90003539 * D) * DToR)
W2 = modpi2((16.838 + 870.27003539 * D) * DToR)

'{ 3. }
'{ 4. }
'{ 5. }
'{ 6. }
'{ 7. }
'{ 8. }
'{ Already done }

'{ 9. }
Call EclToEqu(SHelio.l, SHelio.B, Obl, rAS, DeclS)

'{ 10. }
JupiterPhysData.DS = -asin(Sin(Decl0) * Sin(DeclS) + Cos(Decl0) * Cos(DeclS) * Cos(RA0 - rAS))

'{ 11. }
Call SphToRect(SGeo, r)
u = r.Y * Cos(Obl) - r.Z * Sin(Obl)
v = r.Y * Sin(Obl) + r.Z * Cos(Obl)
RA = atan2(u, r.x)
Decl = Atn(v / Sqr(r.x * r.x + u * u))
v = Sin(Decl0) * Cos(Decl) * Cos(RA0 - RA) - Sin(Decl) * Cos(Decl0)
u = Cos(Decl) * Sin(RA0 - RA)
P = atan2(v, u)

'{ 12. }
JupiterPhysData.DE = -asin(Sin(Decl0) * Sin(Decl) + Cos(Decl0) * Cos(Decl) * Cos(RA0 - RA))

'{ 13. }
JupiterPhysData.Om1 = modpi2(W1 - P - 5.07033 * DToR * SGeo.r)
JupiterPhysData.Om2 = modpi2(W2 - P - 5.02626 * DToR * SGeo.r)

'{ 14. }
JupiterPhysData.C = 1 - CalcPhase(SHelio.r, SEarth.r, SGeo.r)
If (Sin(SHelio.l - SEarth.l) < 0) Then
  JupiterPhysData.C = -JupiterPhysData.C
End If
'{ 15. }
Obl = Obl + NutObl

'{ 16. }
RA = RA + 0.005693 * DToR * (Cos(RA) * Cos(SEarth.l) * Cos(Obl) + Sin(RA) * Sin(SEarth.l)) / Cos(Decl)
Decl = Decl + 0.005693 * DToR * (Cos(SEarth.l) * Cos(Obl) * (tan(Obl) * Cos(Decl) - Sin(RA) * Sin(Decl)) + Cos(RA) * Sin(Decl) * Sin(SEarth.l))

'{ 17. }
Call Nutation(NutLon, NutObl, Obl, RA0, Decl0)
Call Nutation(NutLon, NutObl, Obl, RA, Decl)

'{ 18. }
v = Cos(Decl0) * Sin(RA0 - RA)
u = Sin(Decl0) * Cos(Decl) - Cos(Decl0) * Sin(Decl) * Cos(RA0 - RA)
JupiterPhysData.P = modpi2(atan2(v, u))

End Sub

