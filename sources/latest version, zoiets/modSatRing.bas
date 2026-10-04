Attribute VB_Name = "modSatRing"
'(*****************************************************************************)
'(* Module: SATRING.PAS                                                       *)
'(* Version 2.0                                                               *)
'(* Last modified: October 1, 1992                                            *)
'(*****************************************************************************)
Sub ElementsToU1B1P1(I As Double, l_om As Double, B As Double, ByRef u1 As Double, ByRef b1 As Double, ByRef P1 As Double)

Dim cB1sP1 As Double, cB1cP1 As Double, sB1 As Double, cB1sU1 As Double, cB1cU1  As Double
     cB1sP1 = -Sin(I) * Cos(l_om)
     cB1cP1 = Cos(I) * Cos(B) + Sin(I) * Sin(B) * Sin(l_om)
     sB1 = -Cos(I) * Sin(B) + Sin(I) * Cos(B) * Sin(l_om)
     cB1sU1 = Sin(I) * Sin(B) + Cos(I) * Cos(B) * Sin(l_om)
     cB1cU1 = Cos(B) * Cos(l_om)

     P1 = atan2(cB1sP1, cB1cP1)
     u1 = atan2(cB1sU1, cB1cU1)
     b1 = atan2(sB1, cB1sU1 / Sin(u1))
End Sub

Sub ElementsToUBP(j As Double, d As Double, a_N As Double, ByRef u As Double, ByRef B As Double, ByRef P As Double)
Dim cBsU As Double, cBcU As Double, sB As Double, cBsP As Double, cBcP  As Double
     cBsU = Cos(j) * Cos(d) * Sin(a_N) + Sin(j) * Sin(d)
     cBcU = Cos(d) * Cos(a_N)
     sB = Sin(j) * Cos(d) * Sin(a_N) - Cos(j) * Sin(d)
     cBsP = -Sin(j) * Cos(a_N)
     cBcP = Sin(j) * Sin(d) * Sin(a_N) + Cos(j) * Cos(d)

     u = atan2(cBsU, cBcU)
     P = atan2(cBsP, cBcP)
     B = atan2(sB, cBsP / Sin(P))
End Sub

Sub ElementsToJNW(I As Double, Om As Double, eps As Double, ByRef j As Double, ByRef n As Double, ByRef W As Double)
Dim sJsN As Double, sJcN As Double, cJ As Double, sJsW As Double, sJcW  As Double
     sJsN = Sin(I) * Sin(Om)
     sJcN = Cos(I) * Sin(eps) + Sin(I) * Cos(eps) * Cos(Om)
     cJ = Cos(I) * Cos(eps) - Sin(I) * Sin(eps) * Cos(Om)
     sJsW = Sin(eps) * Sin(Om)
     sJcW = Sin(I) * Cos(eps) + Cos(I) * Sin(eps) * Cos(Om)

     n = atan2(sJsN, sJcN)
     W = atan2(sJsW, sJcW)
     j = atan2(sJsW / Sin(W), cJ)
End Sub

'(*****************************************************************************)
'(* Name:    SaturnRing                                                       *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Calculate the position of Saturn's ring system.                  *)
'(* Arguments:                                                                *)
'(*   T : Julian centuries since J2000                                        *)
'(*   SHelio, SGeo : TSVECTOR records holding the ecliptical coordinates of   *)
'(*                  Saturn (heliocentric and geocentric)                     *)
'(*   Obl : mean obliquity of the ecliptic                                    *)
'(*   NutLon, NutObl : nutation in longitude and obliquity                    *)
'(*   SaturnRingData : TSATURNRINGDATA record to hold the results             *)
'(*****************************************************************************)

Sub SaturnRing(T As Double, sHelio As TSVECTOR, sGeo As TSVECTOR, _
                     obl As Double, NutLon As Double, NutObl As Double, _
                     ByRef SaturnRingData As TSATURNRINGDATA)

Dim I As Double, Om As Double
Dim n As Double, u As Double, v As Double, u1 As Double, u2   As Double
Dim L0 As Double, b0   As Double
Dim RA As Double, Decl As Double, RA0 As Double, Decl0   As Double
'{ 1. }
I = (28.075216 - T * (0.012998 - T * 0.000004)) * DToR
Om = (169.50847 + T * (1.394681 + T * 0.000412)) * DToR

'{ 2. }
' { 3. }
'{ 4. }
'{ 5. }
'{ Already done }

'{ 6. }
SaturnRingData.B = asin(Sin(I) * Cos(sGeo.B) * Sin(sGeo.L - Om) - Cos(I) * Sin(sGeo.B))
SaturnRingData.aAxis = 375.35 / sGeo.r
SaturnRingData.bAxis = SaturnRingData.aAxis * Abs(Sin(SaturnRingData.B))
SaturnRingData.ioaAxis = SaturnRingData.aAxis * 0.8801
SaturnRingData.iobAxis = SaturnRingData.bAxis * 0.8801
SaturnRingData.oiaAxis = SaturnRingData.aAxis * 0.8599
SaturnRingData.oibAxis = SaturnRingData.bAxis * 0.8599
SaturnRingData.iiaAxis = SaturnRingData.aAxis * 0.665
SaturnRingData.iibAxis = SaturnRingData.bAxis * 0.665
SaturnRingData.idaAxis = SaturnRingData.aAxis * 0.5486
SaturnRingData.idbAxis = SaturnRingData.bAxis * 0.5486

'{ 7. }
n = (113.6655 + 0.8771 * T) * DToR
sHelio.L = sHelio.L - 0.01759 * DToR / sHelio.r
sHelio.B = sHelio.B - 0.000764 * DToR * Cos(sHelio.L - n) / sHelio.r

'{ 8. }
SaturnRingData.Bd = asin(Sin(I) * Cos(sHelio.B) * Sin(sHelio.L - Om) - Cos(I) * Sin(sHelio.B))

'{ 9. }
v = Sin(I) * Sin(sHelio.B) + Cos(I) * Cos(sHelio.B) * Sin(sHelio.L - Om)
u = Cos(sHelio.B) * Cos(sHelio.L - Om)
u1 = atan2(v, u)
v = Sin(I) * Sin(sGeo.B) + Cos(I) * Cos(sGeo.B) * Sin(sGeo.L - Om)
u = Cos(sGeo.B) * Cos(sGeo.L - Om)
u2 = atan2(v, u)
SaturnRingData.DeltaU = Abs(u1 - u2)
                  '{DeltaU is altijd kleiner dan 7 gr.}
If SaturnRingData.DeltaU > Pi Then SaturnRingData.DeltaU = Pi2 - SaturnRingData.DeltaU

'{ 10. }
'{ Already done }

'{ 11. }
L0 = Om - Pi / 2
b0 = Pi / 2 - I

'{ 12. }
sGeo.L = sGeo.L + 0.005693 * DToR * Cos(L0 - sGeo.L) / Cos(sGeo.B)
sGeo.B = sGeo.B + 0.005693 * DToR * Sin(L0 - sGeo.L) * Sin(sGeo.B)

'{ 13. }
sGeo.L = sGeo.L + NutLon
L0 = L0 + NutLon
obl = obl + NutObl

'{ 14. }
Call EclToEqu(L0, b0, obl, RA0, Decl0)
Call EclToEqu(sGeo.L, sGeo.B, obl, RA, Decl)

'{ 15. }
v = Cos(Decl0) * Sin(RA0 - RA)
u = Sin(Decl0) * Cos(Decl) - Cos(Decl0) * Sin(Decl) * Cos(RA0 - RA)
SaturnRingData.P = modpi2(atan2(v, u))
End Sub

Sub AltSaturnRing(T As Double, sHelio As TSVECTOR, sGeo As TSVECTOR, _
                        obl As Double, NutLon As Double, NutObl As Double, _
                        ByRef AltSaturnRingData As TALTSATURNRINGDATA)

Dim I As Double, Om  As Double
Dim n As Double, u As Double, v As Double, u1 As Double, u2   As Double
Dim j As Double, W   As Double
Dim L0 As Double, b0   As Double
Dim RA As Double, Decl As Double, RA0 As Double, Decl0   As Double

'{ 1. }
I = (28.075216 - T * (0.012998 - T * 0.000004)) * DToR
Om = (169.50847 + T * (1.394681 + T * 0.000412)) * DToR


'{ 7. }
n = (113.6655 + 0.8771 * T) * DToR
sHelio.L = sHelio.L - 0.01759 * DToR / sHelio.r
sHelio.B = sHelio.B - 0.000764 * DToR * Cos(sHelio.L - n) / sHelio.r

'{ 11. }
L0 = Om - Pi / 2
b0 = Pi / 2 - I

'{ 12. }
sGeo.L = sGeo.L + 0.005693 * DToR * Cos(L0 - sGeo.L) / Cos(sGeo.B)
sGeo.B = sGeo.B + 0.005693 * DToR * Sin(L0 - sGeo.L) * Sin(sGeo.B)

'{ 13. }
sGeo.L = sGeo.L + NutLon
L0 = L0 + NutLon
obl = obl + NutObl

'{ 14. }
Call EclToEqu(L0, b0, obl, RA0, Decl0)
Call EclToEqu(sGeo.L, sGeo.B, obl, RA, Decl)

With AltSaturnRingData
     Call ElementsToJNW(I, Om, obl, .j, .n, .W)
     Call ElementsToUBP(.j, Decl, RA - .n, .u, .B, .P)
     Call ElementsToU1B1P1(I, sHelio.L - Om, sHelio.B, .u1, .b1, .P1)
End With
End Sub

