Attribute VB_Name = "modSatRing"
'(*****************************************************************************)
'(* Module: SATRING.PAS                                                       *)
'(* Version 2.0                                                               *)
'(* Last modified: October 1, 1992                                            *)
'(*****************************************************************************)
Sub ElementsToU1B1P1(i As Double, l_om As Double, B As Double, ByRef U1 As Double, ByRef b1 As Double, ByRef P1 As Double)

Dim cB1sP1 As Double, cB1cP1 As Double, sB1 As Double, cB1sU1 As Double, cB1cU1  As Double
     cB1sP1 = -Sin(i) * Cos(l_om)
     cB1cP1 = Cos(i) * Cos(B) + Sin(i) * Sin(B) * Sin(l_om)
     sB1 = -Cos(i) * Sin(B) + Sin(i) * Cos(B) * Sin(l_om)
     cB1sU1 = Sin(i) * Sin(B) + Cos(i) * Cos(B) * Sin(l_om)
     cB1cU1 = Cos(B) * Cos(l_om)

     P1 = atan2(cB1sP1, cB1cP1)
     U1 = atan2(cB1sU1, cB1cU1)
     b1 = atan2(sB1, cB1sU1 / Sin(U1))
End Sub

Sub ElementsToUBP(j As Double, D As Double, a_N As Double, ByRef u As Double, ByRef B As Double, ByRef P As Double)
Dim cBsU As Double, cBcU As Double, sB As Double, cBsP As Double, cBcP  As Double
     cBsU = Cos(j) * Cos(D) * Sin(a_N) + Sin(j) * Sin(D)
     cBcU = Cos(D) * Cos(a_N)
     sB = Sin(j) * Cos(D) * Sin(a_N) - Cos(j) * Sin(D)
     cBsP = -Sin(j) * Cos(a_N)
     cBcP = Sin(j) * Sin(D) * Sin(a_N) + Cos(j) * Cos(D)

     u = atan2(cBsU, cBcU)
     P = atan2(cBsP, cBcP)
     B = atan2(sB, cBsP / Sin(P))
End Sub

Sub ElementsToJNW(i As Double, Om As Double, eps As Double, ByRef j As Double, ByRef n As Double, ByRef W As Double)
Dim sJsN As Double, sJcN As Double, cJ As Double, sJsW As Double, sJcW  As Double
     sJsN = Sin(i) * Sin(Om)
     sJcN = Cos(i) * Sin(eps) + Sin(i) * Cos(eps) * Cos(Om)
     cJ = Cos(i) * Cos(eps) - Sin(i) * Sin(eps) * Cos(Om)
     sJsW = Sin(eps) * Sin(Om)
     sJcW = Sin(i) * Cos(eps) + Cos(i) * Sin(eps) * Cos(Om)

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

Sub SaturnRing(T As Double, SHelio As TSVECTOR, SGeo As TSVECTOR, _
                     Obl As Double, NutLon As Double, NutObl As Double, _
                     ByRef SaturnRingData As TSATURNRINGDATA)

Dim i As Double, Om As Double
Dim n As Double, u As Double, v As Double, U1 As Double, U2   As Double
Dim l0 As Double, b0   As Double
Dim RA As Double, Decl As Double, RA0 As Double, Decl0   As Double
'{ 1. }
i = (28.075216 - T * (0.012998 - T * 0.000004)) * DToR
Om = (169.50847 + T * (1.394681 + T * 0.000412)) * DToR

'{ 2. }
' { 3. }
'{ 4. }
'{ 5. }
'{ Already done }

'{ 6. }
SaturnRingData.B = asin(Sin(i) * Cos(SGeo.B) * Sin(SGeo.l - Om) - Cos(i) * Sin(SGeo.B))
SaturnRingData.aAxis = 375.35 / SGeo.r
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
SHelio.l = SHelio.l - 0.01759 * DToR / SHelio.r
SHelio.B = SHelio.B - 0.000764 * DToR * Cos(SHelio.l - n) / SHelio.r

'{ 8. }
SaturnRingData.Bd = asin(Sin(i) * Cos(SHelio.B) * Sin(SHelio.l - Om) - Cos(i) * Sin(SHelio.B))

'{ 9. }
v = Sin(i) * Sin(SHelio.B) + Cos(i) * Cos(SHelio.B) * Sin(SHelio.l - Om)
u = Cos(SHelio.B) * Cos(SHelio.l - Om)
U1 = atan2(v, u)
v = Sin(i) * Sin(SGeo.B) + Cos(i) * Cos(SGeo.B) * Sin(SGeo.l - Om)
u = Cos(SGeo.B) * Cos(SGeo.l - Om)
U2 = atan2(v, u)
SaturnRingData.DeltaU = Abs(U1 - U2)
                  '{DeltaU is altijd kleiner dan 7 gr.}
If SaturnRingData.DeltaU > PI Then SaturnRingData.DeltaU = Pi2 - SaturnRingData.DeltaU

'{ 10. }
'{ Already done }

'{ 11. }
l0 = Om - PI / 2
b0 = PI / 2 - i

'{ 12. }
SGeo.l = SGeo.l + 0.005693 * DToR * Cos(l0 - SGeo.l) / Cos(SGeo.B)
SGeo.B = SGeo.B + 0.005693 * DToR * Sin(l0 - SGeo.l) * Sin(SGeo.B)

'{ 13. }
SGeo.l = SGeo.l + NutLon
l0 = l0 + NutLon
Obl = Obl + NutObl

'{ 14. }
Call EclToEqu(l0, b0, Obl, RA0, Decl0)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)

'{ 15. }
v = Cos(Decl0) * Sin(RA0 - RA)
u = Sin(Decl0) * Cos(Decl) - Cos(Decl0) * Sin(Decl) * Cos(RA0 - RA)
SaturnRingData.P = modpi2(atan2(v, u))
End Sub

Sub AltSaturnRing(T As Double, SHelio As TSVECTOR, SGeo As TSVECTOR, _
                        Obl As Double, NutLon As Double, NutObl As Double, _
                        ByRef AltSaturnRingData As TALTSATURNRINGDATA)

Dim i As Double, Om  As Double
Dim n As Double, u As Double, v As Double, U1 As Double, U2   As Double
Dim j As Double, W   As Double
Dim l0 As Double, b0   As Double
Dim RA As Double, Decl As Double, RA0 As Double, Decl0   As Double

'{ 1. }
i = (28.075216 - T * (0.012998 - T * 0.000004)) * DToR
Om = (169.50847 + T * (1.394681 + T * 0.000412)) * DToR


'{ 7. }
n = (113.6655 + 0.8771 * T) * DToR
SHelio.l = SHelio.l - 0.01759 * DToR / SHelio.r
SHelio.B = SHelio.B - 0.000764 * DToR * Cos(SHelio.l - n) / SHelio.r

'{ 11. }
l0 = Om - PI / 2
b0 = PI / 2 - i

'{ 12. }
SGeo.l = SGeo.l + 0.005693 * DToR * Cos(l0 - SGeo.l) / Cos(SGeo.B)
SGeo.B = SGeo.B + 0.005693 * DToR * Sin(l0 - SGeo.l) * Sin(SGeo.B)

'{ 13. }
SGeo.l = SGeo.l + NutLon
l0 = l0 + NutLon
Obl = Obl + NutObl

'{ 14. }
Call EclToEqu(l0, b0, Obl, RA0, Decl0)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)

With AltSaturnRingData
     Call ElementsToJNW(i, Om, Obl, .j, .n, .W)
     Call ElementsToUBP(.j, Decl, RA - .n, .u, .B, .P)
     Call ElementsToU1B1P1(i, SHelio.l - Om, SHelio.B, .U1, .b1, .P1)
End With
End Sub

