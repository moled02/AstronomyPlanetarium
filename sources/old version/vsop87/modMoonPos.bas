Attribute VB_Name = "modMoonPos"
'(*****************************************************************************)
'(*                                                                           *)
'(*                  Copyright (c) 1991-1992 by Jeffrey Sax                   *)
'(*                            All rights reserved                            *)
'(*                        Published and Distributed by                       *)
'(*                           Willmann-Bell, Inc.                             *)
'(*                             P.O. Box 35025                                *)
'(*                        Richmond, Virginia 23235                           *)
'(*                Voice (804) 320-7016 FAX (804) 272-5920                    *)
'(*                                                                           *)
'(*                                                                           *)
'(*                NOTICE TO COMMERCIAL SOFTWARE DEVELOPERS                   *)
'(*                                                                           *)
'(*        Prior to distributing software incorporating this code             *)
'(*        you MUST write Willmann-Bell, Inc. at the above address            *)
'(*        for validation of your book's (Astronomical Algorithms             *)
'(*        by Jean Meeus) and software Serial Numbers.  No additional         *)
'(*        fees will be required BUT you MUST have the following              *)
'(*        notice at the start of your program(s):                            *)
'(*                                                                           *)
'(*                    This program contains code                             *)
'(*              Copyright (c) 1991-1992 by Jeffrey Sax                       *)
'(*              and Distributed by Willmann-Bell, Inc.                       *)
'(*                         Serial #######                                    *)
'(*                                                                           *)
'(*****************************************************************************)
'(* Module: MOONPOS.PAS                                                       *)
'(* Version 2.0                                                               *)
'(* Last modified: October 1, 1992                                            *)
'(*****************************************************************************)

'(*****************************************************************************)
'(* Name:    EvalMoonAngles                                                   *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Calculate the angles Om, D, M, Md and F, as well as E.           *)
'(* Arguments:                                                                *)
'(*   T : number of Julian centuries since J2000                              *)
'(*   Om, D, M, Md, F, E : the results                                        *)
'(*                                                                           *)
'(* Note: This procedure is also used by MoonPhysEphemeris and Eclipse        *)
'(*****************************************************************************)

Private Sub CalcMoonAngles(ByVal T As Double, ByRef Om As Double, ByRef D As Double, ByRef M As Double, ByRef Md As Double, ByRef F As Double, ByRef e As Double)
Dim pOm As Variant, PD As Variant, PM As Variant, PMd As Variant, PF As Variant
    
    pOm = Array(125.044555, -1934.1361849, 0.0020762, 0.000002139449, 0.0000000164973)
    PD = Array(297.8502042, 445267.1115168, -0.00163, 0.000001831945, 0.00000884447)
    PM = Array(357.5291092, 35999.0502909, -0.0001536, 0.00000004083299, 0)
    PMd = Array(134.9634114, 477198.8676313, 0.008997, 0.00001434741, 0.0000000679717)
    PF = Array(93.2720993, 483202.0175273, -0.0034029, 0.0000002836075, 0.00000000115833)
Om = modpi2(Eval4Poly(C4P(pOm), T) * DToR)
D = modpi2(Eval4Poly(C4P(PD), T) * DToR)
M = modpi2(Eval4Poly(C4P(PM), T) * DToR)
Md = modpi2(Eval4Poly(C4P(PMd), T) * DToR)
F = modpi2(Eval4Poly(C4P(PF), T) * DToR)
e = 1 - T * (0.002516 + T * 0.0000074)
End Sub

'(*****************************************************************************)
'(* Name:    MoonPos                                                          *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Calculate the ecliptic coordinates and radius vector of the Moon *)
'(* Arguments:                                                                *)
'(*   T : number of Julian centuries since J2000                              *)
'(*   S : TSVECTOR record to hold the longitude, latitude and distance to     *)
'(*       Earth (in km)                                                       *)
'(*****************************************************************************)

Sub MoonPos(ByVal T As Double, ByRef s As TSVECTOR)
Dim i As Long, j As Long, k As Long, k1     As Long
Dim Om     As Double, ld As Double, D As Double, M As Double, Md As Double, F As Double
Dim dl As Double, db As Double, dr As Double
Dim a1     As Double, a2 As Double, A3 As Double
Dim Flag As Boolean
Dim SinVal     As Double, CosVal As Double, ArgSin As Double, ArgCos As Double, tmp As Double
Dim LonArg     As Double, RadArg As Double
Dim SinTab(4) As TSINCOSTAB, CosTab(4) As TSINCOSTAB
Dim e(2) As Double

Dim PLd As Variant
    PLd = Array(218.3164591, 481267.88134236, -0.0013268, 0.000001855835, 0.0000000153388)
    '{Table 45.A}
    MOONTABTERMS = 60
Dim MoonTermsLR As Variant
'    MoonTermsLR : array[1..MOONTABTERMS,1..6] of longint =

MoonTermsLR = Array(Array(0, 0, 1, 0, 6288774, -20905355), Array(2, 0, -1, 0, 1274027, -3699111), _
    Array(2, 0, 0, 0, 658314, -2955968), Array(0, 0, 2, 0, 213618, -569925), Array(0, 1, 0, 0, -185116, 48888), Array(0, 0, 0, 2, -114332, -3149), _
    Array(2, 0, -2, 0, 58793, 246158), Array(2, -1, -1, 0, 57066, -152138), Array(2, 0, 1, 0, 53322, -170733), Array(2, -1, 0, 0, 45758, -204586), _
    Array(0, 1, -1, 0, -40923, -129620), Array(1, 0, 0, 0, -34720, 108743), Array(0, 1, 1, 0, -30383, 104755), Array(2, 0, 0, -2, 15327, 10321), _
    Array(0, 0, 1, 2, -12528, 0), Array(0, 0, 1, -2, 10980, 79661), Array(4, 0, -1, 0, 10675, -34782), Array(0, 0, 3, 0, 10034, -23210), _
    Array(4, 0, -2, 0, 8548, -21636), Array(2, 1, -1, 0, -7888, 24208), Array(2, 1, 0, 0, -6766, 30824), Array(1, 0, -1, 0, -5163, -8379), _
    Array(1, 1, 0, 0, 4987, -16675), Array(2, -1, 1, 0, 4036, -12831), Array(2, 0, 2, 0, 3994, -10445), Array(4, 0, 0, 0, 3861, -11650), _
    Array(2, 0, -3, 0, 3665, 14403), Array(0, 1, -2, 0, -2689, -7003), Array(2, 0, -1, 2, -2602, 0), Array(2, -1, -2, 0, 2390, 10056), _
    Array(1, 0, 1, 0, -2348, 6322), Array(2, -2, 0, 0, 2236, -9884), Array(0, 1, 2, 0, -2120, 5751), Array(0, 2, 0, 0, -2069, 0), _
    Array(2, -2, -1, 0, 2048, -4950), Array(2, 0, 1, -2, -1773, 4130), Array(2, 0, 0, 2, -1595, 0), Array(4, -1, -1, 0, 1215, -3958), _
    Array(0, 0, 2, 2, -1110, 0), Array(3, 0, -1, 0, -892, 3258), Array(2, 1, 1, 0, -810, 2616), Array(4, -1, -2, 0, 759, -1897), _
    Array(0, 2, -1, 0, -713, -2117), Array(2, 2, -1, 0, -700, 2354), Array(2, 1, -2, 0, 691, 0), Array(2, -1, 0, -2, 596, 0), _
    Array(4, 0, 1, 0, 549, -1423), Array(0, 0, 4, 0, 537, -1117), Array(4, -1, 0, 0, 520, -1571), Array(1, 0, -2, 0, -487, -1739), _
    Array(2, 1, 0, -2, -399, 0), Array(0, 0, 2, -2, -381, -4421), Array(1, 1, 1, 0, 351, 0), Array(3, 0, -2, 0, -340, 0), _
    Array(4, 0, -3, 0, 330, 0), Array(2, -1, 2, 0, 327, 0), Array(0, 2, 1, 0, -323, 1165), Array(1, 1, -1, 0, 299, 0), _
    Array(2, 0, 3, 0, 294, 0), Array(2, 0, -1, -2, 0, 8752))

    '{Table 45.B}
Dim MoonTermsB As Variant
'    MoonTermsB : array[1..MOONTABTERMS,1..5] of longint =
    MoonTermsB = Array(Array(0, 0, 0, 1, 5128122), Array(0, 0, 1, 1, 280602), Array(0, 0, 1, -1, 277693), _
    Array(2, 0, 0, -1, 173237), Array(2, 0, -1, 1, 55413), Array(2, 0, -1, -1, 46271), _
    Array(2, 0, 0, 1, 32573), Array(0, 0, 2, 1, 17198), Array(2, 0, 1, -1, 9266), _
    Array(0, 0, 2, -1, 8822), Array(2, -1, 0, -1, 8216), Array(2, 0, -2, -1, 4324), _
    Array(2, 0, 1, 1, 4200), Array(2, 1, 0, -1, -3359), Array(2, -1, -1, 1, 2463), _
    Array(2, -1, 0, 1, 2211), Array(2, -1, -1, -1, 2065), Array(0, 1, -1, -1, -1870), _
    Array(4, 0, -1, -1, 1828), Array(0, 1, 0, 1, -1794), Array(0, 0, 0, 3, -1749), _
    Array(0, 1, -1, 1, -1565), Array(1, 0, 0, 1, -1491), Array(0, 1, 1, 1, -1475), _
    Array(0, 1, 1, -1, -1410), Array(0, 1, 0, -1, -1344), Array(1, 0, 0, -1, -1335), _
    Array(0, 0, 3, 1, 1107), Array(4, 0, 0, -1, 1021), Array(4, 0, -1, 1, 833), _
    Array(0, 0, 1, -3, 777), Array(4, 0, -2, 1, 671), Array(2, 0, 0, -3, 607), _
    Array(2, 0, 2, -1, 596), Array(2, -1, 1, -1, 491), Array(2, 0, -2, 1, -451), _
    Array(0, 0, 3, -1, 439), Array(2, 0, 2, 1, 422), Array(2, 0, -3, -1, 421), _
    Array(2, 1, -1, 1, -366), Array(2, 1, 0, 1, -351), Array(4, 0, 0, 1, 331), _
    Array(2, -1, 1, 1, 315), Array(2, -2, 0, -1, 302), Array(0, 0, 1, 3, -283), _
    Array(2, 1, 1, -1, -229), Array(1, 1, 0, -1, 223), Array(1, 1, 0, 1, 223), _
    Array(0, 1, -2, -1, -220), Array(2, 1, -1, -1, -220), Array(1, 0, 1, 1, -185), _
    Array(2, -1, -2, -1, 181), Array(0, 1, 2, 1, -177), Array(4, 0, -2, -1, 176), _
    Array(4, -1, -1, -1, 166), Array(1, 0, 1, -1, -164), Array(4, 0, 1, -1, 132), _
    Array(1, 0, -1, -1, -119), Array(4, -1, 0, -1, 115), Array(2, -2, 0, 1, 107))

ld = modpi2(Eval4Poly(C4P(PLd), T) * DToR)
Call CalcMoonAngles(T, Om, D, M, Md, F, e(1))

a1 = (119.75 + 131.849 * T) * DToR
a2 = (53.09 + 479264.29 * T) * DToR
A3 = (313.45 + 481266.484 * T) * DToR

e(1) = 1 - T * (0.002516 + T * 0.0000074)
e(2) = e(1) * e(1)

Call CalcSinCosTab(D, 4, SinTab(1), CosTab(1))
Call CalcSinCosTab(M, 2, SinTab(2), CosTab(2))
Call CalcSinCosTab(Md, 4, SinTab(3), CosTab(3))
Call CalcSinCosTab(F, 3, SinTab(4), CosTab(4))

dl = 0: db = 0: dr = 0

'{ Calculate dl and dr }
For i = 1 To MOONTABTERMS
    Flag = True   '{First non-zero coefficient of one of the five anlges}
    For j = 1 To 4
        k = MoonTermsLR(i - 1)(j - 1)
        If (k <> 0) Then
            If (k < 0) Then k1 = -k Else k1 = k
            SinVal = SinTab(j).w(k1)
            If (k < 0) Then SinVal = -SinVal
            CosVal = CosTab(j).w(k1)
            If j = 2 Then
                SinVal = SinVal * e(k1)
                CosVal = CosVal * e(k1)
             End If
            If Flag Then
                ArgSin = SinVal
                ArgCos = CosVal
                Flag = False
            Else
                tmp = ArgSin * CosVal + ArgCos * SinVal
                ArgCos = ArgCos * CosVal - ArgSin * SinVal
                ArgSin = tmp
            End If
        End If
    Next
    dl = dl + MoonTermsLR(i - 1)(5 - 1) * ArgSin
    dr = dr + MoonTermsLR(i - 1)(6 - 1) * ArgCos
Next

'{ Calculate db }
For i = 1 To MOONTABTERMS
    Flag = True   '{First non-zero coefficient of one of the five anlges}
    For j = 1 To 4
        k = MoonTermsB(i - 1)(j - 1)
        If (k <> 0) Then
            If (k < 0) Then k1 = -k Else k1 = k
            SinVal = SinTab(j).w(k1)
            If (k < 0) Then SinVal = -SinVal
            CosVal = CosTab(j).w(k1)
            If j = 2 Then
                SinVal = SinVal * e(k1)
                CosVal = CosVal * e(k1)
            End If
            If Flag Then
                ArgSin = SinVal
                ArgCos = CosVal
                Flag = False
            Else
                tmp = ArgSin * CosVal + ArgCos * SinVal
                ArgCos = ArgCos * CosVal - ArgSin * SinVal
                ArgSin = tmp
            End If
        End If
    Next
    db = db + MoonTermsB(i - 1)(5 - 1) * ArgSin
Next
dl = dl + 3958 * Sin(a1) + 1962 * Sin(ld - F) + 318 * Sin(a2)
db = db + -2235 * Sin(ld) + 382 * Sin(A3) + 350 * Sin(a1) * Cos(F) + 127 * Sin(ld - Md) - 115 * Sin(ld + Md)
s.l = ld + 0.000001 * DToR * dl
If s.l < 0 Then s.l = s.l + 2 * Pi
s.B = 0.000001 * db * DToR
s.r = 385000.56 + 0.001 * dr
End Sub

