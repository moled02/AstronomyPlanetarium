Attribute VB_Name = "modEclipse"
Option Explicit
'(*****************************************************************************)
'(* Module: ECLIPSE.PAS                                                       *)
'(* Version 2.0                                                               *)
'(* Last modified: Februari, 1998                                             *)
'(* D. Molenkamp                                                              *)
'(*****************************************************************************)

'(*****************************************************************************)
'(* Name:    GeneralEclipseData                                               *)
'(* Type:    Support procedure                                                *)
'(* Purpose: Calculate the general data associated with an eclipse.           *)
'(* Arguments:                                                                *)
'(*   k : lunation number                                                     *)
'(*   JD : Julian Day corresponding to maximum eclipse                        *)
'(*   u, Gamma : see text.  If certainly no eclipse, Gamma = 2                *)
'(*   Md : mean anomaly of the Moon, only needed for lunar eclipses           *)
'(*****************************************************************************)

Sub GeneralEclipseData(ByVal k As Double, ByRef JD As Double, ByRef u As Double, ByRef Gamma As Double, ByRef Md As Double)
Dim T As Double
Dim Om As Double, D As Double, M As Double, F As Double, E   As Double
Dim F1 As Double, a1 As Double, sum As Double, P As Double, Q As Double, W   As Double

T = k / 1236.85
JD = 2451550.09765 + 29.530588853 * k
JD = JD + T * T * (0.0001337 + T * (-0.00000015 + T * 0.00000000073))
T = JDToT(JD)
Call CalcMoonAngles(T, Om, D, M, Md, F, E)
If (Abs(Sin(F)) < 0.36) Then
  '{ Possible eclipse }
  F1 = F - 0.02665 * DToR * Sin(Om)
  a1 = (299.77 + 0.107408 * k - 0.009173 * T * T) * DToR
  sum = _
    -0.4075 * Sin(Md) _
    + 0.1721 * E * Sin(M) _
    + 0.0161 * Sin(2 * Md) _
    - 0.0097 * Sin(2 * F1) _
    + 0.0073 * E * Sin(Md - M) _
    - 0.005 * E * Sin(Md + M) _
    - 0.0023 * Sin(Md - 2 * F1) _
    + 0.0021 * E * Sin(2 * M) _
    + 0.0012 * Sin(Md + 2 * F1) _
    + 0.0006 * E * Sin(2 * Md + M) _
    - 0.0004 * Sin(3 * Md) _
    - 0.0003 * E * Sin(M + 2 * F1) _
    + 0.0003 * Sin(a1) _
    - 0.0002 * E * Sin(M - 2 * F1) _
    - 0.0002 * E * Sin(2 * Md - M) _
    - 0.0002 * Sin(Om)
  JD = JD + sum
  P = _
    0.207 * E * Sin(M) _
    + 0.0024 * E * Sin(2 * M) _
    - 0.0392 * Sin(Md) _
    + 0.0116 * Sin(2 * Md) _
    - 0.0073 * E * Sin(Md + M) _
    + 0.0067 * E * Sin(Md - M) _
    + 0.0118 * Sin(2 * F1)
  Q = _
    5.2207 _
    - 0.0048 * E * Cos(M) _
    + 0.002 * E * Cos(2 * M) _
    - 0.3299 * Cos(Md) _
    - 0.006 * E * Cos(Md + M) _
    + 0.0041 * E * Cos(Md - M)
  W = Abs(Cos(F1))
  '{ We aren't interested here if the eclipse occurs }
  '{ in the northern or southern hemisphere, so we   }
  '{ take the absolute value.                        }
  Gamma = (P * Cos(F1) + Q * Sin(F1)) * (1 - 0.0048 * W)
  u = _
    0.0059 _
    + 0.0046 * E * Cos(M) _
    - 0.0182 * Cos(Md) _
    + 0.0004 * Cos(2 * Md) _
    - 0.0005 * Cos(Md + M)
Else
  u = 0
  Gamma = 2
End If
End Sub

'(*****************************************************************************)
'(* Name:    NextSolarEclipse                                                 *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Find the next solar eclipse following a given Julian Day.        *)
'(* Arguments:                                                                *)
'(*   JD : lower limit for Julian Day                                         *)
'(*   Solareclips : data of the eclips                                        *)
'(*   Max : maximum magnitude in case of a partial eclipse                    *)
'(*   EclipseType : one of TOTAL, ANNULAR, ANNULARTOTAL, NOT_CENTRAL or       *)
'(*                 PARTIAL                                                   *)
'(*   Gamma :
'(*****************************************************************************)

Sub NextSolarEclipse(ByVal JD As Double, ByRef SolarEclipse As SOLARECLIPSEDATA)
Dim k As Double, u As Double, Gamma As Double, Dummy As Double, Max As Double
Dim EclipseType As Long


k = floor((JD - 2451550.09765) / 29.530588853 + 1.03)
Do Until Max > 0
    Gamma = 9999
    Do Until Gamma < 1.5433 + u
        Call GeneralEclipseData(k, JD, u, Gamma, Dummy)
        k = k + 1
    Loop
    Max = 0
    If Abs(Gamma) < 0.9972 Then
        If u < 0 Then
           EclipseType = TOTAL
        ElseIf u > 0.0047 Then
           EclipseType = ANNULAR
        ElseIf u < 0.00464 * Sqr(1 - Gamma * Gamma) Then
           EclipseType = ANNULARTOTAL
        Else
           EclipseType = ANNULAR
        End If
    Else
        If (Abs(Gamma) < 0.9972 + Abs(u)) Then
           If (u > 0.0047) Then
              EclipseType = ANNULAR_NOT_CENTRAL
           Else
              EclipseType = TOTAL_NOT_CENTRAL
           End If
        Else
            EclipseType = PARTIAL
        End If
    End If
    Max = (1.5433 + u - Abs(Gamma)) / (0.5461 + 2 * u)
Loop
SolarEclipse.Gamma = Gamma
SolarEclipse.JD = JD
SolarEclipse.Maxmag = Max
SolarEclipse.EclipseType = EclipseType
End Sub
Sub LastSolarEclipse(ByVal JD As Double, ByRef SolarEclipse As SOLARECLIPSEDATA)
Dim k As Double, u As Double, Gamma As Double, Dummy As Double, Max As Double
Dim EclipseType As Long


k = floor((JD - 2451550.09765) / 29.530588853 - 0.03)
Do Until Max > 0
    Gamma = 9999
    Do Until Gamma < 1.5433 + u
        Call GeneralEclipseData(k, JD, u, Gamma, Dummy)
        k = k - 1
    Loop
    Max = 0
    If Abs(Gamma) < 0.9972 Then
        If u < 0 Then
           EclipseType = TOTAL
        ElseIf u > 0.0047 Then
           EclipseType = ANNULAR
        ElseIf u < 0.00464 * Sqr(1 - Gamma * Gamma) Then
           EclipseType = ANNULARTOTAL
        Else
           EclipseType = ANNULAR
        End If
    Else
        If (Abs(Gamma) < 0.9972 + Abs(u)) Then
           If (u > 0.0047) Then
              EclipseType = ANNULAR_NOT_CENTRAL
           Else
              EclipseType = TOTAL_NOT_CENTRAL
           End If
        Else
            EclipseType = PARTIAL
        End If
    End If
    Max = (1.5433 + u - Abs(Gamma)) / (0.5461 + 2 * u)
Loop
SolarEclipse.Gamma = Gamma
SolarEclipse.JD = JD
SolarEclipse.Maxmag = Max
SolarEclipse.EclipseType = EclipseType
End Sub
'(*****************************************************************************)
'(* Name:    NextLunarEclipse                                                 *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Find the next lunar eclipse following a given Julian Day.        *)
'(* Arguments:                                                                *)
'(*   JD : lower limit for Julian Day                                         *)
'(*   System : normally 0, TRADITIONAL if the 'traditional' system of         *)
'(*            constants is to be used                                        *)
'(*   LunarEclipse : contains data about the lunareclips, in which             *)
'(*   JD : Julian Day at maximum eclips                                       *)
'(*   MagPenum, MagUmbra : magnitude eclips in Penumbra, Umbra                *)
'(*   SpartUmbra, StotalUmbra : semi-duration of partial, total phase, in     *)
'(*                             parts of one day                              *)
'(*   SpartPenum : semi-duration of penumbral eclips, in parts of one day     *)
'(*   EclipseType : one of TOTAL, PARTIAL or PENUMBRAL                        *)
'(*****************************************************************************)

Sub NextLunarEclipse(ByVal JD As Double, ByVal System As Long, _
                           ByRef LunarEclipse As LUNARECLIPSEDATA)
Dim k As Double, u As Double, Max As Double, g As Double, Md As Double
Dim H As Double, P As Double, T As Double, n As Double

Dim LunarEclipseParams

'LunarEclipseParams : array(0..1, 1..5) of real =
LunarEclipseParams = Array( _
    Array(0, 1.2848, 0.7403, 1.5573, 1.0128, 0.4678), _
    Array(0, 1.2985, 0.7432, 1.571, 1.0157, 0.4707))

k = floor((JD - 2451550.09765) / 29.530588853 + 0.53) + 0.5
Do Until Max > 0
  Call GeneralEclipseData(k, JD, u, g, Md)
  g = Abs(g)
  k = k + 1
  Max = (LunarEclipseParams(System)(3) + u - g)
Loop
n = 24 * (0.5458 + 0.04 * Cos(Md))

With LunarEclipse
  .SpartPenumbra = 0
  .SpartUmbra = 0
  .StotUmbra = 0
  P = LunarEclipseParams(System)(4) - u
  T = LunarEclipseParams(System)(5) - u
  H = LunarEclipseParams(System)(3) + u
  .sumbra = (LunarEclipseParams(System)(2) + u) * 3.6698
  .spenumbra = (LunarEclipseParams(System)(1) - u) * 3.6698
  .MagPenumbra = (H - g) / 0.545
  .MagUmbra = (P - g) / 0.545

  .SpartPenumbra = Sqr(H * H - g * g) / n
  If .MagPenumbra < 1 Then .EclipseType = PARTPENUMBRAL

  If .MagUmbra > 0 Then
      .SpartUmbra = Sqr(P * P - g * g) / n
      If .MagUmbra > 1 Then
          .EclipseType = TOTAL
          .StotUmbra = Sqr(T * T - g * g) / n
      Else
          .EclipseType = PARTIAL
      End If
   Else
       .EclipseType = PENUMBRAL
   End If
End With
LunarEclipse.JD = JD
End Sub

Sub LastLunarEclipse(ByVal JD As Double, ByVal System As Long, _
                           ByRef LunarEclipse As LUNARECLIPSEDATA)
Dim k As Double, u As Double, Max As Double, g As Double, Md As Double
Dim H As Double, P As Double, T As Double, n As Double

Dim LunarEclipseParams

'LunarEclipseParams : array(0..1, 1..5) of real =
LunarEclipseParams = Array( _
    Array(0, 1.2848, 0.7403, 1.5573, 1.0128, 0.4678), _
    Array(0, 1.2985, 0.7432, 1.571, 1.0157, 0.4707))

k = floor((JD - 2451550.09765) / 29.530588853 - 0.53) + 0.5
Do Until Max > 0
  Call GeneralEclipseData(k, JD, u, g, Md)
  g = Abs(g)
  k = k - 1
  Max = (LunarEclipseParams(System)(3) + u - g)
Loop
n = 24 * (0.5458 + 0.04 * Cos(Md))

With LunarEclipse
  .SpartPenumbra = 0
  .SpartUmbra = 0
  .StotUmbra = 0
  P = LunarEclipseParams(System)(4) - u
  T = LunarEclipseParams(System)(5) - u
  H = LunarEclipseParams(System)(3) + u
  .sumbra = (LunarEclipseParams(System)(2) + u) * 3.6698
  .spenumbra = (LunarEclipseParams(System)(1) - u) * 3.6698
  .MagPenumbra = (H - g) / 0.545
  .MagUmbra = (P - g) / 0.545

  .SpartPenumbra = Sqr(H * H - g * g) / n
  If .MagPenumbra < 1 Then .EclipseType = PARTPENUMBRAL

  If .MagUmbra > 0 Then
      .SpartUmbra = Sqr(P * P - g * g) / n
      If .MagUmbra > 1 Then
          .EclipseType = TOTAL
          .StotUmbra = Sqr(T * T - g * g) / n
      Else
          .EclipseType = PARTIAL
      End If
   Else
       .EclipseType = PENUMBRAL
   End If
End With
LunarEclipse.JD = JD
End Sub
'{ Vanaf hier (c) 1997 D.A.M. Molenkamp }
Sub Bess_elmts(RkM As Double, DecM As Double, ParM As Double, RkZ As Double, DecZ As Double, ParZ As Double, RZ As Double, EphTime As Double, ByRef BessElmt As tBessElmt)
Dim sRkM As Double, cRkM As Double, sDecM As Double, cDecM As Double, sRkZ As Double, cRkZ As Double, sDecZ As Double, cDecz As Double, _
   sParZ0 As Double, ParZ0 As Double, sParM As Double, RsParM As Double, RM As Double, B As Double, _
   cDecZcRkZ As Double, cDecZsRkZ As Double, BcDecMcRkM As Double, BcDecMsRkM As Double, _
   GcDcA As Double, GcDsA As Double, GsD As Double, tA As Double, A As Double, g As Double, GR As Double, sF1 As Double, sF2 As Double, _
   k As Double, KcsF1 As Double, KcsF2 As Double, C1 As Double, C2  As Double, _
   Z As Double
Dim k1 As Double, k2 As Double

 ParZ0 = 4.26352084003365E-05
 With BessElmt
   sRkM = Sin(RkM): cRkM = Cos(RkM)
   sDecM = Sin(DecM): cDecM = Cos(DecM)
   sRkZ = Sin(RkZ): cRkZ = Cos(RkZ)
   sDecZ = Sin(DecZ): cDecz = Cos(DecZ)

   sParZ0 = Sin(ParZ0)
   sParM = Sin(ParM)
   RsParM = RZ * sParM
   RM = 1 / sParM
   B = sParZ0 / RsParM

   cDecZcRkZ = cDecz * cRkZ: cDecZsRkZ = cDecz * sRkZ
   BcDecMcRkM = B * cDecM * cRkM: BcDecMsRkM = B * cDecM * sRkM
   GcDcA = cDecZcRkZ - BcDecMcRkM
   GcDsA = cDecZsRkZ - BcDecMsRkM
   GsD = sDecZ - B * sDecM

   tA = GcDsA / GcDcA
   A = atan2(GcDsA, GcDcA)
   g = Sqr(GcDcA * GcDcA + GcDsA * GcDsA + GsD * GsD)
   .SD = GsD / g
   .cD = GcDsA / (g * Sin(A))

   .x = RM * (cDecM * Sin(RkM - A))
   .Y = RM * (sDecM * BessElmt.cD - cDecM * BessElmt.SD * Cos(RkM - A))
   Z = RM * (sDecM * BessElmt.SD + cDecM * BessElmt.cD * Cos(RkM - A))

   GR = g * RZ
   'constanten bepaald op: pi0 = 8.794143; afst zon = 696000km, 1 ea =1.495978...
   sF1 = 4.66404070395725E-03 / GR
   sF2 = 4.64081352846916E-03 / GR
'   k = 0.272488
   'k1 = Lunar radius Penumbra; k2 = Lunar radius Umbra (vanaf Fred Espenak)
   k1 = 0.2725076: k2 = 0.272281
   KcsF1 = k1 / sF1
   KcsF2 = k2 / sF2

   C1 = Z + KcsF1
   C2 = Z - KcsF2

   .tF1 = tan(asin(sF1))
   .tF2 = tan(asin(sF2))

   .l1 = C1 * BessElmt.tF1
   .l2 = C2 * BessElmt.tF2

   .mu = EphTime - A
   If .mu < 0 Then
      .mu = .mu + Pi2
   End If
End With
End Sub

Sub Aux_elmts(BessElmt As tBessElmt, ByRef AuxElmt As tAuxElmt, ByRef DiffBess As tDiffBess)
Dim e2c2d As Double, e2s2d As Double, e2sdcd As Double, rho12 As Double, rho22 As Double, rho1rho2  As Double

   e2c2d = e2 * BessElmt.cD * BessElmt.cD
   e2s2d = e2 * BessElmt.SD * BessElmt.SD
   e2sdcd = e2 * BessElmt.SD * BessElmt.cD
   
   rho12 = 1 - e2c2d
   rho22 = 1 - e2s2d
   AuxElmt.rho1 = Sqr(rho12)
   AuxElmt.rho2 = Sqr(rho22)
   rho1rho2 = AuxElmt.rho1 * AuxElmt.rho2
   AuxElmt.sd1 = BessElmt.SD / AuxElmt.rho1
   AuxElmt.cd1 = BessElmt.cD * Sqr(1 - e2) / AuxElmt.rho1
   AuxElmt.sd1_d2 = e2sdcd / rho1rho2
   AuxElmt.cd1_d2 = Sqr(1 - e2) / rho1rho2
   DiffBess.a11 = -DiffBess.l11 - DiffBess.mu1 * BessElmt.x * BessElmt.tF1 * BessElmt.cD
   DiffBess.a21 = -DiffBess.l21 - DiffBess.mu1 * BessElmt.x * BessElmt.tF2 * BessElmt.cD
   DiffBess.b1 = -DiffBess.y1 + DiffBess.mu1 * BessElmt.x * BessElmt.SD
   DiffBess.c11 = DiffBess.x1 + DiffBess.mu1 * BessElmt.Y * BessElmt.SD + DiffBess.mu1 * BessElmt.l1 * BessElmt.tF1 * BessElmt.cD
   DiffBess.c21 = DiffBess.x1 + DiffBess.mu1 * BessElmt.Y * BessElmt.SD + DiffBess.mu1 * BessElmt.l2 * BessElmt.tF2 * BessElmt.cD
End Sub

Sub DiffBess(Bess0 As tBessElmt, Bess1 As tBessElmt, ByRef dBess As tDiffBess)
    dBess.mu1 = modpi(Bess1.mu - Bess0.mu)
    dBess.x1 = Bess1.x - Bess0.x
    dBess.y1 = Bess1.Y - Bess0.Y
    dBess.l11 = modpi(Bess1.l1 - Bess0.l1)
    dBess.l21 = modpi(Bess1.l2 - Bess0.l2)
    dBess.d1 = asin(Bess1.SD) - asin(Bess0.SD)
End Sub

Function PredDataSolarEcl(BessElmt As tBessElmt, AuxElmt As tAuxElmt, dBess As tDiffBess, ByRef PredData As tPredData) As Boolean
Dim cPhi1sDel As Double, cPhi1cDel As Double, x1_Xi1 As Double, y1_nu1 As Double
Dim sPhi1 As Double, tPhi1 As Double

On Error GoTo fout_PredDataSolarEcl
    PredData.nu1 = BessElmt.Y / AuxElmt.rho1
    PredData.Xi = BessElmt.x
    PredData.Psi1 = Sqr(1 - PredData.Xi * PredData.Xi - PredData.nu1 * PredData.nu1)

    cPhi1sDel = PredData.Xi
    cPhi1cDel = -PredData.nu1 * AuxElmt.sd1 + PredData.Psi1 * AuxElmt.cd1
    PredData.Psi = AuxElmt.rho2 * (PredData.Psi1 * AuxElmt.cd1_d2 - PredData.nu1 * AuxElmt.sd1_d2)
    PredData.l2 = BessElmt.l2 - PredData.Psi * BessElmt.tF2
    PredData.Del = atan2(cPhi1sDel, cPhi1cDel)
    x1_Xi1 = dBess.c21 - dBess.mu1 * PredData.Psi * BessElmt.cD
    y1_nu1 = -dBess.b1
    PredData.n = Sqr(x1_Xi1 * x1_Xi1 + y1_nu1 * y1_nu1)
    
    PredData.s = Abs(PredData.l2 / PredData.n)
    sPhi1 = PredData.nu1 * AuxElmt.cd1 + PredData.Psi1 * AuxElmt.sd1
    tPhi1 = sPhi1 * Cos(PredData.Del) / cPhi1cDel
    PredData.tPhi = tPhi1 / Sqr(1 - e2)
    PredData.Phi = Atn(PredData.tPhi)
    PredData.lambda = BessElmt.mu - PredData.Del
    PredData.tQ0 = -(y1_nu1) / (x1_Xi1)
    PredDataSolarEcl = True
    Exit Function
    
fout_PredDataSolarEcl:
    PredDataSolarEcl = False
End Function

Function LimitsUmbraN(BessElmt As tBessElmt, AuxElmt As tAuxElmt, DiffBess As tDiffBess, PredData As tPredData, _
                              ByRef limits As tLimits) As Boolean
Dim tQ As Double, sQ As Double, cQ As Double, PsitF2 As Double, oPsi As Double, nPsi  As Double, _
    lXi As Double, cQqRho1 As Double, l2sQ As Double, L2cQqRho1 As Double, lNu1 As Double, lPsi12 As Double, lPsi1 As Double, _
    cPhi1sDel As Double, cPhi1cDel As Double, tDel  As Double, lDel As Double, sPhi1 As Double, ltPhi1 As Double, ltPhi As Double, _
    PsitF1 As Double, l1sQ As Double, L1cQqRho1  As Double
Dim I As Long
On Error GoTo fout_LimitsUmbraN
'{-----------------------UMBRA--------------------------}
      '{Noordelijke benadering}

    oPsi = 1000: nPsi = PredData.Psi
    tQ = DiffBess.b1 / (DiffBess.c21 - nPsi * DiffBess.mu1 * BessElmt.cD)
    sQ = Sin(Atn(tQ))
    cQ = sQ / tQ
    cQqRho1 = cQ / AuxElmt.rho1
    
    For I = 1 To 50
       oPsi = nPsi
       PsitF2 = oPsi * BessElmt.tF2
       limits.l2 = BessElmt.l2 - PsitF2

       l2sQ = limits.l2 * sQ
       lXi = BessElmt.x - l2sQ

       L2cQqRho1 = limits.l2 * cQ / AuxElmt.rho1
       lNu1 = BessElmt.Y / AuxElmt.rho1 - L2cQqRho1
       If Abs(lNu1) <= 1 Then
            lPsi12 = 1 - lXi * lXi - lNu1 * lNu1
            lPsi1 = Sqr(lPsi12)
            nPsi = AuxElmt.rho2 * (lPsi1 * AuxElmt.cd1_d2 - lNu1 * AuxElmt.sd1_d2)
            tQ = (DiffBess.b1 - nPsi * DiffBess.d1 - DiffBess.a21 / cQ) / (DiffBess.c21 - nPsi * DiffBess.mu1 * BessElmt.cD)
            sQ = Sin(Atn(tQ))
            cQ = sQ / tQ
            cQqRho1 = cQ / AuxElmt.rho1
       End If
       If Abs(oPsi - nPsi) <= 0.0002 Then Exit For
    Next
    If Abs(lNu1) <= 1 Then 'And Abs(oPsi - nPsi) <= 0.0002 Then
         cPhi1sDel = lXi
         cPhi1cDel = -lNu1 * AuxElmt.sd1 + lPsi1 * AuxElmt.cd1
         tDel = cPhi1sDel / cPhi1cDel
         lDel = atan2(cPhi1sDel, cPhi1cDel)
         limits.ULimN.lng = BessElmt.mu - lDel
         sPhi1 = lNu1 * AuxElmt.cd1 + lPsi1 * AuxElmt.sd1
         ltPhi1 = sPhi1 * Cos(lDel) / cPhi1cDel
         ltPhi = ltPhi1 / Sqr(1 - e2)
         limits.ULimN.nb = Atn(ltPhi)
    Else
        LimitsUmbraN = False
        Exit Function
    End If
    LimitsUmbraN = True
    Exit Function
    
fout_LimitsUmbraN:
    LimitsUmbraN = False
End Function
Function LimitsUmbraZ(BessElmt As tBessElmt, AuxElmt As tAuxElmt, DiffBess As tDiffBess, PredData As tPredData, _
                              ByRef limits As tLimits) As Boolean
Dim tQ As Double, sQ As Double, cQ As Double, PsitF2 As Double, oPsi As Double, nPsi  As Double, _
    lXi As Double, cQqRho1 As Double, l2sQ As Double, L2cQqRho1 As Double, lNu1 As Double, lPsi12 As Double, lPsi1 As Double, _
    cPhi1sDel As Double, cPhi1cDel As Double, tDel  As Double, lDel As Double, sPhi1 As Double, ltPhi1 As Double, ltPhi As Double, _
    PsitF1 As Double, l1sQ As Double, L1cQqRho1  As Double
Dim vVorigVerschil As Double
Dim I As Long
On Error GoTo fout_LimitsUmbraZ
'{-----------------------UMBRA--------------------------}
      '{Zuidelijke benadering}

    oPsi = 1000: nPsi = PredData.Psi
    tQ = DiffBess.b1 / (DiffBess.c21 - nPsi * DiffBess.mu1 * BessElmt.cD)
    sQ = -Sin(Atn(tQ))
    cQ = sQ / tQ
    cQqRho1 = cQ / AuxElmt.rho1

    For I = 1 To 50
 '   Do While Abs(oPsi - nPsi) > 0.0002 And Abs(oPsi - nPsi) <= vVorigVerschil
'       vVorigVerschil = Abs(oPsi - nPsi)
       oPsi = nPsi
       PsitF2 = oPsi * BessElmt.tF2
       limits.l2 = BessElmt.l2 - PsitF2


       l2sQ = limits.l2 * sQ
       lXi = BessElmt.x - l2sQ

       L2cQqRho1 = limits.l2 * cQ / AuxElmt.rho1
       lNu1 = BessElmt.Y / AuxElmt.rho1 - L2cQqRho1
       If Abs(lNu1) <= 1 Then
            lPsi12 = 1 - lXi * lXi - lNu1 * lNu1
            lPsi1 = Sqr(lPsi12)
            nPsi = AuxElmt.rho2 * (lPsi1 * AuxElmt.cd1_d2 - lNu1 * AuxElmt.sd1_d2)
            tQ = (DiffBess.b1 - nPsi * DiffBess.d1 - DiffBess.a21 / cQ) / (DiffBess.c21 - nPsi * DiffBess.mu1 * BessElmt.cD)
            sQ = -Sin(Atn(tQ))
            cQ = sQ / tQ
            cQqRho1 = cQ / AuxElmt.rho1
       End If
'    Loop
        If Abs(oPsi - nPsi) <= 0.0002 Then Exit For
    Next
    If Abs(lNu1) <= 1 And Abs(oPsi - nPsi) <= 0.0002 Then
         cPhi1sDel = lXi
         cPhi1cDel = -lNu1 * AuxElmt.sd1 + lPsi1 * AuxElmt.cd1
         tDel = cPhi1sDel / cPhi1cDel
         lDel = atan2(cPhi1sDel, cPhi1cDel)
         limits.ULimZ.lng = BessElmt.mu - lDel
         sPhi1 = lNu1 * AuxElmt.cd1 + lPsi1 * AuxElmt.sd1
         ltPhi1 = sPhi1 * Cos(lDel) / cPhi1cDel
         ltPhi = ltPhi1 / Sqr(1 - e2)
         limits.ULimZ.nb = Atn(ltPhi)
    Else
        LimitsUmbraZ = False
        Exit Function
    End If
    LimitsUmbraZ = True
    Exit Function
    
fout_LimitsUmbraZ:
    LimitsUmbraZ = False
End Function

Function LimitsPenumbraN(BessElmt As tBessElmt, AuxElmt As tAuxElmt, DiffBess As tDiffBess, PredData As tPredData, _
                              ByRef limits As tLimits) As Boolean
Dim tQ As Double, sQ As Double, cQ As Double, PsitF2 As Double, oPsi As Double, nPsi  As Double, _
    lXi As Double, cQqRho1 As Double, l2sQ As Double, L2cQqRho1 As Double, lNu1 As Double, lPsi12 As Double, lPsi1 As Double, _
    cPhi1sDel As Double, cPhi1cDel As Double, tDel  As Double, lDel As Double, sPhi1 As Double, ltPhi1 As Double, ltPhi As Double, _
    PsitF1 As Double, l1sQ As Double, L1cQqRho1  As Double
Dim I As Long
On Error GoTo fout_LimitsPenumbraN
'{-----------------------PENUMBRA--------------------------}
      '{Noordelijke benadering}

    oPsi = 1000: nPsi = PredData.Psi
    tQ = DiffBess.b1 / (DiffBess.c11 - nPsi * DiffBess.mu1 * BessElmt.cD)
    sQ = -Sin(Atn(tQ))
    cQ = sQ / tQ
    cQqRho1 = cQ / AuxElmt.rho1

    For I = 1 To 50
       oPsi = nPsi
       PsitF1 = oPsi * BessElmt.tF1
       limits.l1 = BessElmt.l1 - PsitF1


       l1sQ = limits.l1 * sQ
       lXi = BessElmt.x - l1sQ

       L1cQqRho1 = limits.l1 * cQ / AuxElmt.rho1
       lNu1 = BessElmt.Y / AuxElmt.rho1 - L1cQqRho1
       If Abs(lNu1) <= 1 Then
          lPsi12 = 1 - lXi * lXi - lNu1 * lNu1
          lPsi1 = Sqr(lPsi12)
          nPsi = AuxElmt.rho2 * (lPsi1 * AuxElmt.cd1_d2 - lNu1 * AuxElmt.sd1_d2)
          tQ = (DiffBess.b1 - nPsi * DiffBess.d1 - DiffBess.a21 / cQ) / (DiffBess.c11 - nPsi * DiffBess.mu1 * BessElmt.cD)
          sQ = -Sin(Atn(tQ))
          cQ = sQ / tQ
          cQqRho1 = cQ / AuxElmt.rho1
          cPhi1sDel = lXi
       End If
       If Abs(oPsi - nPsi) <= 0.0002 Then Exit For
    Next
    If Abs(lNu1) <= 1 Then
         cPhi1cDel = -lNu1 * AuxElmt.sd1 + lPsi1 * AuxElmt.cd1
         tDel = cPhi1sDel / cPhi1cDel
         lDel = atan2(cPhi1sDel, cPhi1cDel)
         limits.PLimN.lng = BessElmt.mu - lDel
         sPhi1 = lNu1 * AuxElmt.cd1 + lPsi1 * AuxElmt.sd1
         ltPhi1 = sPhi1 * Cos(lDel) / cPhi1cDel
         ltPhi = ltPhi1 / Sqr(1 - e2)
         limits.PLimN.nb = Atn(ltPhi)
    Else
        LimitsPenumbraN = False
        Exit Function
    End If
    LimitsPenumbraN = True
    Exit Function
    
fout_LimitsPenumbraN:
    LimitsPenumbraN = False
End Function

Function LimitsPenumbraZ(BessElmt As tBessElmt, AuxElmt As tAuxElmt, DiffBess As tDiffBess, PredData As tPredData, _
                              ByRef limits As tLimits) As Boolean
Dim tQ As Double, sQ As Double, cQ As Double, PsitF2 As Double, oPsi As Double, nPsi  As Double, _
    lXi As Double, cQqRho1 As Double, l2sQ As Double, L2cQqRho1 As Double, lNu1 As Double, lPsi12 As Double, lPsi1 As Double, _
    cPhi1sDel As Double, cPhi1cDel As Double, tDel  As Double, lDel As Double, sPhi1 As Double, ltPhi1 As Double, ltPhi As Double, _
    PsitF1 As Double, l1sQ As Double, L1cQqRho1  As Double
Dim I As Long
On Error GoTo fout_LimitsPenumbraZ
'{-----------------------PENUMBRA--------------------------}
      '{Zuidelijke benadering}

    oPsi = 1000: nPsi = 0
    tQ = DiffBess.b1 / (DiffBess.c11 - nPsi * DiffBess.mu1 * BessElmt.cD)
    sQ = Sin(Atn(tQ))
    cQ = sQ / tQ
    cQqRho1 = cQ / AuxElmt.rho1

    For I = 1 To 50
       oPsi = nPsi
       PsitF1 = oPsi * BessElmt.tF1
       limits.l1 = BessElmt.l1 - PsitF1


       l1sQ = limits.l1 * sQ
       lXi = BessElmt.x - l1sQ

       L1cQqRho1 = limits.l1 * cQ
       lNu1 = BessElmt.Y - L1cQqRho1

       If Abs(lNu1) <= 1 Then
            lPsi12 = 1 - lXi * lXi - lNu1 * lNu1
            lPsi1 = Sqr(lPsi12)
            nPsi = AuxElmt.rho2 * (lPsi1 * AuxElmt.cd1_d2 - lNu1 * AuxElmt.sd1_d2)
            tQ = (DiffBess.b1 - nPsi * DiffBess.d1 - DiffBess.a21 / cQ) / (DiffBess.c11 - nPsi * DiffBess.mu1 * BessElmt.cD)
            sQ = Sin(Atn(tQ))
            cQ = sQ / tQ
            cQqRho1 = cQ / AuxElmt.rho1
       End If
       If Abs(oPsi - nPsi) <= 0.0002 Then Exit For
    Next
    If Abs(lNu1) <= 1 Then
         cPhi1sDel = lXi
         cPhi1cDel = -lNu1 * AuxElmt.sd1 + lPsi1 * AuxElmt.cd1
         tDel = cPhi1sDel / cPhi1cDel
         lDel = atan2(cPhi1sDel, cPhi1cDel)
         limits.PLimZ.lng = BessElmt.mu - lDel
         sPhi1 = lNu1 * AuxElmt.cd1 + lPsi1 * AuxElmt.sd1
         ltPhi1 = sPhi1 * Cos(lDel) / cPhi1cDel
         ltPhi = ltPhi1 / Sqr(1 - e2)
         limits.PLimZ.nb = Atn(ltPhi)
    Else
        LimitsPenumbraZ = False
        Exit Function
    End If
    LimitsPenumbraZ = True
    Exit Function
    
fout_LimitsPenumbraZ:
    LimitsPenumbraZ = False
End Function

Function LimitsUmbraPenumbra(BessElmt As tBessElmt, AuxElmt As tAuxElmt, DiffBess As tDiffBess, PredData As tPredData, _
                              ByRef limits As tLimits) As Long
'0 = geen waarden gevonden
'1 = n umbra
'2 = z umbra
'4 = n penumbra
'8 = z penumbra
'3 = umbra n + z gevonden etc.
Dim tQ As Double, sQ As Double, cQ As Double, PsitF2 As Double, oPsi As Double, nPsi  As Double, _
    lXi As Double, cQqRho1 As Double, l2sQ As Double, L2cQqRho1 As Double, lNu1 As Double, lPsi12 As Double, lPsi1 As Double, _
    cPhi1sDel As Double, cPhi1cDel As Double, tDel  As Double, lDel As Double, sPhi1 As Double, ltPhi1 As Double, ltPhi As Double, _
    PsitF1 As Double, l1sQ As Double, L1cQqRho1  As Double
Dim nRes As Long
Dim sUmbraPenumbra As String
' WITH BessElmt, AuxElmt, DiffBess, PredData, Limits DO
    limits.ULimN.nb = 0: limits.ULimN.lng = 0: limits.ULimZ.nb = 0: limits.ULimZ.lng = 0

    nRes = 15
    
'{-----------------------UMBRA--------------------------}
      '{Noordelijke benadering}
    sUmbraPenumbra = "Umbra N"
    If Not LimitsUmbraN(BessElmt, AuxElmt, DiffBess, PredData, limits) Then
        nRes = nRes - 1
    End If
    
    '{Nu voor de zuidelijke}
    sUmbraPenumbra = "Umbra Z"
    If Not LimitsUmbraZ(BessElmt, AuxElmt, DiffBess, PredData, limits) Then
        nRes = nRes - 2
    End If
    

'{--------------------------PENUMBRA----------------------------}
      '{Noordelijke benadering}
    sUmbraPenumbra = "Penumbra N"
    If Not LimitsPenumbraN(BessElmt, AuxElmt, DiffBess, PredData, limits) Then
        nRes = nRes - 4
    End If
    
    '{Nu voor de zuidelijke}
    sUmbraPenumbra = "Penumbra Z"
    If Not LimitsPenumbraZ(BessElmt, AuxElmt, DiffBess, PredData, limits) Then
        nRes = nRes - 8
    End If

    LimitsUmbraPenumbra = nRes
End Function

Function Outline1Curve(BessElmt As tBessElmt, AuxElmt As tAuxElmt, DiffBess As tDiffBess, PredData As tPredData, _
                              ByRef OutCurve As tOutCurve) As Boolean
Dim Tm As Double, M As Double, cQ_M  As Double
Dim hM As Double, Q_M1 As Double, Q_M2 As Double, Q1 As Double, Q2 As Double, tempQ   As Double


On Error GoTo fout_Outline1Curve

Tm = BessElmt.x / BessElmt.Y
hM = atan2(BessElmt.x, BessElmt.Y)
M = BessElmt.x / Sin(hM)
cQ_M = (M * M + BessElmt.l1 * BessElmt.l1 - 1) / (2 * BessElmt.l1 * M)
If Abs(cQ_M) > 1 Then
    Outline1Curve = False
    OutCurve.bQ = 0
    OutCurve.eQ = 2 * Pi
Else
    Q_M1 = acos(cQ_M)
    Q_M2 = -acos(cQ_M)
    Q1 = modpi2(Q_M1 + hM)
    Q2 = modpi2(Q_M2 + hM)
    If Q1 > Q2 Then
        tempQ = Q1
        Q1 = Q2
        Q2 = tempQ
    End If
    '1 =  (x - l1*sin(Q))^2 + (y-l1*cos(Q))^2 + phi^2
    'indien geldig bij Q = 0, dan zit Q=0 dus in het bereik, dan moet van Q2 2pi afgetrokken worden en wordt bovenwaarde
    If 1 - BessElmt.x * BessElmt.x - (BessElmt.Y - BessElmt.l1) * (BessElmt.Y - BessElmt.l1) > 0 Then
        OutCurve.eQ = Q1
        OutCurve.bQ = Q2 - Pi2
    Else
        OutCurve.bQ = Q1
        OutCurve.eQ = Q2
    End If
    Outline1Curve = True
End If

Exit Function

fout_Outline1Curve:
    Outline1Curve = False
End Function

Function Outline2Curve(BessElmt As tBessElmt, AuxElmt As tAuxElmt, DiffBess As tDiffBess, PredData As tPredData, _
                              HoekQ As Double, ByRef OutCurve As tOutCurve) As Boolean

Dim lsQ As Double, lcQ As Double, lXi As Double, lNu As Double, lPsi As Double, lPsi2  As Double
Dim cPhisDel As Double, cPhicDel As Double, tDel As Double, lDel As Double, sPhi1 As Double, ltPhi1 As Double, ltPhi     As Double


On Error GoTo fout_Outline2Curve

    lsQ = Sin(HoekQ)
    lcQ = Cos(HoekQ)

    lXi = BessElmt.x - BessElmt.l1 * lsQ
    lNu = BessElmt.Y - BessElmt.l1 * lcQ

    lPsi2 = 1 - lXi * lXi - lNu * lNu
    lPsi = Sqr(lPsi2)

    cPhisDel = lXi
    cPhicDel = -lNu * AuxElmt.sd1 + lPsi * AuxElmt.cd1
    tDel = cPhisDel / cPhicDel
    lDel = atan2(cPhisDel, cPhicDel)
    OutCurve.pos.lng = modpi2(BessElmt.mu - lDel)
    sPhi1 = lNu * AuxElmt.cd1 + lPsi * AuxElmt.sd1
    ltPhi1 = sPhi1 * Cos(lDel) / cPhicDel
    ltPhi = ltPhi1 / Sqr(1 - e2)
    OutCurve.pos.nb = Atn(ltPhi)
    Outline2Curve = True
    Exit Function
    
fout_Outline2Curve:
    Outline2Curve = False
End Function

Function MaxEclipseCurveU(BessElmt As tBessElmt, AuxElmt As tAuxElmt, DiffBess As tDiffBess, PredData As tPredData, _
                          pPsi As Double, ByRef MaxEclCurve As tMaxEclCurve) As Boolean

Dim ltQ As Double, lL1 As Double, lL12 As Double, lL2 As Double, lC As Double, l1Q As Double, l2Q As Double, l1sQ As Double, l1cQ As Double, l2sQ As Double, l2cQ As Double
Dim l1xcQ_ysQ As Double, l2xcQ_ysQ As Double, l1sg_Q As Double, l2sg_Q As Double, l1g_Q As Double, l2g_Q As Double, l1g As Double, l2g As Double
Dim l1sg As Double, l2sg As Double, l1cg As Double, l2cg As Double, l1Xi As Double, l2Xi As Double, l1Nu As Double, l2Nu     As Double
Dim l1x_Xi As Double, l2x_Xi As Double, l1y_Nu As Double, l2y_Nu As Double, l1Delta2 As Double, l2Delta2 As Double, _
    l1cPhisd As Double, l2cPhisd As Double, l1cPhicd As Double, l2cPhicd As Double, l1tDel As Double, l2tDel As Double, l1Del As Double, l2del As Double, _
    l1Lambda As Double, l2Lambda As Double, l1sPhi As Double, l2sPhi As Double, l1Phi As Double, l2Phi As Double, _
    l1Delta As Double, l1L1_Delta As Double, l1L1pL2 As Double, l1sW As Double, l1cW As Double, l1L1cW As Double, l1M1 As Double, _
    l1Xi1 As Double, l1Nu1 As Double, l1n2 As Double, l1n As Double, l1s  As Double
Dim l2Delta As Double, l2L1_Delta As Double, l2L1pL2 As Double, l2sW As Double, l2cW As Double, l2L1cW As Double, l2M1 As Double, _
    l2Xi1 As Double, l2Nu1 As Double, l2n2 As Double, l2n As Double, l2s  As Double

On Error GoTo fout_MaxEclipseCurveU
' WITH BessElmt, AuxElmt, DiffBess, PredData, MaxEclCurve DO
     ltQ = DiffBess.b1 / (DiffBess.c11 - pPsi * DiffBess.mu1 * BessElmt.cD)
     lL1 = BessElmt.l1 - pPsi * BessElmt.tF1
     lL12 = lL1 * lL1
     lL2 = BessElmt.l2 - pPsi * BessElmt.tF2
     lC = Sqr(1 - pPsi * pPsi)

     l1Q = Atn(ltQ): l2Q = Pi + l1Q
     l1sQ = Sin(l1Q): l2sQ = Sin(l2Q)
     l1cQ = Cos(l1Q): l2cQ = Cos(l2Q)

     l1xcQ_ysQ = BessElmt.x * l1cQ - BessElmt.Y * l1sQ: l2xcQ_ysQ = BessElmt.x * l2cQ - BessElmt.Y * l2sQ
     l1sg_Q = l1xcQ_ysQ / lC: l2sg_Q = l2xcQ_ysQ / lC
     l1g_Q = asin(l1sg_Q): l2g_Q = asin(l2sg_Q)
     l1g = l1g_Q + l1Q: l2g = l2g_Q + l2Q
     l1sg = Sin(l1g): l2sg = Sin(l2g)
     l1cg = Cos(l1g): l2cg = Cos(l2g)

     l1Xi = lC * l1sg: l2Xi = lC * l2sg
     l1Nu = lC * l1cg: l2Nu = lC * l2cg

     l1x_Xi = BessElmt.x - l1Xi: l2x_Xi = BessElmt.x - l2Xi
     l1y_Nu = BessElmt.Y - l1Nu: l2y_Nu = BessElmt.Y - l2Nu
     l1Delta2 = l1x_Xi * l1x_Xi + l1y_Nu * l1y_Nu
     l2Delta2 = l2x_Xi * l2x_Xi + l2y_Nu * l2y_Nu
     If l1Delta2 <= lL12 Then
         l1cPhisd = l1Xi
         l1cPhicd = -l1Nu * BessElmt.SD + pPsi * BessElmt.cD
         l1tDel = l1cPhisd / l1cPhicd
         l1Del = atan2(l1cPhisd, l1cPhicd)
         l1Lambda = BessElmt.mu - l1Del
         l1sPhi = l1Nu * BessElmt.cD + pPsi * BessElmt.SD
         l1Phi = asin(l1sPhi)
         MaxEclCurve.pos1.lng = l1Lambda
         MaxEclCurve.pos1.nb = l1Phi

         l1Delta = Sqr(l1Delta2)
         l1L1_Delta = lL1 - l1Delta
         l1L1pL2 = lL1 + lL2
         l1sW = l1Delta / lL1
         l1cW = Cos(asin(l1sW))
         l1L1cW = lL1 * l1cW
         l1M1 = (lL1 - l1Delta) / (lL1 + lL2)
         l1Xi1 = DiffBess.mu1 * Cos(l1Phi) * Cos(l1Del)
         l1Nu1 = DiffBess.mu1 * l1Xi * BessElmt.SD
         l1n2 = (DiffBess.x1 - l1Xi1) * (DiffBess.x1 - l1Xi1) + (DiffBess.y1 - l1Nu1) * (DiffBess.y1 - l1Nu1)
         l1n = Sqr(l1n2)
         l1s = l1L1cW / l1n
         MaxEclCurve.M1 = l1M1
         MaxEclCurve.s1 = l1s
         MaxEclipseCurveU = True
     Else
         MaxEclipseCurveU = False
     End If
     Exit Function
     
fout_MaxEclipseCurveU:
    MaxEclipseCurveU = False
End Function

Function MaxEclipseCurveP(BessElmt As tBessElmt, AuxElmt As tAuxElmt, DiffBess As tDiffBess, PredData As tPredData, _
                          pPsi As Double, ByRef MaxEclCurve As tMaxEclCurve) As Boolean

Dim ltQ As Double, lL1 As Double, lL12 As Double, lL2 As Double, lC As Double, l1Q As Double, l2Q As Double, l1sQ As Double, l1cQ As Double, l2sQ As Double, l2cQ As Double
Dim l1xcQ_ysQ As Double, l2xcQ_ysQ As Double, l1sg_Q As Double, l2sg_Q As Double, l1g_Q As Double, l2g_Q As Double, l1g As Double, l2g As Double
Dim l1sg As Double, l2sg As Double, l1cg As Double, l2cg As Double, l1Xi As Double, l2Xi As Double, l1Nu As Double, l2Nu     As Double
Dim l1x_Xi As Double, l2x_Xi As Double, l1y_Nu As Double, l2y_Nu As Double, l1Delta2 As Double, l2Delta2 As Double, _
    l1cPhisd As Double, l2cPhisd As Double, l1cPhicd As Double, l2cPhicd As Double, l1tDel As Double, l2tDel As Double, l1Del As Double, l2del As Double, _
    l1Lambda As Double, l2Lambda As Double, l1sPhi As Double, l2sPhi As Double, l1Phi As Double, l2Phi As Double, _
    l1Delta As Double, l1L1_Delta As Double, l1L1pL2 As Double, l1sW As Double, l1cW As Double, l1L1cW As Double, l1M1 As Double, _
    l1Xi1 As Double, l1Nu1 As Double, l1n2 As Double, l1n As Double, l1s  As Double
Dim l2Delta As Double, l2L1_Delta As Double, l2L1pL2 As Double, l2sW As Double, l2cW As Double, l2L1cW As Double, l2M1 As Double, _
    l2Xi1 As Double, l2Nu1 As Double, l2n2 As Double, l2n As Double, l2s  As Double

On Error GoTo fout_MaxEclipseCurveP
' WITH BessElmt, AuxElmt, DiffBess, PredData, MaxEclCurve DO
     ltQ = DiffBess.b1 / (DiffBess.c11 - pPsi * DiffBess.mu1 * BessElmt.cD)
     lL1 = BessElmt.l1 - pPsi * BessElmt.tF1
     lL12 = lL1 * lL1
     lL2 = BessElmt.l2 - pPsi * BessElmt.tF2
     lC = Sqr(1 - pPsi * pPsi)

     l1Q = Atn(ltQ): l2Q = Pi + l1Q
     l1sQ = Sin(l1Q): l2sQ = Sin(l2Q)
     l1cQ = Cos(l1Q): l2cQ = Cos(l2Q)

     l1xcQ_ysQ = BessElmt.x * l1cQ - BessElmt.Y * l1sQ: l2xcQ_ysQ = BessElmt.x * l2cQ - BessElmt.Y * l2sQ
     l1sg_Q = l1xcQ_ysQ / lC: l2sg_Q = l2xcQ_ysQ / lC
     l1g_Q = asin(l1sg_Q): l2g_Q = asin(l2sg_Q)
     l1g = l1g_Q + l1Q: l2g = l2g_Q + l2Q
     l1sg = Sin(l1g): l2sg = Sin(l2g)
     l1cg = Cos(l1g): l2cg = Cos(l2g)

     l1Xi = lC * l1sg: l2Xi = lC * l2sg
     l1Nu = lC * l1cg: l2Nu = lC * l2cg

     l1x_Xi = BessElmt.x - l1Xi: l2x_Xi = BessElmt.x - l2Xi
     l1y_Nu = BessElmt.Y - l1Nu: l2y_Nu = BessElmt.Y - l2Nu
     l1Delta2 = l1x_Xi * l1x_Xi + l1y_Nu * l1y_Nu
     l2Delta2 = l2x_Xi * l2x_Xi + l2y_Nu * l2y_Nu

     If l2Delta2 <= lL12 Then
         l2cPhisd = l2Xi
         l2cPhicd = -l2Nu * BessElmt.SD + pPsi * BessElmt.cD
         l2tDel = l2cPhisd / l2cPhicd
         l2del = atan2(l2cPhisd, l2cPhicd)
         l2Lambda = BessElmt.mu - l2del
         l2sPhi = l2Nu * BessElmt.cD + pPsi * BessElmt.SD
         l2Phi = asin(l2sPhi)
         MaxEclCurve.pos2.lng = l2Lambda
         MaxEclCurve.pos2.nb = l2Phi

         l2Delta = Sqr(l2Delta2)
         l2L1_Delta = lL1 - l2Delta
         l2L1pL2 = lL1 + lL2
         l2sW = l1Delta / lL1
         l2cW = Cos(asin(l2sW))
         l2L1cW = lL1 * l2cW
         l2M1 = (lL1 - l2Delta) / (lL1 + lL2)
         l2Xi1 = DiffBess.mu1 * Cos(l2Phi) * Cos(l2del)
         l2Nu1 = DiffBess.mu1 * l2Xi * BessElmt.SD
         l2n2 = (DiffBess.x1 - l2Xi1) * (DiffBess.x1 - l2Xi1) + (DiffBess.y1 - l2Nu1) * (DiffBess.y1 - l2Nu1)
         l2n = Sqr(l2n2)
         l2s = l2L1cW / l2n
         MaxEclCurve.m2 = l2M1
         MaxEclCurve.s2 = l2s
         MaxEclipseCurveP = True
     Else
         MaxEclipseCurveP = False
     End If
     Exit Function
     
fout_MaxEclipseCurveP:
     MaxEclipseCurveP = False
End Function

Function MaxEclipseCurve(BessElmt As tBessElmt, AuxElmt As tAuxElmt, DiffBess As tDiffBess, PredData As tPredData, _
                          pPsi As Double, ByRef MaxEclCurve As tMaxEclCurve) As Long
MaxEclipseCurve = 0
If MaxEclipseCurveU(BessElmt, AuxElmt, DiffBess, PredData, pPsi, MaxEclCurve) Then
    MaxEclipseCurve = MaxEclipseCurve + 1
End If
If MaxEclipseCurveP(BessElmt, AuxElmt, DiffBess, PredData, pPsi, MaxEclCurve) Then
    MaxEclipseCurve = MaxEclipseCurve + 2
End If
End Function
Function RiseCurve(BessElmt As tBessElmt, AuxElmt As tAuxElmt, DiffBess As tDiffBess, PredData As tPredData, _
                          ByRef RiseSet As tRiseSetCurve) As Boolean


Dim lm  As Double, Tm As Double, l2m As Double, lcg_M As Double, hM As Double, l1g As Double, l2g As Double, l1Xi As Double, l1Nu As Double, l2Xi As Double, l2Nu As Double, _
    l1g_M As Double, l2g_M  As Double
Dim l1cPhisd As Double, l2cPhisd As Double, l1cPhicd As Double, l2cPhicd As Double, l1tDel As Double, l2tDel As Double, l1Del As Double, l2del As Double, _
    l1Lambda As Double, l2Lambda As Double, l1sPhi As Double, l2sPhi As Double, l1Phi As Double, l2Phi  As Double

On Error GoTo fout_RiseCurve
' WITH BessElmt, AuxElmt, DiffBess, PredData, RiseSet DO
     Tm = BessElmt.x / BessElmt.Y
     hM = atan2(BessElmt.x, BessElmt.Y)
     lm = BessElmt.x / Sin(hM)
     lcg_M = (lm * lm + 1 - BessElmt.l1 * BessElmt.l1) / (2 * lm)
     l1g_M = acos(lcg_M): l2g_M = -l1g_M
     l1g = l1g_M + hM: l2g = l2g_M + hM
     l1Xi = Sin(l1g): l2Xi = Sin(l2g)
     l1Nu = Cos(l1g): l2Nu = Cos(l2g)

     l1cPhisd = l1Xi
     l1cPhicd = -l1Nu * BessElmt.SD
     l1tDel = l1cPhisd / l1cPhicd
     l1Del = atan2(l1cPhisd, l1cPhicd)
     l1Lambda = BessElmt.mu - l1Del
     l1sPhi = l1Nu * BessElmt.cD
     l1Phi = asin(l1sPhi)
     RiseSet.pos1.lng = l1Lambda
     RiseSet.pos1.nb = l1Phi
     RiseCurve = True
     Exit Function

fout_RiseCurve:
    RiseCurve = False
End Function


Function SetCurve(BessElmt As tBessElmt, AuxElmt As tAuxElmt, DiffBess As tDiffBess, PredData As tPredData, _
                          ByRef RiseSet As tRiseSetCurve) As Boolean

Dim lm  As Double, Tm As Double, l2m As Double, lcg_M As Double, hM As Double, l1g As Double, l2g As Double, l1Xi As Double, l1Nu As Double, l2Xi As Double, l2Nu As Double, _
    l1g_M As Double, l2g_M  As Double
Dim l1cPhisd As Double, l2cPhisd As Double, l1cPhicd As Double, l2cPhicd As Double, l1tDel As Double, l2tDel As Double, l1Del As Double, l2del As Double, _
    l1Lambda As Double, l2Lambda As Double, l1sPhi As Double, l2sPhi As Double, l1Phi As Double, l2Phi  As Double

On Error GoTo fout_SetCurve
' WITH BessElmt, AuxElmt, DiffBess, PredData, RiseSet DO
     Tm = BessElmt.x / BessElmt.Y
     hM = atan2(BessElmt.x, BessElmt.Y)
     lm = BessElmt.x / Sin(hM)
     lcg_M = (lm * lm + 1 - BessElmt.l1 * BessElmt.l1) / (2 * lm)
     l1g_M = acos(lcg_M): l2g_M = -l1g_M
     l1g = l1g_M + hM: l2g = l2g_M + hM
     l1Xi = Sin(l1g): l2Xi = Sin(l2g)
     l1Nu = Cos(l1g): l2Nu = Cos(l2g)

     l2cPhisd = l2Xi
     l2cPhicd = -l2Nu * BessElmt.SD
     l2tDel = l2cPhisd / l2cPhicd
     l2del = atan2(l2cPhisd, l2cPhicd)
     l2Lambda = BessElmt.mu - l2del
     l2sPhi = l2Nu * BessElmt.cD
     l2Phi = asin(l2sPhi)
     RiseSet.pos2.lng = l2Lambda
     RiseSet.pos2.nb = l2Phi
     SetCurve = True
     Exit Function
     
fout_SetCurve:
    SetCurve = False
End Function

Function RiseSetCurve(BessElmt As tBessElmt, AuxElmt As tAuxElmt, DiffBess As tDiffBess, PredData As tPredData, _
                          ByRef RiseSet As tRiseSetCurve) As Long
RiseSetCurve = 0
If RiseCurve(BessElmt, AuxElmt, DiffBess, PredData, RiseSet) Then
    RiseSetCurve = 1
End If
If SetCurve(BessElmt, AuxElmt, DiffBess, PredData, RiseSet) Then
    RiseSetCurve = RiseSetCurve + 2
End If
End Function

Function RMaxCurve(BessElmt As tBessElmt, AuxElmt As tAuxElmt, DiffBess As tDiffBess, PredData As tPredData, _
                     ByRef RSMax As tRSMaxCurve) As Boolean

Dim lm  As Double, Tm As Double, l2m As Double, lcg_M As Double, hM As Double, l1g As Double, l2g As Double, l1Xi As Double, l1Nu As Double, l2Xi As Double, l2Nu As Double, _
    l1g_M As Double, l2g_M As Double, ltQ As Double, l1Q As Double, l2Q As Double, l1sQ As Double, l2sQ As Double, l1cQ As Double, l2cQ As Double, l1sg_Q As Double, l2sg_Q As Double
Dim l1cPhisd As Double, l2cPhisd As Double, l1cPhicd As Double, l2cPhicd As Double, l1tDel As Double, l2tDel As Double, l1Del As Double, l2del As Double, _
    l1Lambda As Double, l2Lambda As Double, l1sPhi As Double, l2sPhi As Double, l1Phi As Double, l2Phi  As Double
Dim l1sg As Double, l2sg As Double, l1cg As Double, l2cg     As Double
Dim l1x_Xi As Double, l2x_Xi As Double, l1y_Nu As Double, l2y_Nu As Double, l1Delta2 As Double, l2Delta2     As Double
Dim l1g_Q As Double, l2g_Q As Double, lL1 As Double, lL12      As Double

On Error GoTo fout_RMaxCurve
' WITH BessElmt, AuxElmt, DiffBess, PredData, RSMax DO
     lL1 = BessElmt.l1
     lL12 = lL1 * lL1
     ltQ = DiffBess.b1 / DiffBess.c11
     'l1Q = Atn(ltQ)
     l1Q = modpi2(Atn(ltQ))
     If l1Q > Pi Then l1Q = l1Q - Pi

     l1sQ = Sin(l1Q)
     l1cQ = Cos(l1Q)
     l1sg_Q = BessElmt.x * l1cQ - BessElmt.Y * l1sQ
     l1g_Q = asin(l1sg_Q)
     l1g = l1g_Q + l1Q
     l1sg = Sin(l1g)
     l1cg = Cos(l1g)

     l1Xi = l1sg
     l1Nu = l1cg

     l1x_Xi = BessElmt.x - l1Xi
     l1y_Nu = BessElmt.Y - l1Nu
     l1Delta2 = l1x_Xi * l1x_Xi + l1y_Nu * l1y_Nu
     
     If l1Delta2 <= lL12 Then
         l1cPhisd = l1Xi
         l1cPhicd = -l1Nu * BessElmt.SD
         l1tDel = l1cPhisd / l1cPhicd
         l1Del = atan2(l1cPhisd, l1cPhicd)
         l1Lambda = BessElmt.mu - l1Del
         l1sPhi = l1Nu * BessElmt.cD
         l1Phi = asin(l1sPhi)
         RSMax.pos1.lng = l1Lambda
         RSMax.pos1.nb = l1Phi
         RMaxCurve = True
     Else
         RMaxCurve = False
     End If
     Exit Function

fout_RMaxCurve:
     RMaxCurve = False
End Function
Function SMaxCurve(BessElmt As tBessElmt, AuxElmt As tAuxElmt, DiffBess As tDiffBess, PredData As tPredData, _
                     ByRef RSMax As tRSMaxCurve) As Boolean

Dim lm  As Double, Tm As Double, l2m As Double, lcg_M As Double, hM As Double, l1g As Double, l2g As Double, l1Xi As Double, l1Nu As Double, l2Xi As Double, l2Nu As Double, _
    l1g_M As Double, l2g_M As Double, ltQ As Double, l1Q As Double, l2Q As Double, l1sQ As Double, l2sQ As Double, l1cQ As Double, l2cQ As Double, l1sg_Q As Double, l2sg_Q As Double
Dim l1cPhisd As Double, l2cPhisd As Double, l1cPhicd As Double, l2cPhicd As Double, l1tDel As Double, l2tDel As Double, l1Del As Double, l2del As Double, _
    l1Lambda As Double, l2Lambda As Double, l1sPhi As Double, l2sPhi As Double, l1Phi As Double, l2Phi  As Double
Dim l1sg As Double, l2sg As Double, l1cg As Double, l2cg     As Double
Dim l1x_Xi As Double, l2x_Xi As Double, l1y_Nu As Double, l2y_Nu As Double, l1Delta2 As Double, l2Delta2     As Double
Dim l1g_Q As Double, l2g_Q As Double, lL1 As Double, lL12      As Double

On Error GoTo fout_SMaxCurve
' WITH BessElmt, AuxElmt, DiffBess, PredData, RSMax DO
     lL1 = BessElmt.l1
     lL12 = lL1 * lL1
     ltQ = DiffBess.b1 / DiffBess.c11
     l2Q = modpi2(Atn(ltQ))
     If l2Q < Pi Then l2Q = l2Q + Pi
     l2sQ = Sin(l2Q)
     l2cQ = Cos(l2Q)
     l2sg_Q = BessElmt.x * l2cQ - BessElmt.Y * l2sQ
     l2g_Q = asin(l2sg_Q)
     l2g = l2g_Q + l2Q
     l2sg = Sin(l2g)
     l2cg = Cos(l2g)

     l2Xi = l2sg
     l2Nu = l2cg

     l2x_Xi = BessElmt.x - l2Xi
     l2y_Nu = BessElmt.Y - l2Nu
     l2Delta2 = l2x_Xi * l2x_Xi + l2y_Nu * l2y_Nu

     If l2Delta2 <= lL12 Then
         l2cPhisd = l2Xi
         l2cPhicd = -l2Nu * BessElmt.SD
         l2tDel = l2cPhisd / l2cPhicd
         l2del = atan2(l2cPhisd, l2cPhicd)
         l2Lambda = BessElmt.mu - l2del
         l2sPhi = l2Nu * BessElmt.cD
         l2Phi = asin(l2sPhi)
         RSMax.pos2.lng = l2Lambda
         RSMax.pos2.nb = l2Phi
         SMaxCurve = True
     Else
         SMaxCurve = False
     End If
     Exit Function
     
fout_SMaxCurve:
     SMaxCurve = False
End Function

Function RSMaxCurve(BessElmt As tBessElmt, AuxElmt As tAuxElmt, DiffBess As tDiffBess, PredData As tPredData, _
                     ByRef RSMax As tRSMaxCurve) As Long
RSMaxCurve = 0
If RMaxCurve(BessElmt, AuxElmt, DiffBess, PredData, RSMax) Then
    RSMaxCurve = 1
End If
If SMaxCurve(BessElmt, AuxElmt, DiffBess, PredData, RSMax) Then
    RSMaxCurve = RSMaxCurve + 2
End If
End Function

Sub PositieZonMaan(T As Double, ByRef RkM As Double, DecM As Double, ParM As Double, RkZ As Double, DecZ As Double, ParZ As Double, RZ As Double, EphTime As Double)
Dim Obl As Double, NutLon As Double, NutObl As Double, deltaT  As Double
Dim SHelio   As TSVECTOR, SEarth As TSVECTOR, SGeo As TSVECTOR, s As TSVECTOR
Dim l As Double, B As Double, dRkM As Double, dDecM As Double
Obl = Obliquity(T)
Call NutationConst(T, NutLon, NutObl)
EphTime = SiderealTime(T) + NutLon * Cos(Obl)

'{ Main Calculations }
'{ Allereerst de Zon-gegevens }
SHelio.l = 0: SHelio.B = 0: SHelio.r = 0
Call PlanetPosHi(0, T, SEarth)
Call HelioToGeo(SHelio, SEarth, SGeo)
Call PlanetPosHi(0, T - SGeo.r * LightTimeConst, SEarth)
Call HelioToGeo(SHelio, SEarth, SGeo)
Call EclToEqu(SGeo.l, SGeo.B, Obl, RkZ, DecZ)
Call ConvertVSOP_FK5(T, SGeo.l, SGeo.B)
Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RkZ, DecZ)

ParZ = SolarParallax / SGeo.r
RZ = SGeo.r

'{ en nu de gegevens van de Maan }

Dim Dist As Double, dkm As Double, illum As Double, phase As Double, diam As Double
Call Lune(TToJD(T), RkM, DecM, Dist, dkm, diam, phase, illum)
Call Lune(TToJD(T - Dist * LightTimeConst), RkM, DecM, Dist, dkm, diam, phase, illum)
RkM = RkM * Pi / 12
DecM = DecM * Pi / 180
'coordinaten zijn voor J2000. Omzetten naar huidige dag, en daarna appearent berekenen
Call PrecessFK5(0, T, RkM, DecM)
Call Nutation(NutLon, NutObl, Obl, RkM, DecM)

ParM = asin(6378.14 / dkm)

'In verband met de verduistering moet het centrum van de Maan gebruikt worden en niet het centrum van de massa
'dit verschilt lichtelijk. Correcties hieronder:
Call EquToEcl(RkM, DecM, Obl, l, B)
dRkM = -Sin(Obl) * Cos(l) / (Cos(DecM) * Cos(DecM)) * -0.3 / 3600 * Pi / 180
dDecM = (Cos(Obl) * Cos(l) * Cos(RkM) + Sin(l) * Sin(RkM)) * -0.3 / 3600 * Pi / 180
' RkM = RkM + dRkM
'DecM = DecM + dDecM
'  SHelio.l = 0
'  SHelio.B = 0
'  SHelio.r = 0

'  Call PlanetPosHi(EPDATE, T, SEarth)
'  Call HelioToGeo(SHelio, SEarth, SGeo)

'  Call EclToEqu(SGeo.l + NutLon, SGeo.B, Obl + NutObl, RkZ, DecZ)
'  Call Aberration(T, Obl, FK5System, RkZ, DecZ)
'{  EclToEqu(SGeo.l, SGeo.b, Obl, RkZ, DecZ)}
'  ParZ = SolarParallax / SGeo.r
'  RZ = SGeo.r

  'Call MoonPos(T, s)
'  ParM = asin(EarthRadius / s.r)

'  Call EclToEqu(s.l + NutLon, s.B, Obl + NutObl, RkM, DecM)
End Sub

Function ExtremesN(ByVal I As Long, ByVal T As Double, BessElmt As tBessElmt, AuxElmt As tAuxElmt, dBess As tDiffBess, PredData As tPredData, _
                   ByRef Extr As tExtremes) As Boolean

Dim EphTime        As Double
Dim RkM As Double, DecM As Double, ParM As Double, RkZ As Double, DecZ As Double, ParZ As Double, RZ     As Double
Dim BessElmt1 As tBessElmt, BessElmt2 As tBessElmt
Dim lx0 As Double, ly0 As Double, lm2 As Double, ly10 As Double, lm12 As Double, lrho2 As Double, lrho As Double, lx1 As Double, ly1     As Double
Dim ln1 As Double, ln12 As Double, ll1pRho As Double, ln1ll1pRho As Double, lx0x1py0y1 As Double, lx0y1mx1y0 As Double, lsX1 As Double, lcX1     As Double
Dim lt As Double, T0 As Double, lXi As Double, lNu As Double, lcPhisd As Double, lcPhicd As Double, ltDel As Double, lDel As Double, lLambda As Double, lsPhi As Double, lPhi     As Double
Dim lsPhi1 As Double, lPhi1     As Double

On Error GoTo fout_ExtremesN
' WITH BessElmt1, AuxElmt, DBess, PredData, Extr DO
  T0 = T
  T = T0
  lt = 1
    Do Until Abs(lt) < 0.0000001
       Call PositieZonMaan(T, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
       Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt1)
       Call PositieZonMaan(T + 1 / 876600, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
       Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt2)
       Call DiffBess(BessElmt1, BessElmt2, dBess)
       Call Aux_elmts(BessElmt1, AuxElmt, dBess)
       lx0 = BessElmt1.x
       ly0 = BessElmt1.Y
       lm2 = lx0 * lx0 + ly0 * ly0
       ly10 = ly0 / AuxElmt.rho1
       lm12 = lx0 * lx0 + ly10 * ly10
       lrho2 = lm2 / lm12
       lrho = Sqr(lrho2)
       lx1 = dBess.x1
       ly1 = dBess.y1
       ln12 = lx1 * lx1 + ly1 * ly1
       ln1 = Sqr(ln12)
       ll1pRho = BessElmt1.l1 + lrho
       ln1ll1pRho = ln1 * ll1pRho
       lx0x1py0y1 = lx0 * lx1 + ly0 * ly1
       lx0y1mx1y0 = lx0 * ly1 - lx1 * ly0
       lsX1 = lx0y1mx1y0 / ln1ll1pRho
       lcX1 = Abs(Cos(asin(lsX1)))
       If I = 1 Then lcX1 = -lcX1
       lt = ll1pRho / ln1 * lcX1 - lx0x1py0y1 / ln12
       T = T + lt / 876600
    Loop
    
    Call PositieZonMaan(T, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt1)
    Call PositieZonMaan(T + 1 / 876600, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt2)
    Call DiffBess(BessElmt1, BessElmt2, dBess)
    Call Aux_elmts(BessElmt1, AuxElmt, dBess)
    lx0 = BessElmt1.x
    ly0 = BessElmt1.Y
    lm2 = lx0 * lx0 + ly0 * ly0
    ly1 = ly0 / AuxElmt.rho1
    lm12 = lx0 * lx0 + ly1 * ly1
    lrho2 = lm2 / lm12
    lrho = Sqr(lrho2)
    lXi = Sin(atan2(lx0, ly0)) * lrho
    lNu = Cos(atan2(lx0, ly0)) * lrho

    lcPhisd = lXi
    lcPhicd = -lNu * BessElmt1.SD
    ltDel = lcPhisd / lcPhicd
    lDel = atan2(lcPhisd, lcPhicd)
    lLambda = lDel - BessElmt1.mu
    lsPhi1 = lNu * BessElmt1.cD
    lPhi1 = asin(lsPhi1)
    lPhi = Atn(tan(lPhi1) / Sqr(1 - e2))
    If I = 1 Then
       Extr.fPenPos.time = T
       Extr.fPenPos.lng = lLambda
       Extr.fPenPos.nb = lPhi
    Else
       Extr.lPenpos.time = T
       Extr.lPenpos.lng = lLambda
       Extr.lPenpos.nb = lPhi
    End If
    ExtremesN = True
    Exit Function
  
fout_ExtremesN:
    ExtremesN = False
End Function
Function ExtremesZ(ByVal I As Long, ByVal T As Double, BessElmt As tBessElmt, AuxElmt As tAuxElmt, dBess As tDiffBess, PredData As tPredData, _
                   ByRef Extr As tExtremes) As Boolean

Dim EphTime        As Double
Dim RkM As Double, DecM As Double, ParM As Double, RkZ As Double, DecZ As Double, ParZ As Double, RZ     As Double
Dim BessElmt1 As tBessElmt, BessElmt2 As tBessElmt
Dim lx0 As Double, ly0 As Double, lm2 As Double, ly10 As Double, lm12 As Double, lrho2 As Double, lrho As Double, lx1 As Double, ly1     As Double
Dim ln1 As Double, ln12 As Double, ll1pRho As Double, ln1ll1pRho As Double, lx0x1py0y1 As Double, lx0y1mx1y0 As Double, lsX1 As Double, lcX1     As Double
Dim lt As Double, T0 As Double, lXi As Double, lNu As Double, lcPhisd As Double, lcPhicd As Double, ltDel As Double, lDel As Double, lLambda As Double, lsPhi As Double, lPhi     As Double
Dim lsPhi1 As Double, lPhi1     As Double

On Error GoTo fout_ExtremesZ
    T0 = T
    T = T0
    lt = 1
    Do Until Abs(lt) < 0.0000001
       Call PositieZonMaan(T, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
       Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt1)
       Call PositieZonMaan(T + 1 / 876600, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
       Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt2)
       Call DiffBess(BessElmt1, BessElmt2, dBess)
       Call Aux_elmts(BessElmt1, AuxElmt, dBess)
       lx0 = BessElmt1.x
       ly0 = BessElmt1.Y
       lm2 = lx0 * lx0 + ly0 * ly0
       ly10 = ly0 / AuxElmt.rho1
       lm12 = lx0 * lx0 + ly10 * ly10
       lrho2 = lm2 / lm12
       lrho = Sqr(lrho2)
       lx1 = dBess.x1
       ly1 = dBess.y1
       ln12 = lx1 * lx1 + ly1 * ly1
       ln1 = Sqr(ln12)
       lx0x1py0y1 = lx0 * lx1 + ly0 * ly1
       lx0y1mx1y0 = lx0 * ly1 - lx1 * ly0
       lsX1 = lx0y1mx1y0 / ln1
       lcX1 = Abs(Cos(asin(lsX1)))
       If I = 1 Then lcX1 = -lcX1
       lt = lcX1 / ln1 - lx0x1py0y1 / ln12
       T = T + lt / 876600
    Loop

    Call PositieZonMaan(T, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt1)
    Call PositieZonMaan(T + 1 / 876600, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt2)
    Call DiffBess(BessElmt1, BessElmt2, dBess)
    Call Aux_elmts(BessElmt1, AuxElmt, dBess)
    lx0 = BessElmt1.x
    ly0 = BessElmt1.Y
    lm2 = lx0 * lx0 + ly0 * ly0
    ly1 = ly0 / AuxElmt.rho1
    lm12 = lx0 * lx0 + ly1 * ly1
    lrho2 = lm2 / lm12
    lrho = Sqr(lrho2)
    lXi = Sin(atan2(lx0, ly0)) * lrho
    lNu = Cos(atan2(lx0, ly0)) * lrho

    lcPhisd = lXi
    lcPhicd = -lNu * BessElmt1.SD
    ltDel = lcPhisd / lcPhicd
    lDel = atan2(lcPhisd, lcPhicd)
    lLambda = BessElmt1.mu - lDel
    lsPhi1 = lNu * BessElmt1.cD
    lPhi1 = asin(lsPhi1)
    lPhi = Atn(tan(lPhi1) / Sqr(1 - e2))

    If I = 1 Then
       Extr.fUmbPos.time = T
       Extr.fUmbPos.lng = lLambda
       Extr.fUmbPos.nb = lPhi
    Else
       Extr.lUmbPos.time = T
       Extr.lUmbPos.lng = lLambda
       Extr.lUmbPos.nb = lPhi
    End If
 
    ExtremesZ = True
    Exit Function
fout_ExtremesZ:
    ExtremesZ = False
 
End Function

Function Extremes(ByVal T As Double, BessElmt As tBessElmt, AuxElmt As tAuxElmt, dBess As tDiffBess, PredData As tPredData, _
                   ByRef Extr As tExtremes) As Long
Extremes = 0
If ExtremesN(1, T, BessElmt, AuxElmt, dBess, PredData, Extr) Then
    Extremes = Extremes + 1
End If
If ExtremesN(2, T, BessElmt, AuxElmt, dBess, PredData, Extr) Then
    Extremes = Extremes + 2
End If
If ExtremesZ(1, T, BessElmt, AuxElmt, dBess, PredData, Extr) Then
    Extremes = Extremes + 4
End If
If ExtremesZ(2, T, BessElmt, AuxElmt, dBess, PredData, Extr) Then
    Extremes = Extremes + 8
End If
End Function

Function CentralEclipseLocalAppNoon(ByVal T As Double, BessElmt As tBessElmt, AuxElmt As tAuxElmt, DiffBess1 As tDiffBess, PredData As tPredData, _
                              ByRef GreatestEclipse As tGreatestEclipse) As Boolean
'                              ByRef CentralEcliseLAN As tCentralEcliseLAN) As Boolean
Dim EphTime        As Double
Dim RkM As Double, DecM As Double, ParM As Double, RkZ As Double, DecZ As Double, ParZ As Double, RZ     As Double
Dim BessElmt1 As tBessElmt, BessElmt2 As tBessElmt, dBess As tDiffBess
Dim lx0 As Double, ly0 As Double, lm2 As Double, ly10 As Double, lm12 As Double, lrho2 As Double, lrho As Double, lx1 As Double, ly1     As Double
Dim ln1 As Double, ln12 As Double, ll1pRho As Double, ln1ll1pRho As Double, lx0x1py0y1 As Double, lx0y1mx1y0 As Double, lsX1 As Double, lcX1     As Double
Dim lt As Double, T0 As Double, lXi As Double, lNu As Double, lcPhisd As Double, lcPhicd As Double, ltDel As Double, lDel As Double, lLambda As Double, lsPhi As Double, lPhi     As Double
Dim lm As Double, ld As Double, ll1_d As Double, ll1pl2 As Double
Dim lsPhi1 As Double, lPhi1     As Double

On Error GoTo fout_CentralEclipseLocalAppNoon

    lx0x1py0y1 = BessElmt.x * DiffBess1.x1 + BessElmt.Y * DiffBess1.y1
    ln12 = DiffBess1.x1 * DiffBess1.x1 + DiffBess1.y1 * DiffBess1.y1
    lt = -lx0x1py0y1 / ln12
    T = T + lt / 876600
    GreatestEclipse.T = T
    
    Call PositieZonMaan(T, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt1)
    Call PositieZonMaan(T + 1 / 876600, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt2)
    Call DiffBess(BessElmt1, BessElmt2, dBess)
    Call Aux_elmts(BessElmt1, AuxElmt, dBess)
    
    lx0 = BessElmt1.x
    ly0 = BessElmt1.Y
    lm2 = lx0 * lx0 + ly0 * ly0
    lm = Sqr(lm2)
    ly10 = ly0 / AuxElmt.rho1
    lm12 = lx0 * lx0 + ly10 * ly10
    lrho2 = lm2 / lm12
    lrho = Sqr(lrho2)
    
    ld = lm - lrho
    ll1_d = BessElmt1.l1 - ld
    ll1pl2 = 2 * BessElmt1.l1 - 0.5464
    
    GreatestEclipse.magn = (BessElmt1.l1 - ld) / (ll1pl2)
    
'    lXi = Sin(atan2(lx0, ly0)) * lrho
'    lNu = Cos(atan2(lx0, ly0)) * lrho
    
    lXi = lx0 / Sqr(lm12)
    lNu = ly10 / Sqr(lm12)
    lcPhisd = lXi
    lcPhicd = -lNu * BessElmt1.SD
    ltDel = lcPhicd / lcPhisd
    lDel = atan2(lcPhisd, lcPhicd)
    lLambda = BessElmt1.mu - lDel
    lsPhi1 = lNu * BessElmt1.cD
    lPhi1 = asin(lsPhi1)
    lPhi = Atn(tan(lPhi1) / Sqr(1 - e2))
    GreatestEclipse.pos.lng = lLambda
    GreatestEclipse.pos.nb = lPhi
    
    CentralEclipseLocalAppNoon = True
    Exit Function
    
fout_CentralEclipseLocalAppNoon:
    CentralEclipseLocalAppNoon = False
End Function

Function Greatest_Eclipse(ByVal T As Double, ByRef GreatestEclipse As tGreatestEclipse) As Boolean
Dim EphTime        As Double
Dim RkM As Double, DecM As Double, ParM As Double, RkZ As Double, DecZ As Double, ParZ As Double, RZ     As Double
Dim BessElmt1 As tBessElmt, BessElmt2 As tBessElmt, dBess As tDiffBess, AuxElmt As tAuxElmt
Dim lx0 As Double, ly0 As Double, lm2 As Double, ly10 As Double, lm12 As Double, lrho2 As Double, lrho As Double, lx1 As Double, ly1     As Double
Dim ln1 As Double, ln12 As Double, ll1pRho As Double, ln1ll1pRho As Double, lx0x1py0y1 As Double, lx0y1mx1y0 As Double, lsX1 As Double, lcX1     As Double
Dim lt As Double, T0 As Double, lXi As Double, lNu As Double, lcPhisd As Double, lcPhicd As Double, ltDel As Double, lDel As Double, lLambda As Double, lsPhi As Double, lPhi     As Double
Dim lm As Double, ld As Double, ll1_d As Double, ll1pl2 As Double
Dim lsPhi1 As Double, lPhi1     As Double
Dim ltv As Double
Dim lm1 As Double

On Error GoTo fout_Greatest_Eclipse

lt = 1
Do While Abs(lt) > 0.000001
    Call PositieZonMaan(T, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt1)
    Call PositieZonMaan(T + 1 / 876600, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt2)
    Call DiffBess(BessElmt1, BessElmt2, dBess)
    Call Aux_elmts(BessElmt1, AuxElmt, dBess)
    
    lx0x1py0y1 = BessElmt1.x * dBess.x1 + BessElmt1.Y * dBess.y1
    ln12 = dBess.x1 * dBess.x1 + dBess.y1 * dBess.y1
    lt = -lx0x1py0y1 / ln12
    T = T + lt / 876600
Loop
    
    GreatestEclipse.T = T
    Call PositieZonMaan(T, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt1)
    Call PositieZonMaan(T + 1 / 876600, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt2)
    Call DiffBess(BessElmt1, BessElmt2, dBess)
    Call Aux_elmts(BessElmt1, AuxElmt, dBess)
    
    lx0 = BessElmt1.x
    ly0 = BessElmt1.Y
    lm2 = lx0 * lx0 + ly0 * ly0
    lm = Sqr(lm2)
    ly10 = ly0 / AuxElmt.rho1
    
    lrho = AuxElmt.rho1
    lm12 = lx0 * lx0 + ly10 * ly10
    lm1 = Sqr(lm12)
    
   'ly10 = ly0 / AuxElmt.rho1
    lrho2 = lm2 / lm12
    lrho = Sqr(lrho2)
    
    ld = lm - lrho
    ll1_d = BessElmt1.l1 - ld
    ll1pl2 = 2 * BessElmt1.l1 - 0.5464
    
    GreatestEclipse.magn = (BessElmt1.l1 - ld) / (BessElmt1.l1 + BessElmt1.l2)
    
'    lXi = Sin(atan2(lx0, ly0)) * lrho
'    lNu = Cos(atan2(lx0, ly0)) * lrho
    
    lXi = lx0 / lm1
    lNu = ly10 / lm1
    lcPhisd = lXi
    lcPhicd = -lNu * BessElmt1.SD
    ltDel = lcPhicd / lcPhisd
    lDel = atan2(lcPhisd, lcPhicd)
    lLambda = BessElmt1.mu - lDel
    lsPhi1 = lNu * BessElmt1.cD
    lPhi1 = asin(lsPhi1)
    lPhi = Atn(tan(lPhi1) / Sqr(1 - e2))
    GreatestEclipse.pos.lng = lLambda
    GreatestEclipse.pos.nb = lPhi
    
    Greatest_Eclipse = True
    Exit Function
    
fout_Greatest_Eclipse:
    Greatest_Eclipse = False
End Function

Function Local_Eclipse(ByVal T As Double, ByVal Latitude As Double, ByVal Longitude As Double, ByVal Altitude As Double, ByRef localeclipse As tLocalEclipse, BeginEindMax As String) As Boolean

Dim BessElmt As tBessElmt, AuxElmt As tAuxElmt, dBess As tDiffBess, PredData As tPredData, _
    Extr As tExtremes, limits As tLimits, OutCurve As tOutCurve, MaxEclCurve As tMaxEclCurve, _
    RiseSet As tRiseSetCurve, RSMax As tRSMaxCurve, GreatestEclipse As tGreatestEclipse
Dim EphTime        As Double
Dim RkM As Double, DecM As Double, ParM As Double, RkZ As Double, DecZ As Double, ParZ As Double, RZ     As Double
Dim BessElmt1 As tBessElmt, BessElmt2 As tBessElmt
Dim lx0 As Double, ly0 As Double, lm2 As Double, ly10 As Double, lm12 As Double, lrho2 As Double, lrho As Double, lx1 As Double, ly1     As Double
Dim ln1 As Double, ln12 As Double, ll1pRho As Double, ln1ll1pRho As Double, lx0x1py0y1 As Double, lx0y1mx1y0 As Double, lsX1 As Double, lcX1     As Double
Dim lt As Double, T0 As Double, lXi As Double, lNu As Double, lcPhisd As Double, lcPhicd As Double, ltDel As Double, lDel As Double, lLambda As Double, lsPhi As Double, lPhi     As Double
Dim lm As Double, ld As Double, ll1_d As Double, ll1pl2 As Double
Dim lsPhi1 As Double, lPhi1     As Double, lcPhi As Double
Dim sLatitude As String, sLongitude As String
Dim lambda As Double, LambdaEph As Double, Phi As Double, lrhocphi1 As Double
Dim lH As Double, lC As Double, lS As Double, lA As Double, lB As Double, lMu As Double, lDelta As Double
Dim lsDelta As Double, lcDelta As Double, lsd As Double, lcd As Double
Dim lx As Double, lu As Double, ly As Double
Dim lv As Double, lPsi As Double, lL1 As Double, lL2 As Double
Dim lmu1 As Double, ld1 As Double, lXisD As Double
Dim lXi1 As Double, lu1 As Double
Dim lNu1 As Double, lv1 As Double
Dim ln2 As Double, ln As Double
Dim lDELTAG As Double, lsFi As Double, lcFi As Double
Dim lQ As Double
Dim sAltitude As String
Dim lXib As Double, lXie As Double, lCb As Double, lCe As Double
Dim tDat As tDatum
Dim lSign As Long
Dim ltv As Double
Dim I As Long

On Error GoTo fout_Local_Eclipse

    ltv = 9999
    
    Phi = Latitude * Pi / 180
    lambda = Longitude * Pi / 180
    lH = Altitude * 0.000000156785
    
    lsPhi = Sin(Phi): lcPhi = Cos(Phi)
    lC = 1 / Sqr(1 - e2 * lsPhi * lsPhi)
    lS = (1 - e2) * lC
    LambdaEph = lambda + ApproxDeltaT(T) * 1.002738 * 15 / 3600 * Pi / 180
    
    lA = (lS + lH) * lsPhi '
    lrhocphi1 = (lC + lH) * lcPhi
    lt = -1
    
    I = 0
    While Abs(lt) > 0.00001
        I = I + 1
        With localeclipse
            Call PositieZonMaan(T, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
            Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt1)
            Call PositieZonMaan(T + 1 / 876600, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
            Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt2)
            Call DiffBess(BessElmt1, BessElmt2, dBess)
            Call Aux_elmts(BessElmt1, AuxElmt, dBess)
        
            lMu = BessElmt1.mu
            lDelta = lMu - LambdaEph
            lsDelta = Sin(lDelta): lcDelta = Cos(lDelta)
            lB = lrhocphi1 * lcDelta
            lsd = BessElmt1.SD
            lcd = BessElmt1.cD
            lx = BessElmt1.x
            lXi = lrhocphi1 * lsDelta
            lu = lx - lXi
            
            ly = BessElmt1.Y
            lNu = lA * lcd - lB * lsd
            lv = ly - lNu
            
            lPsi = lA * lsd + lB * lcd
            lL1 = BessElmt1.l1 - lPsi * BessElmt1.tF1
            'If BeginEindMax = "T" Then
                lL2 = BessElmt1.l2 - lPsi * BessElmt1.tF1
            'End If
                
            lmu1 = dBess.mu1
            ld1 = dBess.d1
            lXisD = lXi * lsd
            
            lx1 = dBess.x1
            lXi1 = lB * lmu1
            lu1 = lx1 - lXi1
            
            ly1 = dBess.y1
            lNu1 = lmu1 * lXisD - lPhi * ld1
            lv1 = ly1 - lNu1
            
            ln2 = lu1 * lu1 + lv1 * lv1
            ln = Sqr(ln2)
            
            ld = lu * lu1 + lv * lv1
            lDELTAG = (lu * lv1 - lu1 * lv) / ln
            
            If BeginEindMax = "T" Then
                lsFi = lDELTAG / lL2
            Else
                lsFi = lDELTAG / lL1
            End If
            Select Case BeginEindMax
            Case "B"
                If Abs(lsFi) <= 1 Then
                    lcFi = -(Abs(Cos(asin(lsFi))))
                End If
            Case "E"
                If Abs(lsFi) <= 1 Then
                    lcFi = Abs(Cos(asin(lsFi)))
                End If
            Case "M"
                lcFi = 0
            Case "T"
                lcFi = 0
            End Select
            If BeginEindMax = "T" Then
                lt = -ld / ln2 + lL2 * lcFi / ln
            Else
                If Abs(lsFi) <= 1 Then
                    lt = -ld / ln2 + lL1 * lcFi / ln
                Else 'lager nauwkeurigheid maar zonder problemen met bereik sin.
                    lt = sign(lL1) * (lL1 - lu * lu - lv * lv) / (2 * ld)
                End If
            End If
        '    lt = -lD / ln2 + lL1 * lcFi / ln
            If Abs(lt) < Abs(ltv) And sign(lt) = sign(ltv) Then
            ElseIf I > 20 Then GoTo fout_Local_Eclipse
            End If
            ltv = lt
            T = T + lt / 876600
        End With
    Wend
    tDat = JDNaarKalender(TToJD(T))
'    Debug.Print Int(tDat.DD) & "-" & tDat.mm & "-" & tDat.jj & " : ";
'    Debug.Print StrHMS_DMS(Frac(tDat.DD) * 360, 3, 3, False, False, "h", 2)
    Select Case BeginEindMax
        Case "B"
            localeclipse.Tb = T - ApproxDeltaT(T) / 86400 / 36525
            lQ = atan2(lu, lv)
            lC = atan2(lXi, lNu)
            localeclipse.Vb = lQ - lC
        Case "E"
            localeclipse.Te = T - ApproxDeltaT(T) / 86400 / 36525
            lQ = atan2(lu, lv)
            lC = atan2(lXi, lNu)
            localeclipse.Ve = lQ - lC
        Case "M"
            localeclipse.Tm = T - ApproxDeltaT(T) / 86400 / 36525
            lm2 = lu * lu + lv * lv
            lm = Sqr(lm2)
            localeclipse.mag = (lL1 - lm) / (lL1 + lL2) '/(2 * lL1 - 0.5459)
        Case "T"
            lcFi = (Abs(Cos(asin(lsFi)))) * sign(lL2)
            'lSign: indien L2 > 0 dan totale verduistering
            'L2<0 dan ringvormige verduistering
            'bepaling tijd begin: bij L2>0 cFi negatief nemen, L2<0, cFi pos. nemen
            localeclipse.Ttotaalb = T - (lL2 * lcFi / ln) / 876600 - ApproxDeltaT(T) / 86400 / 36525
            localeclipse.Ttotaale = T + (lL2 * lcFi / ln) / 876600 - ApproxDeltaT(T) / 86400 / 36525
            localeclipse.Tm = T - ApproxDeltaT(T) / 86400 / 36525
            localeclipse.MagTotaal = (lL1 - lL2) / (lL1 + lL2)
            'vanwege nauwkeurigheid moet Xi opnieuw berekend worden
            'hiervoor gebruiken we geïnterpoleerde waarden Mu  + mu1
            'Xi = rhocosphi1 * sin(Mu + Mu1*dt - LambdaEhp)
            lXib = Sin(BessElmt1.mu + (-(lL2 * lcFi / ln) / 24) * dBess.mu1 - LambdaEph) * lrhocphi1
            lXie = Sin(BessElmt1.mu + ((lL2 * lcFi / ln) / 24) * dBess.mu1 - LambdaEph) * lrhocphi1
            lQ = atan2(lu1, lv1) ' +/- asin(lsFi)
            lCb = atan2(lXib, lNu)
            lCe = atan2(lXie, lNu)
            localeclipse.Vtotaalb = lQ + asin(lsFi) - lCb
            localeclipse.Vtotaale = lQ + Pi - asin(lsFi) - lCe
'            tDat = JDNaarKalender(TToJD(LocalEclipse.Ttotaalb))
'            Debug.Print "------------------"
'            Debug.Print Int(tDat.DD) & "-" & tDat.mm & "-" & tDat.jj & " : ";
'            Debug.Print StrHMS_DMS(Frac(tDat.DD) * 360, 3, 3, False, False, "h", 2)
'            tDat = JDNaarKalender(TToJD(LocalEclipse.Ttotaale))
'            Debug.Print Int(tDat.DD) & "-" & tDat.mm & "-" & tDat.jj & " : ";
'            Debug.Print StrHMS_DMS(Frac(tDat.DD) * 360, 3, 3, False, False, "h", 2)
'            Debug.Print "------------------"
    End Select
        
    
    Local_Eclipse = True
    Exit Function
    
fout_Local_Eclipse:
    Local_Eclipse = False
End Function

Sub TestEclipse(T As Double)
Dim BessElmt As tBessElmt, AuxElmt As tAuxElmt, dBess As tDiffBess, PredData As tPredData, _
    Extr As tExtremes, limits As tLimits, OutCurve As tOutCurve, MaxEclCurve As tMaxEclCurve, _
    RiseSet As tRiseSetCurve, RSMax As tRSMaxCurve, GreatestEclipse As tGreatestEclipse, _
    localeclipse As tLocalEclipse

Dim EphTime        As Double
Dim RkM As Double, DecM As Double, ParM As Double, RkZ As Double, DecZ As Double, ParZ As Double, RZ     As Double
Dim BessElmt1 As tBessElmt, BessElmt2 As tBessElmt
Dim lx0 As Double, ly0 As Double, lm2 As Double, ly10 As Double, lm12 As Double, lrho2 As Double, lrho As Double, lx1 As Double, ly1     As Double
Dim ln1 As Double, ln12 As Double, ll1pRho As Double, ln1ll1pRho As Double, lx0x1py0y1 As Double, lx0y1mx1y0 As Double, lsX1 As Double, lcX1     As Double
Dim lt As Double, T0 As Double, lXi As Double, lNu As Double, lcPhisd As Double, lcPhicd As Double, ltDel As Double, lDel As Double, lLambda As Double, lsPhi As Double, lPhi     As Double
Dim lsPhi1 As Double, lPhi1     As Double
Dim I As Long, pPsi As Double
Dim nRes As Boolean
Dim nRes2 As Long
Dim sLatitude As String, sLongitude As String, sAltitude As String

T0 = T - 0.5 / 876600

'Debug.Print "Predicted data, point on central line, duration of eclipse"
While T0 < T + 0.5 / 876600
    Call PositieZonMaan(T0 + 1 / 876600, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt2)
    Call PositieZonMaan(T0, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt1)
    Call DiffBess(BessElmt1, BessElmt2, dBess)
    Call Aux_elmts(BessElmt1, AuxElmt, dBess)
    nRes = PredDataSolarEcl(BessElmt1, AuxElmt, dBess, PredData)
    If nRes = True Then
        Debug.Print T0 & vbTab & StrHMS_DMS(PredData.lambda * 180 / Pi, 7, 1, False, True, "g", 3) _
        & vbTab & StrHMS_DMS(PredData.Phi * 180 / Pi, 7, 1, True, False, "g", 3) _
        & vbTab & StrHMS_DMS(2 * PredData.s * 15, 4, 1, True, False, "h", 5)
    End If
    T0 = T0 + 0.05 / 876600
Wend
'--------------------------------------------------------------------------------
T0 = T - 0.5 / 876600

Debug.Print "Northern and southern limits of the penumbra"
While T0 < T + 0.5 / 876600
    Call PositieZonMaan(T0, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt1)
    Call PositieZonMaan(T0 + 1 / 876600, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt2)
    Call DiffBess(BessElmt1, BessElmt2, dBess)
    Call Aux_elmts(BessElmt1, AuxElmt, dBess)
    nRes2 = LimitsUmbraPenumbra(BessElmt1, AuxElmt, dBess, PredData, limits)
    If nRes2 <> 0 Then Debug.Print T0;
    If nRes2 And 1 Then
        Debug.Print vbTab & "N: " & StrHMS_DMS(limits.ULimN.lng * 180 / Pi, 7, 1, False, True, "g", 3) _
        & vbTab & StrHMS_DMS(limits.ULimN.nb * 180 / Pi, 7, 1, True, False, "g", 3);
    End If
    If nRes2 And 2 Then
        Debug.Print vbTab & "Z: " & StrHMS_DMS(limits.ULimZ.lng * 180 / Pi, 7, 1, False, True, "g", 3) _
        & vbTab & StrHMS_DMS(limits.ULimZ.nb * 180 / Pi, 7, 1, True, False, "g", 3);
    End If
    If nRes2 And 4 Then
        Debug.Print vbTab & "N: " & StrHMS_DMS(limits.PLimN.lng * 180 / Pi, 7, 1, False, True, "g", 3) _
        & vbTab & StrHMS_DMS(limits.ULimN.nb * 180 / Pi, 7, 1, True, False, "g", 3);
    End If
    If nRes2 And 8 Then
        Debug.Print vbTab & "Z: " & StrHMS_DMS(limits.PLimZ.lng * 180 / Pi, 7, 1, False, True, "g", 3) _
        & vbTab & StrHMS_DMS(limits.PLimZ.nb * 180 / Pi, 7, 1, True, False, "g", 3);
    End If
    If nRes2 <> 0 Then Debug.Print
    T0 = T0 + 0.05 / 876600
Wend

'--------------------------------------------------------------------------------
T0 = T - 0.5 / 876600
Debug.Print "Outline curves of an eclipse"
While T0 < T + 0.5 / 876600
    Call PositieZonMaan(T0, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt1)
    Call PositieZonMaan(T0 + 1 / 876600, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt2)
    Call DiffBess(BessElmt1, BessElmt2, dBess)
    Call Aux_elmts(BessElmt1, AuxElmt, dBess)
    nRes = Outline1Curve(BessElmt1, AuxElmt, dBess, PredData, OutCurve)
    If nRes Then
        Debug.Print T0;
        Debug.Print vbTab & "N: " & StrHMS_DMS(OutCurve.bQ * 180 / Pi, 7, 1, False, False, "g", 3) _
        & vbTab & StrHMS_DMS(OutCurve.eQ * 180 / Pi, 7, 1, False, False, "g", 3);
        nRes = Outline2Curve(BessElmt1, AuxElmt, dBess, PredData, 20 * Pi / 180, OutCurve)
        If nRes Then
            Debug.Print vbTab & "N: " & StrHMS_DMS(OutCurve.pos.lng * 180 / Pi, 7, 1, False, True, "g", 3) _
            & vbTab & StrHMS_DMS(OutCurve.pos.nb * 180 / Pi, 7, 1, True, False, "g", 3)
        Else
            Debug.Print
        End If
    End If
    T0 = T0 + 0.05 / 876600
Wend
'--------------------------------------------------------------------------------

T0 = T - 0.5 / 876600

T0 = T
Debug.Print "Curves of max and middle eclipse, semi-dur and equal magn."
While T0 < T + 0.5 / 876600
    Call PositieZonMaan(T0, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt1)
    Call PositieZonMaan(T0 + 1 / 876600, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt2)
    Call DiffBess(BessElmt1, BessElmt2, dBess)
    Call Aux_elmts(BessElmt1, AuxElmt, dBess)
    pPsi = 0.2: nRes2 = MaxEclipseCurve(BessElmt1, AuxElmt, dBess, PredData, pPsi, MaxEclCurve)
    If nRes2 <> 0 Then
        Debug.Print T0;
        If nRes2 And 1 Then
            Debug.Print vbTab & "MN: " & StrHMS_DMS(MaxEclCurve.pos1.lng * 180 / Pi, 7, 1, False, False, "g", 3) _
            & vbTab & StrHMS_DMS(MaxEclCurve.pos1.nb * 180 / Pi, 7, 1, False, False, "g", 3);
            Debug.Print vbTab & StrHMS_DMS(MaxEclCurve.s1 * 15, 2, 1, False, False, "h", 4);
            Debug.Print vbTab; Format(MaxEclCurve.M1, "0.000");
        End If
        If nRes2 And 2 Then
            Debug.Print vbTab & "MZ: " & StrHMS_DMS(MaxEclCurve.pos2.lng * 180 / Pi, 7, 1, False, False, "g", 3) _
            & vbTab & StrHMS_DMS(MaxEclCurve.pos2.nb * 180 / Pi, 7, 1, False, False, "g", 3);
            Debug.Print vbTab & StrHMS_DMS(MaxEclCurve.s2 * 15, 2, 1, False, False, "h", 4);
            Debug.Print vbTab; Format(MaxEclCurve.M1, "0.000");
        End If
        Debug.Print
    End If
    T0 = T0 + 0.05 / 876600
Wend

'--------------------------------------------------------------------------------
T0 = T - 0.5 / 876600

Debug.Print "Points on the rising and setting curves"
While T0 < T + 0.5 / 876600
    Call PositieZonMaan(T0, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt1)
    Call PositieZonMaan(T0 + 1 / 876600, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt2)
    Call DiffBess(BessElmt1, BessElmt2, dBess)
    Call Aux_elmts(BessElmt1, AuxElmt, dBess)
    nRes2 = RiseSetCurve(BessElmt1, AuxElmt, dBess, PredData, RiseSet)
    If nRes2 <> 0 Then
        Debug.Print T0;
        If nRes2 And 1 Then
            Debug.Print vbTab & "R: " & StrHMS_DMS(RiseSet.pos1.lng * 180 / Pi, 7, 1, False, False, "g", 3) _
            & vbTab & StrHMS_DMS(RiseSet.pos1.nb * 180 / Pi, 7, 1, False, False, "g", 3);
        End If
        If nRes2 And 2 Then
            Debug.Print vbTab & "S: " & StrHMS_DMS(RiseSet.pos2.lng * 180 / Pi, 7, 1, False, True, "g", 3) _
            & vbTab & StrHMS_DMS(RiseSet.pos2.nb * 180 / Pi, 7, 1, True, False, "g", 3);
        End If
        Debug.Print
    End If
    T0 = T0 + 0.05 / 876600
Wend

'--------------------------------------------------------------------------------

T0 = T - 0.5 / 876600

Debug.Print "Points on the curve of maximum eclipse at sunrise and sunset"
While T0 < T + 0.5 / 876600
    Call PositieZonMaan(T0, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt1)
    Call PositieZonMaan(T0 + 1 / 876600, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
    Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt2)
    Call DiffBess(BessElmt1, BessElmt2, dBess)
    Call Aux_elmts(BessElmt1, AuxElmt, dBess)
    nRes2 = RSMaxCurve(BessElmt1, AuxElmt, dBess, PredData, RSMax)
    If nRes2 <> 0 Then
        Debug.Print T0;
        If nRes2 And 1 Then
            Debug.Print vbTab & "RM: " & StrHMS_DMS(RSMax.pos1.lng * 180 / Pi, 7, 1, False, False, "g", 3) _
            & vbTab & StrHMS_DMS(RSMax.pos1.nb * 180 / Pi, 7, 1, False, False, "g", 3);
        End If
        If nRes2 And 2 Then
            Debug.Print vbTab & "SM: " & StrHMS_DMS(RSMax.pos2.lng * 180 / Pi, 7, 1, False, True, "g", 3) _
            & vbTab & StrHMS_DMS(RSMax.pos2.nb * 180 / Pi, 7, 1, True, False, "g", 3);
        End If
        Debug.Print
    End If
    T0 = T0 + 0.05 / 876600
Wend
'--------------------------------------------------------------------------------

T0 = T
Debug.Print "Time and position of first/last contact of the umbra/penumbra"

Call PositieZonMaan(T0, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt1)
Call PositieZonMaan(T0 + 1 / 876600, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt2)
Call DiffBess(BessElmt1, BessElmt2, dBess)
Call Aux_elmts(BessElmt1, AuxElmt, dBess)
nRes2 = Extremes(T0, BessElmt1, AuxElmt, dBess, PredData, Extr)
If nRes2 <> 0 Then
    Debug.Print T0;
    If nRes2 And 1 Then
        Debug.Print vbTab & "FP: " & StrHMS_DMS(Extr.fPenPos.lng * 180 / Pi, 7, 1, False, False, "g", 3) _
        & vbTab & StrHMS_DMS(Extr.fPenPos.nb * 180 / Pi, 7, 1, False, False, "g", 3);
        Debug.Print vbTab; StrHMS_DMS(Frac(JDNaarKalender(TToJD(Extr.fPenPos.time)).DD) * 360, 3, 3, False, False, "h", 2);
    End If
    If nRes2 And 2 Then
        Debug.Print vbTab & "LP: " & StrHMS_DMS(Extr.lPenpos.lng * 180 / Pi, 7, 1, False, False, "g", 3) _
        & vbTab & StrHMS_DMS(Extr.lPenpos.nb * 180 / Pi, 7, 1, False, False, "g", 3);
        Debug.Print vbTab; StrHMS_DMS(Frac(JDNaarKalender(TToJD(Extr.lPenpos.time)).DD) * 360, 3, 3, False, False, "h", 2);
    End If
    If nRes2 And 4 Then
        Debug.Print vbTab & "FU: " & StrHMS_DMS(Extr.fUmbPos.lng * 180 / Pi, 7, 1, False, False, "g", 3) _
        & vbTab & StrHMS_DMS(Extr.fUmbPos.nb * 180 / Pi, 7, 1, False, False, "g", 3);
        Debug.Print vbTab; StrHMS_DMS(Frac(JDNaarKalender(TToJD(Extr.fUmbPos.time)).DD) * 360, 3, 3, False, False, "h", 2);
    End If
    If nRes2 And 8 Then
        Debug.Print vbTab & "LU: " & StrHMS_DMS(Extr.lUmbPos.lng * 180 / Pi, 7, 1, False, False, "g", 3) _
        & vbTab & StrHMS_DMS(Extr.lUmbPos.nb * 180 / Pi, 7, 1, False, False, "g", 3);
        Debug.Print vbTab; StrHMS_DMS(Frac(JDNaarKalender(TToJD(Extr.lUmbPos.time)).DD) * 360, 3, 3, False, False, "h", 2);
    End If
    Debug.Print
End If

'T0 = T
'Debug.Print "Time, position and magnitude of greatest eclipse"

'Call PositieZonMaan(T0, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
'Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt1)
'Call PositieZonMaan(T0 + 1 / 876600, RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime)
'Call Bess_elmts(RkM, DecM, ParM, RkZ, DecZ, ParZ, RZ, EphTime, BessElmt2)
'Call DiffBess(BessElmt1, BessElmt2, dBess)
'Call Aux_elmts(BessElmt1, AuxElmt, dBess)
'If Greatest_Eclipse(T0, BessElmt1, AuxElmt, dBess, PredData, GreatestEclipse) Then
'    Debug.Print T0;
'    Debug.Print "GE : " & StrHMS_DMS(GreatestEclipse.pos.lng * 180 / PI, 7, 1, False, False, "g", 3) _
'            & vbTab & StrHMS_DMS(GreatestEclipse.pos.nb * 180 / PI, 7, 1, False, False, "g", 3) _
'            & vbTab & Format(GreatestEclipse.magn, "0.000")
'End If


'============================
'BEPALEN GLOBALE OMSTANDIGHEDEN WERKT NU CORRECT
'NU NOG BEPALEN VAN PLAATSELIJKE OMSTANDIGHEDEN (VANAF PAGINA 38, NOG DOEN: TIJD/PLAATS MAXIMALE ECLIPS
    Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Latitude", _
            sLatitude)
    Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Longitude", _
            sLongitude)
    Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Altitude", _
            sAltitude)
    T0 = T
Call Local_Eclipse(T, Val(sLatitude), Val(sLongitude), Val(sAltitude), localeclipse, "B")
Call Local_Eclipse(T, Val(sLatitude), Val(sLongitude), Val(sAltitude), localeclipse, "E")
Call Local_Eclipse(T, Val(sLatitude), Val(sLongitude), Val(sAltitude), localeclipse, "M")
Call Local_Eclipse(T, Val(sLatitude), Val(sLongitude), Val(sAltitude), localeclipse, "T")


End Sub
