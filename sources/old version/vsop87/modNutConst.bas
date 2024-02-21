Attribute VB_Name = "modNutConst"
'(*****************************************************************************)
'(* Name:    NutationConst                                                    *)
'(* Type:    Procedure                                                        *)
'(* Purpose: calculate the nutation in longitude and obliquity                *)
'(* Arguments:                                                                *)
'(*   T : number of centuries since J2000                                     *)
'(*   NutLon, NutObl : the requested nutation in longitude and obliquity      *)
'(*****************************************************************************)

Sub NutationConst(T As Double, ByRef NutLon As Double, ByRef NutObl As Double)

Dim i As Long, j As Long, k As Long, l As Long
Dim flag As Boolean
Dim MS As Double, MM As Double, FF As Double, DD As Double, Om As Double
Dim SinVal As Double, CosVal As Double, ArgSin As Double, ArgCos As Double, Tmp As Double
Dim LonArg As Double, OblArg As Double
    
Dim SinTab(5) As TSINCOSTAB, CosTab(5) As TSINCOSTAB

'  {Polynomials page 132}

Dim POm    As Variant 'T3POLY ' = (125.0445222,  -1934.1362608,  0.00207833, 2.22e-6);
Dim PMSun  As Variant 'T3POLY  '= (357.5277233,  35999.0503400, -0.0001603 ,-3.33e-6);
Dim PMMoon As Variant 'T3POLY '= (134.9629814, 477198.8673981,  0.0086972 , 1.778e-5);
Dim PF     As Variant 'T3POLY '= ( 93.2719103, 483202.0175381, -0.00368250, 3.056e-6);
Dim PD     As Variant 'T3POLY '= (297.8503631, 445267.1114800, -0.0019142 , 5.278e-6);

'   { Table 27.A
'  { The coefficients of the angles D, M, M'... are in the same order as }
'  { in the Supplement to the Astronomical Almanac of 1984 (pS23-S25),   }
'  { which is different from table 27.A.  The order here is M, M', F, D, }
'  { Omega.  Arguments are stored in units of 0.0001" and 0.00001"       }
'  { respecitively for the constant and T-coefficients of the nutation.  }
'  { This allows all coefficients to be stored as integers.  The first   }
'  { terms are too large and are therefore calculated seperately.        }
Const NUTTABTERMS = 62
Const NUTTABSIZE = 9

Dim NutTab As Variant '(NUTTABTERMS, NUTTABSIZE) As Long


POm = Array(125.0445222, -1934.1362608, 0.00207833, 0.00000222)
PMSun = Array(357.5277233, 35999.05034, -0.0001603, -0.00000333)
PMMoon = Array(134.9629814, 477198.8673981, 0.0086972, 0.00001778)
PF = Array(93.2719103, 483202.0175381, -0.0036825, 0.000003056)
PD = Array(297.8503631, 445267.11148, -0.0019142, 0.000005278)

NutTab = Array(Array(0, 0, 2, -2, 2, -13187, -16, 5736, -31), _
    Array(0, 0, 2, 0, 2, -2274, -2, 977, -5), Array(0, 0, 0, 0, 2, 2062, 2, -895, 5), Array(0, 1, 0, 0, 0, 1426, -34, 54, -1), Array(1, 0, 0, 0, 0, 712, 1, -7, 0), _
    Array(0, 1, 2, -2, 2, -517, 12, 224, -6), Array(0, 0, 2, 0, 1, -386, -4, 200, 0), Array(1, 0, 2, 0, 2, -301, 0, 129, -1), Array(0, -1, 2, -2, 2, 217, -5, -95, 3), _
    Array(1, 0, 0, -2, 0, -158, 0, 0, 0), Array(0, 0, 2, -2, 1, 129, 1, -70, 0), Array(-1, 0, 2, 0, 2, 123, 0, -53, 0), Array(0, 0, 0, 2, 0, 63, 0, 0, 0), _
    Array(1, 0, 0, 0, 1, 63, 1, -33, 0), Array(-1, 0, 2, 2, 2, -59, 0, 26, 0), Array(-1, 0, 0, 0, 1, -58, -1, 32, 0), Array(1, 0, 2, 0, 1, -51, 0, 27, 0), _
    Array(2, 0, 0, -2, 0, 48, 0, 0, 0), Array(-2, 0, 2, 0, 1, 46, 0, -24, 0), Array(0, 0, 2, 2, 2, -38, 0, 16, 0), Array(2, 0, 2, 0, 2, -31, 0, 13, 0), _
    Array(2, 0, 0, 0, 0, 29, 0, 0, 0), Array(1, 0, 2, -2, 2, 29, 0, -12, 0), Array(0, 0, 2, 0, 0, 26, 0, 0, 0), Array(0, 0, 2, -2, 0, -22, 0, 0, 0), _
    Array(-1, 0, 2, 0, 1, 21, 0, -10, 0), Array(0, 2, 0, 0, 0, 17, -1, 0, 0), Array(-1, 0, 0, 2, 1, 16, 0, -8, 0), Array(0, 2, 2, -2, 2, -16, 1, 7, 0), _
    Array(0, 1, 0, 0, 1, -15, 0, 9, 0), Array(1, 0, 0, -2, 1, -13, 0, 7, 0), Array(0, -1, 0, 0, 1, -12, 0, 6, 0), Array(2, 0, -2, 0, 0, 11, 0, 0, 0), _
    Array(-1, 0, 2, 2, 1, -10, 0, 5, 0), Array(1, 0, 2, 2, 2, -8, 0, 3, 0), Array(0, 1, 2, 0, 2, 7, 0, -3, 0), Array(1, 1, 0, -2, 0, -7, 0, 0, 0), _
    Array(0, -1, 2, 0, 2, -7, 0, 3, 0), Array(0, 0, 2, 2, 1, -7, 0, 3, 0), Array(1, 0, 0, 2, 0, 6, 0, 0, 0), Array(2, 0, 2, -2, 2, 6, 0, -3, 0), _
    Array(1, 0, 2, -2, 1, 6, 0, -3, 0), Array(-2, 0, 0, 2, 1, -6, 0, 3, 0), Array(0, 0, 0, 2, 1, -6, 0, 3, 0), Array(1, -1, 0, 0, 0, 5, 0, 0, 0), _
    Array(0, -1, 2, -2, 1, -5, 0, 3, 0), Array(0, 0, 0, -2, 1, -5, 0, 3, 0), Array(2, 0, 2, 0, 1, -5, 0, 3, 0), Array(-2, 0, 2, 0, 2, -3, 0, 0, 0), _
    Array(2, 0, 0, -2, 1, 4, 0, 0, 0), Array(0, 1, 2, -2, 1, 4, 0, 0, 0), Array(1, -1, 2, 0, 2, -3, 0, 0, 0), Array(-1, -1, 2, 2, 2, -3, 0, 0, 0), _
    Array(3, 0, 2, 0, 2, -3, 0, 0, 0), Array(0, -1, 2, 2, 2, -3, 0, 0, 0), Array(1, -1, 0, -1, 0, -3, 0, 0, 0), Array(1, 0, 0, -1, 0, -4, 0, 0, 0), _
    Array(0, 1, 0, -2, 0, -4, 0, 0, 0), Array(1, 0, -2, 0, 0, 4, 0, 0, 0), Array(0, 0, 0, 1, 0, -4, 0, 0, 0), Array(1, 1, 0, 0, 0, -3, 0, 0, 0), _
    Array(1, 0, 2, 0, 0, 3, 0, 0, 0))

MM = Eval3Poly(C3P(PMMoon), T) * DToR
MS = Eval3Poly(C3P(PMSun), T) * DToR
FF = Eval3Poly(C3P(PF), T) * DToR
DD = Eval3Poly(C3P(PD), T) * DToR
Om = Eval3Poly(C3P(POm), T) * DToR

Call CalcSinCosTab(MM, 3, SinTab(0), CosTab(0))
Call CalcSinCosTab(MS, 2, SinTab(1), CosTab(1))
Call CalcSinCosTab(FF, 4, SinTab(2), CosTab(2))
Call CalcSinCosTab(DD, 4, SinTab(3), CosTab(3))
Call CalcSinCosTab(Om, 2, SinTab(4), CosTab(4))

'{ the first terms are too big for the table : }
NutLon = (-0.01742 * T - 17.1996) * SinTab(4).W(1)      '    { sin(OM) }
NutObl = (0.00089 * T + 9.2025) * CosTab(4).W(1) '    { cos(OM) }

For i = 0 To NUTTABTERMS - 1
    flag = True    '{First non-zero coefficient of one of the five anlges}
    For j = 0 To 4
        k = NutTab(i)(j)
        If (k <> 0) Then
            If (k < 0) Then l = -k Else l = k
            SinVal = SinTab(j).W(l)
            If (k < 0) Then SinVal = -SinVal
            CosVal = CosTab(j).W(l)
            If flag Then
                ArgSin = SinVal
                ArgCos = CosVal
                flag = False
            Else
                Tmp = ArgSin * CosVal + ArgCos * SinVal
                ArgCos = ArgCos * CosVal - ArgSin * SinVal
                ArgSin = Tmp
            End If
        End If
    Next
    OblArg = 0#
    LonArg = NutTab(i)(5) * 0.0001  '        {constant coefficient of sine}
    k = NutTab(i)(6)  '      {T-coefficient of sine}
    If (k <> 0) Then LonArg = LonArg + 0.00001 * T * k
    k = NutTab(i)(7)  '      {constant coefficient of cosine}
    If (k <> 0) Then
        OblArg = 0.0001 * k
        k = NutTab(i)(8)  '      {T-coefficient of cosine}
        If (k <> 0) Then OblArg = OblArg + 0.00001 * T * k
    End If
    NutLon = NutLon + LonArg * ArgSin
    NutObl = NutObl + OblArg * ArgCos
Next
NutLon = NutLon * SToR
NutObl = NutObl * SToR
End Sub


