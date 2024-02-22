Attribute VB_Name = "modSidTime"
'(*****************************************************************************)
'(* Name:    SiderealTime                                                     *)
'(* Type:    Function                                                         *)
'(* Purpose: calculate sidereal time at Greenwich                             *)
'(* Arguments:                                                                *)
'(*   T : number of Julian centuries since J2000                              *)
'(* Return value:                                                             *)
'(*   Greenwich sidereal time                                                 *)
'(*****************************************************************************)

Function SiderealTime(T As Double) As Double

Dim Theta As Double
Theta = T * (360.98564736629 * 36525 + T * (0.000387933 - T / 38710000))
SiderealTime = modpi2((280.46061837 + Theta) * DToR)
End Function

'(*****************************************************************************)
'(* Name:   SiderealTime0                                                     *)
'(* Type:   Function                                                          *)
'(* Purpose: calculate sidereal time at Greenwich at 0h UT                    *)
'(* Arguments:                                                                *)
'(*   T : number of Julian centuries since J2000                              *)
'(* Return value:                                                             *)
'(*   Greenwich sidereal time                                                 *)
'(*****************************************************************************)

Function SiderealTime0(T As Double) As Double

Dim Theta As Double
T = T - Frac(T * 36525 + 0.5) * 36525 '{T at 0h of the day}
Theta = T * (36000.770053608 + T * (0.000387933 - T / 38710000))
SiderealTime0 = modpi2((100.46061837 + Theta) * DToR)
End Function
Function PlaatselijkeSterrentijd(dat As tDatum) As Double
' bereken juliaansedatum
Dim A As Integer, B As Integer, M As Integer, j As Integer
Dim D As Double, JD As Double, T As Double
Dim Theta As Double
Dim SiderealTime As Double
Dim nLongitude As Double
Dim JD_ZT As Double, JD_WT As Double
Dim dattijd As Date
Dim txtLongitude As String

JD = KalenderNaarJD(dat)
'correctie voor zomertijd/wintertijd
Call Zomertijd_Wintertijd(dat.jj, JD_ZT, JD_WT)
If JD >= JD_ZT And JD < JD_WT Then
    JD = JD - 2 / 24
Else
    JD = JD - 1 / 24
End If

'Bereken T (juliaanse eeuwen)
T = (JD - 2451545) / 36525
'bereken sid.time Greenwich
Theta = T * (360.98564736629 * 36525 + T * (0.000387933 - T / 38710000))
SiderealTime = modpi2((280.46061837 + Theta) * Pi / 180) * 12 / Pi
'Corrigeer voor de oosterlengte. Dit moet er afgetrokken worden
'wat inhoudt dat er ongeveer 20 min. opgeteld worden, want het is -5.08 oosterlengte
Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Longitude", _
            txtLongitude)
SiderealTime = SiderealTime - 4 * ReadDMS(txtLongitude) / 60
'en op scherm plaatsen
'stBar.Panels(2).Text = Format(SiderealTime / 24, "hh:mm:ss")
PlaatselijkeSterrentijd = SiderealTime / 24
End Function
'(*****************************************************************************)
'(* Name:    modpi2                                                           *)
'(* Type:    Function                                                         *)
'(* Purpose: reduce an angle to the interval (0, 2pi).                        *)
'(*****************************************************************************)

Function ReadDMS(s As String) As Double
Dim D As Long, M As Long, sign As Long
Dim angle As Double, ss   As Double
    angle = Val(s)
    If (angle < 0) Then
        sign = -1
        angle = -angle
    Else
        sign = 1
    End If
    D = Int(angle)
    angle = (angle - D) * 100
    M = Int(angle)
    ss = (angle - M) * 100 + 0.00001
               '{ Otherwise we might get 59.999... seconds }
    ReadDMS = sign * (D + M / 60# + ss / 3600#)
End Function


