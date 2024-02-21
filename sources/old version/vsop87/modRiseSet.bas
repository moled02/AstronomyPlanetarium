Attribute VB_Name = "modRiseSet"
'(*****************************************************************************)
'(* Name:    RiseSet                                                          *)
'(* Type:    Procedure                                                        *)
'(* Purpose: Calculate times of rise, transit and set                         *)
'(* Arguments:                                                                *)
'(*   T : number of Julian centuries since J2000.0 (0h LOCAL TIME)            *)
'(*   DeltaT : the difference DT - UT in seconds                              *)
'(*   RA1, Decl1 : coordinates of object at time T - 1 day                    *)
'(*   RA2, Decl2 : coordinates of object at time T                            *)
'(*   RA3, Decl3 : coordinates of object at time T + 1 day                    *)
'(*   Height0 : height of object above the horizon                            *)
'(*   Rise, Transit, Setting : times (in radians) of rise, transit and set    *)
'(*   Flags : flags special situations                                        *)
'(*     ALWAYS_ABOVE : object is always above horizon                         *)
'(*     ALWAYS_BELOW : object is always below horizon                         *)
'(*                                                                           *)
'(* Note: The special situation handling is still very limited and simple.    *)
'(*       No provision was made for a body reaching the minimum or maximum    *)
'(*       altitude, NutLon - Pi/2, and NutLon,  during the two days' span of  *)
'(*       the RA and Decl values.  If a body is out of bounds on the middle   *)
'(*       date, no further checks are performed.                              *)
'(*****************************************************************************)

Sub RiseSet(ByVal T As Double, ByVal DeltaT As Double, ByVal RA1 As Double, ByVal Decl1 As Double, ByVal RA2 As Double, ByVal Decl2 As Double, ByVal RA3 As Double, ByVal Decl3 As Double, ByVal Height0 As Double, ByVal Lon As Double, ByVal Lat As Double, ByRef RTS As tRiseSetTran)


Dim M(2) As Double
Dim H As Double, cosH As Double, sin_Height As Double, Height As Double, Max_Height As Double, n As Double
Dim GST As Double, GST0 As Double, dmn As Double, dmo As Double
Dim RA    As Double, Decl As Double
Dim i As Long, Dummy As Long
Dim fout As Boolean

'{ Make sure the RAs don't jump from 24h to 0h }
RA1 = RA2 + modpi(RA1 - RA2)
RA3 = RA2 + modpi(RA3 - RA2)
RTS.flags = 0
GST0 = SiderealTime(T) '{ Greenwich sidereal time at 0h LT }
cosH = (Sin(Height0) - Sin(Lat) * Sin(Decl2)) / (Cos(Lat) * Cos(Decl2))
M(0) = (RA2 + Lon - GST0) / Pi2
If M(0) < 0 Then M(0) = Frac(M(0)) + 1
If M(0) > 1 Then M(0) = Frac(M(0))
If (Abs(cosH) <= 1) Then
  H = acos(cosH)
  M(1) = M(0) - H / Pi2
  M(2) = M(0) + H / Pi2
Else
  If cosH < -1 Then
    M(1) = 0
    M(2) = 1
    RTS.flags = ALWAYS_ABOVE
  Else
    M(1) = 1
    M(2) = 0
    RTS.flags = ALWAYS_BELOW
  End If
End If
If (M(1) < 0) Then M(1) = Frac(M(1)) + 1
If (M(2) > 1) Then M(2) = Frac(M(2))
If RTS.flags = 0 Then
  Dummy = 2
Else
  Dummy = 0
End If
For i = 0 To Dummy
  dmo = 2
  dmn = 99999
  Do Until (Abs(dmn) < 0.00002) Or (fout)
    fout = False
    GST = GST0 + (360.985647 * DToR) * M(i)
    n = M(i) + DeltaT / 86400
    RA = Interpol3(RA1, RA2, RA3, n)
    If (i <> 0) Then
      Decl = Interpol3(Decl1, Decl2, Decl3, n)
    End If
    H = modpi(GST - Lon - RA)
    If (i = 0) Then
      dmn = -H / Pi2
    Else
      sin_Height = Sin(Lat) * Sin(Decl) + Cos(Lat) * Cos(Decl) * Cos(H)
      If sin_Height >= 1 Then
        Height = PI * 0.5
      ElseIf sin_Height <= -1 Then
        Height = -PI * 0.5
      Else
        Height = asin(sin_Height)
      End If
      dmn = Cos(Height) * (Height - Height0) / (Pi2 * Cos(Decl) * Cos(Lat) * Sin(H))
    End If
    fout = Abs(dmn) > (Abs(dmo) + 0.05)
    M(i) = M(i) + dmn
    dmo = dmn
    If (i = 0) Then
        Max_Height = Height
    End If
  Loop
  If fout Then
    M(i) = -1
  End If
Next

For i = 0 To 2
  If (M(i) < 0) Or (M(i) >= 1) Then
     M(i) = -1
  End If
Next

RTS.Rise = M(1) * Pi2
RTS.Transit = M(0) * Pi2
RTS.Setting = M(2) * Pi2
End Sub

'(*****************************************************************************)
'(* Name:    MoonSetHeight                                                    *)
'(* Type:    Function                                                         *)
'(* Purpose: Calculate the true height of the Moon at time of rise or set.    *)
'(* Arguments:                                                                *)
'(*   Parallax : horizontal parallax of the Moon                              *)
'(* Return value:                                                             *)
'(*   the true height above the horizon corresponding to Moonrise or Moonset  *)
'(*****************************************************************************)

Function MoonSetHeight(Parallax As Double) As Double
MoonSetHeight = 0.7275 * Parallax + h0Planet
End Function

Function Frac(x As Double) As Double
Frac = x - Fix(x)
End Function
