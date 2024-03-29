Attribute VB_Name = "Module2"
Public Sub EARTH_LBR_FOR(T As Double, ByRef tsRes As TSVECTOR)
' Compute heliocentric, spherical coordinates, LBR
' for Earth.

Dim l, b, r

' Compute heliocentric, ecliptical Longitude L in radians
  l = Earth_L01(T) + Earth_L02(T) _
    + Earth_L11(T) + Earth_L21(T) + Earth_L31(T) _
    + Earth_L41(T) + Earth_L51(T)
    
' Modulate L value between 0 and 2*Pi
  If Abs(l) > (2 * PI) Then l = l - 2 * PI * Int(l / 2 / PI)

' Compute heliocentric, ecliptical Latitude B in radians
  b = Earth_B01(T) _
    + Earth_B11(T) + Earth_B21(T) + Earth_B31(T) _
    + Earth_B41(T)
      
' Compute heliocentric distance R in AU
  r = Earth_R01(T) + Earth_R02(T) _
    + Earth_R11(T) + Earth_R21(T) + Earth_R31(T) _
    + Earth_R41(T) + Earth_R51(T)

  tsRes.l = l
  tsRes.b = b
  tsRes.r = r
  '= "L = " & l & "; B = " & b & "; R = " & r
End Sub

Public Function MERCURY_LBR_FOR(T As Double, tsRes As TSVECTOR)
' Compute heliocentric, spherical coordinates, LBR
' for MERCURY.

Dim l, b, r

' Compute heliocentric, ecliptical Longitude L in radians
  l = Mercury_L01(T) + Mercury_L02(T) + Mercury_L03(T) + Mercury_L11(T) _
    + Mercury_L12(T) + Mercury_L21(T) + Mercury_L31(T) + Mercury_L41(T) _
    + Mercury_L51(T)

' Modulate L value between 0 and 2*Pi
  If Abs(l) > (2 * PI) Then l = l - 2 * PI * Int(l / 2 / PI)

  b = Mercury_B01(T) + Mercury_B02(T) + Mercury_B11(T) + Mercury_B21(T) _
    + Mercury_B31(T) + Mercury_B41(T) + Mercury_B51(T)
    
  r = Mercury_R01(T) + Mercury_R02(T) + Mercury_R03(T) + Mercury_R11(T) _
    + Mercury_R12(T) + Mercury_R21(T) + Mercury_R31(T) + Mercury_R41(T) _
    + Mercury_R51(T)
  

  tsRes.l = l
  tsRes.b = b
  tsRes.r = r


End Function

Public Function VENUS_LBR_FOR(T As Double, tsRes As TSVECTOR)
' Compute heliocentric, spherical coordinates, LBR
' for Venus.

Dim l, b, r

' Compute heliocentric, ecliptical Longitude L in radians
  l = Venus_L01(T) + Venus_L11(T) + Venus_L21(T) + Venus_L31(T) _
    + Venus_L41(T) + Venus_L51(T)

' Modulate L value between 0 and 2*Pi
  If Abs(l) > (2 * PI) Then l = l - 2 * PI * Int(l / 2 / PI)

  b = Venus_B01(T) + Venus_B11(T) + Venus_B21(T) + Venus_B31(T) _
    + Venus_B41(T) + Venus_B51(T)
    
  r = Venus_R01(T) + Venus_R11(T) + Venus_R21(T) + Venus_R31(T) _
    + Venus_R41(T) + Venus_R51(T)
  

  tsRes.l = l
  tsRes.b = b
  tsRes.r = r
  'VENUS_LBR_FOR = "L = " & l & "; B = " & b & "; R = " & r

End Function


Public Function MARS_LBR_FOR(T As Double, tsRes As TSVECTOR)
' Compute heliocentric, spherical coordinates, LBR
' for Mars.

Dim l, b, r

' Compute heliocentric longitude L in radians
  l = Mars_L01(T) + Mars_L02(T) + Mars_L03(T) _
    + Mars_L11(T) + Mars_L12(T) _
    + Mars_L21(T) + Mars_L31(T) + Mars_L41(T) + Mars_L51(T)
    
' Modulate L value between 0 and 2*Pi
  If Abs(l) > (2 * PI) Then l = l - 2 * PI * Int(l / 2 / PI)
    
' Compute heliocentric latitude B in radians
  b = Mars_B01(T) + Mars_B11(T) + Mars_B21(T) + Mars_B31(T) _
    + Mars_B41(T) + Mars_B51(T)
  
' Compute heliocentric distance R in AU
  r = Mars_R01(T) + Mars_R02(T) + Mars_R03(T) + Mars_R11(T) _
    + Mars_R12(T) _
    + Mars_R21(T) + Mars_R31(T) + Mars_R41(T) + Mars_R51(T)
  

  tsRes.l = l
  tsRes.b = b
  tsRes.r = r
  'MARS_LBR_FOR = "L = " & l & "; B = " & b & "; R = " & r

End Function


Public Function Jupiter_LBR_FOR(T As Double, tsRes As TSVECTOR)
' Compute heliocentric, spherical coordinates, LBR
' for Jupiter.

Dim l, b, r

' Compute heliocentric, ecliptical Longitude L in radians
  l = Jupiter_L01(T) + Jupiter_L02(T)
  l = l + Jupiter_L11(T) + Jupiter_L21(T) + Jupiter_L31(T) _
    + Jupiter_L41(T) + Jupiter_L51(T)
    
' Modulate L value between 0 and 2*Pi
  If Abs(l) > (2 * PI) Then l = l - 2 * PI * Int(l / 2 / PI)

' Compute heliocentric, ecliptical Latitude B in radians
  b = Jupiter_B01(T) _
    + Jupiter_B11(T) + Jupiter_B21(T) + Jupiter_B31(T) _
    + Jupiter_B41(T) + Jupiter_B51(T)

' Compute heliocentric distance R in AU
  r = Jupiter_R01(T) + Jupiter_R02(T)
  r = r + Jupiter_R11(T)
  r = r + Jupiter_R21(T)
  r = r + Jupiter_R31(T)
  r = r + Jupiter_R41(T)
  r = r + Jupiter_R51(T)

  
' Return LBR values within a labeled and delimited string.
  tsRes.l = l
  tsRes.b = b
  tsRes.r = r
  'Jupiter_LBR_FOR = "L = " & l & "; B = " & b & "; R = " & r

End Function


Public Function SATURN_LBR_FOR(T As Double, tsRes As TSVECTOR)
' Compute heliocentric, spherical coordinates, LBR
' for Saturn.

Dim l, b, r

' Compute heliocentric, ecliptical Longitude L in radians
  l = Saturn_L01(T) + Saturn_L02(T) + Saturn_L03(T) _
    + Saturn_L11(T) + Saturn_L12(T) _
    + Saturn_L21(T) + Saturn_L31(T) + Saturn_L41(T) _
    + Saturn_L51(T)

' Modulate L value between 0 and 2*Pi
  If Abs(l) > (2 * PI) Then l = l - 2 * PI * Int(l / 2 / PI)

' Compute heliocentric, ecliptical Latitude B in radians
  b = Saturn_B01(T) + Saturn_B02(T) _
    + Saturn_B11(T) + Saturn_B21(T) + Saturn_B31(T) _
    + Saturn_B41(T) + Saturn_B51(T)

' Compute heliocentric distance R in AU
  r = Saturn_R01(T) + Saturn_R02(T) + Saturn_R03(T) _
    + Saturn_R11(T) + Saturn_R12(T) + Saturn_R21(T) _
    + Saturn_R31(T) + Saturn_R41(T) + Saturn_R51(T)
  
' Return LBR values within a labeled and delimited string.
  tsRes.l = l
  tsRes.b = b
  tsRes.r = r
  'SATURN_LBR_FOR = "L = " & l & "; B = " & b & "; R = " & r

End Function


Public Function URANUS_LBR_FOR(T As Double, tsRes As TSVECTOR)
' Compute heliocentric, spherical coordinates, LBR
' for Uranus.


Dim l, b, r

' Compute heliocentric, ecliptical Longitude L in radians
  l = Uranus_L01(T) + Uranus_L02(T) _
    + Uranus_L11(T) + Uranus_L21(T) + Uranus_L31(T) _
    + Uranus_L41(T) + Uranus_L51(T)

' Modulate L value between 0 and 2*Pi
  If Abs(l) > (2 * PI) Then l = l - 2 * PI * Int(l / 2 / PI)

' Compute heliocentric, ecliptical Latitude B in radians
  b = Uranus_B01(T) _
    + Uranus_B11(T) + Uranus_B21(T) + Uranus_B31(T) _
    + Uranus_B41(T)

' Compute heliocentric distance R in AU
  r = Uranus_R01(T) + Uranus_R02(T) + Uranus_R03(T) _
    + Uranus_R11(T) + Uranus_R12(T) + Uranus_R21(T) _
    + Uranus_R31(T) + Uranus_R41(T)
  
' Return LBR values within a labeled and delimited string.
  'URANUS_LBR_FOR = "L = " & l & "; B = " & b & "; R = " & r
  tsRes.l = l
  tsRes.b = b
  tsRes.r = r
End Function


Public Function NEPTUNE_LBR_FOR(T As Double, tsRes As TSVECTOR)
' Compute heliocentric, spherical coordinates, LBR
' for Neptune.

Dim l, b, r

' Compute heliocentric, ecliptical Longitude L in radians
  l = Neptune_L01(T)
  l = l + Neptune_L11(T)
  l = l + Neptune_L21(T)
  l = l + Neptune_L31(T)
  l = l + Neptune_L41(T)
  l = l + Neptune_L51(T)

' Modulate L value between 0 and 2*Pi
  If Abs(l) > (2 * PI) Then l = l - 2 * PI * Int(l / 2 / PI)

' Compute heliocentric, ecliptical Latitude B in radians
  b = Neptune_B01(T)
  b = b + Neptune_B11(T)
  b = b + Neptune_B21(T)
  b = b + Neptune_B31(T)
  b = b + Neptune_B41(T)
  b = b + Neptune_B51(T)

' Compute heliocentric distance R in AU
  r = Neptune_R01(T) + Neptune_R02(T)
  r = r + Neptune_R11(T)
  r = r + Neptune_R21(T)
  r = r + Neptune_R31(T)
  r = r + Neptune_R41(T)
    
' Return LBR values within a labeled and delimited string.
  'NEPTUNE_LBR_FOR = "L = " & l & "; B = " & b & "; R = " & r
  tsRes.l = l
  tsRes.b = b
  tsRes.r = r
End Function




