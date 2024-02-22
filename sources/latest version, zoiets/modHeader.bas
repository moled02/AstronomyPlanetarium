Attribute VB_Name = "modHeader"
Public Sub EARTH_LBR_FOR(ByVal T As Double, ByRef tsRes As TSVECTOR)
' Compute heliocentric, spherical coordinates, LBR
' for Earth.

Dim L, B, r
  T = T / 10

' Compute heliocentric, ecliptical Longitude L in radians
  L = Earth_L01(T) + Earth_L02(T) _
    + Earth_L11(T) + Earth_L21(T) + Earth_L31(T) _
    + Earth_L41(T) + Earth_L51(T)
    
' Modulate L value between 0 and 2*Pi
  If Abs(L) > (2 * Pi) Then L = L - 2 * Pi * Int(L / 2 / Pi)

' Compute heliocentric, ecliptical Latitude B in radians
  B = Earth_B01(T) _
    + Earth_B11(T) + Earth_B21(T) + Earth_B31(T) _
    + Earth_B41(T)
      
' Compute heliocentric distance R in AU
  r = Earth_R01(T) + Earth_R02(T) _
    + Earth_R11(T) + Earth_R21(T) + Earth_R31(T) _
    + Earth_R41(T) + Earth_R51(T)

  tsRes.L = L
  tsRes.B = B
  tsRes.r = r
  '= "L = " & l & "; B = " & b & "; R = " & r
End Sub

Public Function MERCURY_LBR_FOR(ByVal T As Double, tsRes As TSVECTOR)
' Compute heliocentric, spherical coordinates, LBR
' for MERCURY.

Dim L, B, r
  T = T / 10

' Compute heliocentric, ecliptical Longitude L in radians
  L = Mercury_L01(T) + Mercury_L02(T) + Mercury_L03(T) + Mercury_L11(T) _
    + Mercury_L12(T) + Mercury_L21(T) + Mercury_L31(T) + Mercury_L41(T) _
    + Mercury_L51(T)

' Modulate L value between 0 and 2*Pi
  If Abs(L) > (2 * Pi) Then L = L - 2 * Pi * Int(L / 2 / Pi)

  B = Mercury_B01(T) + Mercury_B02(T) + Mercury_B11(T) + Mercury_B21(T) _
    + Mercury_B31(T) + Mercury_B41(T) + Mercury_B51(T)
    
  r = Mercury_R01(T) + Mercury_R02(T) + Mercury_R03(T) + Mercury_R11(T) _
    + Mercury_R12(T) + Mercury_R21(T) + Mercury_R31(T) + Mercury_R41(T) _
    + Mercury_R51(T)
  

  tsRes.L = L
  tsRes.B = B
  tsRes.r = r


End Function

Public Function VENUS_LBR_FOR(ByVal T As Double, tsRes As TSVECTOR)
' Compute heliocentric, spherical coordinates, LBR
' for Venus.

Dim L, B, r
  T = T / 10

' Compute heliocentric, ecliptical Longitude L in radians
  L = Venus_L01(T) + Venus_L11(T) + Venus_L21(T) + Venus_L31(T) _
    + Venus_L41(T) + Venus_L51(T)

' Modulate L value between 0 and 2*Pi
  If Abs(L) > (2 * Pi) Then L = L - 2 * Pi * Int(L / 2 / Pi)

  B = Venus_B01(T) + Venus_B11(T) + Venus_B21(T) + Venus_B31(T) _
    + Venus_B41(T) + Venus_B51(T)
    
  r = Venus_R01(T) + Venus_R11(T) + Venus_R21(T) + Venus_R31(T) _
    + Venus_R41(T) + Venus_R51(T)
  

  tsRes.L = L
  tsRes.B = B
  tsRes.r = r
  'VENUS_LBR_FOR = "L = " & l & "; B = " & b & "; R = " & r

End Function


Public Function MARS_LBR_FOR(ByVal T As Double, tsRes As TSVECTOR)
' Compute heliocentric, spherical coordinates, LBR
' for Mars.

Dim L, B, r
  T = T / 10

' Compute heliocentric longitude L in radians
  L = Mars_L01(T) + Mars_L02(T) + Mars_L03(T) _
    + Mars_L11(T) + Mars_L12(T) _
    + Mars_L21(T) + Mars_L31(T) + Mars_L41(T) + Mars_L51(T)
    
' Modulate L value between 0 and 2*Pi
  If Abs(L) > (2 * Pi) Then L = L - 2 * Pi * Int(L / 2 / Pi)
    
' Compute heliocentric latitude B in radians
  B = Mars_B01(T) + Mars_B11(T) + Mars_B21(T) + Mars_B31(T) _
    + Mars_B41(T) + Mars_B51(T)
  
' Compute heliocentric distance R in AU
  r = Mars_R01(T) + Mars_R02(T) + Mars_R03(T) + Mars_R11(T) _
    + Mars_R12(T) _
    + Mars_R21(T) + Mars_R31(T) + Mars_R41(T) + Mars_R51(T)
  

  tsRes.L = L
  tsRes.B = B
  tsRes.r = r
  'MARS_LBR_FOR = "L = " & l & "; B = " & b & "; R = " & r

End Function


Public Function Jupiter_LBR_FOR(ByVal T As Double, tsRes As TSVECTOR)
' Compute heliocentric, spherical coordinates, LBR
' for Jupiter.

Dim L, B, r
  T = T / 10

' Compute heliocentric, ecliptical Longitude L in radians
  L = Jupiter_L01(T) + Jupiter_L02(T)
  L = L + Jupiter_L11(T) + Jupiter_L21(T) + Jupiter_L31(T) _
    + Jupiter_L41(T) + Jupiter_L51(T)
    
' Modulate L value between 0 and 2*Pi
  If Abs(L) > (2 * Pi) Then L = L - 2 * Pi * Int(L / 2 / Pi)

' Compute heliocentric, ecliptical Latitude B in radians
  B = Jupiter_B01(T) _
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
  tsRes.L = L
  tsRes.B = B
  tsRes.r = r
  'Jupiter_LBR_FOR = "L = " & l & "; B = " & b & "; R = " & r

End Function


Public Function SATURN_LBR_FOR(ByVal T As Double, tsRes As TSVECTOR)
' Compute heliocentric, spherical coordinates, LBR
' for Saturn.

Dim L, B, r
  T = T / 10

' Compute heliocentric, ecliptical Longitude L in radians
  L = Saturn_L01(T) + Saturn_L02(T) + Saturn_L03(T) _
    + Saturn_L11(T) + Saturn_L12(T) _
    + Saturn_L21(T) + Saturn_L31(T) + Saturn_L41(T) _
    + Saturn_L51(T)

' Modulate L value between 0 and 2*Pi
  If Abs(L) > (2 * Pi) Then L = L - 2 * Pi * Int(L / 2 / Pi)

' Compute heliocentric, ecliptical Latitude B in radians
  B = Saturn_B01(T) + Saturn_B02(T) _
    + Saturn_B11(T) + Saturn_B21(T) + Saturn_B31(T) _
    + Saturn_B41(T) + Saturn_B51(T)

' Compute heliocentric distance R in AU
  r = Saturn_R01(T) + Saturn_R02(T) + Saturn_R03(T) _
    + Saturn_R11(T) + Saturn_R12(T) + Saturn_R21(T) _
    + Saturn_R31(T) + Saturn_R41(T) + Saturn_R51(T)
  
' Return LBR values within a labeled and delimited string.
  tsRes.L = L
  tsRes.B = B
  tsRes.r = r
  'SATURN_LBR_FOR = "L = " & l & "; B = " & b & "; R = " & r

End Function


Public Function URANUS_LBR_FOR(ByVal T As Double, tsRes As TSVECTOR)
' Compute heliocentric, spherical coordinates, LBR
' for Uranus.


Dim L, B, r
  T = T / 10

' Compute heliocentric, ecliptical Longitude L in radians
  L = Uranus_L01(T) + Uranus_L02(T) _
    + Uranus_L11(T) + Uranus_L21(T) + Uranus_L31(T) _
    + Uranus_L41(T) + Uranus_L51(T)

' Modulate L value between 0 and 2*Pi
  If Abs(L) > (2 * Pi) Then L = L - 2 * Pi * Int(L / 2 / Pi)

' Compute heliocentric, ecliptical Latitude B in radians
  B = Uranus_B01(T) _
    + Uranus_B11(T) + Uranus_B21(T) + Uranus_B31(T) _
    + Uranus_B41(T)

' Compute heliocentric distance R in AU
  r = Uranus_R01(T) + Uranus_R02(T) + Uranus_R03(T) _
    + Uranus_R11(T) + Uranus_R12(T) + Uranus_R21(T) _
    + Uranus_R31(T) + Uranus_R41(T)
  
' Return LBR values within a labeled and delimited string.
  'URANUS_LBR_FOR = "L = " & l & "; B = " & b & "; R = " & r
  tsRes.L = L
  tsRes.B = B
  tsRes.r = r
End Function


Public Function NEPTUNE_LBR_FOR(ByVal T As Double, tsRes As TSVECTOR)
' Compute heliocentric, spherical coordinates, LBR
' for Neptune.

Dim L, B, r
  T = T / 10

' Compute heliocentric, ecliptical Longitude L in radians
  L = Neptune_L01(T)
  L = L + Neptune_L11(T)
  L = L + Neptune_L21(T)
  L = L + Neptune_L31(T)
  L = L + Neptune_L41(T)
  L = L + Neptune_L51(T)

' Modulate L value between 0 and 2*Pi
  If Abs(L) > (2 * Pi) Then L = L - 2 * Pi * Int(L / 2 / Pi)

' Compute heliocentric, ecliptical Latitude B in radians
  B = Neptune_B01(T)
  B = B + Neptune_B11(T)
  B = B + Neptune_B21(T)
  B = B + Neptune_B31(T)
  B = B + Neptune_B41(T)
  B = B + Neptune_B51(T)

' Compute heliocentric distance R in AU
  r = Neptune_R01(T) + Neptune_R02(T)
  r = r + Neptune_R11(T)
  r = r + Neptune_R21(T)
  r = r + Neptune_R31(T)
  r = r + Neptune_R41(T)
    
' Return LBR values within a labeled and delimited string.
  'NEPTUNE_LBR_FOR = "L = " & l & "; B = " & b & "; R = " & r
  tsRes.L = L
  tsRes.B = B
  tsRes.r = r
End Function




