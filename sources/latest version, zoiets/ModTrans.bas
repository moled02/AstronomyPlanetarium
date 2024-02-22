Attribute VB_Name = "ModTime"
Public Const TJ2000 = 0
Public Const TB1950 = (2433282.4235 - 2451545) / 36525
Public Const secToT = (1 / (86400# * 36525#))

Public Const h0Sun = (-0.8333 * DToR)
Public Const h0Planet = (-0.5667 * DToR)

  '{ RiseSet flags }
Public Const ALWAYS_ABOVE = 1 '  { object is always above horizon  }
Public Const ALWAYS_BELOW = 2 '  { object is always below horizon  }
Type tRiseSetTran
    Rise As Double
    Setting  As Double
    Transit As Double
    flags As Long
End Type
