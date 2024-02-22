Attribute VB_Name = "ModGlobal"
Option Explicit
Public Const Pi = 3.14159265358979
Public Const p11 = Pi / 180#
Public Const Pi2 = 2 * Pi
Public Const LightTimeConst = 0.0057755183 / 36525
Public Const DToR = (Pi / 180)
Public Const HToR = (Pi / 12)
Public Const RToD = (180 / Pi)
Public Const RToH = (12 / Pi)
Public Const SToR = (DToR / 3600)
Public Const SolarParallax = (8.794 * SToR)
Public Const EarthRadius = 6378.14
Public Const FK4System = 1
Public Const FK5System = 2
Public Type TVECTOR '{Rectangular coordinates}
    X As Double
    Y As Double
    Z As Double
End Type
Public Type TSVECTOR  '{Spherical coordinates}
    L As Double
    B As Double
    r As Double
End Type
Public Type T4POLY
    p(4) As Double
End Type
Public Type T3POLY
    p(3) As Double
End Type
Public Type T2POLY
    p(2) As Double
End Type
Public Type TEVENT
    JD As Double
    Description As String * 60
    Precision As Integer
End Type

Public Const GaussConstant = 0.01720209895
Public Const Obl2000 = 23.43929111 * DToR

Type TORBITEL
    T0 As Double      '{ Time of epoch                      }
    M0 As Double        '{ Mean anomaly at T0                 }
    A As Double         '{ Semi-major axis (if applicable)    }
    Q As Double         '{ Perihelion distance                }
    n As Double         '{ DAILY increase in the mean anomaly }
    E As Double         '{ Eccentricity of the orbit          }
    LonPeri As Double   '{ Longitude of the perihelion        }
    ArgPeri As Double   '{ Argument of perihelion             }
    LonNode As Double   '{ Longitude of the ascending node    }
    incl As Double      '{ Orbital inclination                }
    MagParam1 As Double '{ for comets: g for asteroids: G    }
    MagParam2 As Double '{ for comets: k for asteroids: H    }
End Type

Type TORBITCON
    A As Double
    B As Double
    C As Double
    aa As Double
    bb As Double
    cc As Double
End Type


' Define global array to hold current coefficients working
' data.
Global Q(2500)

Public g_word As New Application

Type TSINCOSTAB
    W(6) As Double
End Type


