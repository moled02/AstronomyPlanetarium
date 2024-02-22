Attribute VB_Name = "modPhenom"

  '{ EquinoxSolstice }
Public Const SPRINGEQUINOX = 0
Public Const SUMMERSOLSTICE = 1
Public Const AUTUMNEQUINOX = 2
Public Const WINTERSOLSTICE = 3

  '{ ConjunctionOpposition }
Public Const CONJUNCTION = 1
Public Const OPPOSITION = 0
Public Const INFCONJ = 0
Public Const SUPCONJ = 1
Public Const EASTELONGATION = 0
Public Const WESTELONGATION = 1

  '{ MoonPhase }
Public Const NEWMOON = 0
Public Const FIRSTQUARTER = 1
Public Const FULLMOON = 2
Public Const LASTQUARTER = 3

  '{ Eclipses }
Public Const TOTAL = 1
Public Const ANNULAR = 2
Public Const ANNULARTOTAL = 3
Public Const TOTAL_NOT_CENTRAL = 4
Public Const ANNULAR_NOT_CENTRAL = 5
Public Const PARTIAL = 6

Public Const OLDSYSTEM = 1

Public Const PENUMBRAL = 7
Public Const PARTPENUMBRAL = 8

  '{Aard-constanten}
Public Const e2 = 6.70562132949496E-03
  
Type SOLARECLIPSEDATA
    JD    As Double      '{JD, at maximal eclipse}
    Maxmag As Double     '{maximum magnitude of eclipse}
    Gamma As Double      '{Gamma}
    EclipseType As Long  '{What sort of an eclipse}
End Type

Type LUNARECLIPSEDATA
    JD As Double       '{JD, at maximal eclipse}
    sumbra As Double             '{straal umbra}
    spenumbra As Double          '{straal penumbra}
    MagPenumbra As Double        '{magnitude eclipse in Penumbra}
    MagUmbra As Double           '{magnitude eclipse in Umbra}
    SpartUmbra As Double         '{semi-duration partly eclipse}
    StotUmbra As Double          '{semi-duration total eclipse}
    SpartPenumbra As Double  '{semi-duration penumbral eclipse}
    EclipseType As Long    '{eclipsetype}
End Type

Type tBessElmt
    x As Double
    Y As Double
    SD As Double
    cD As Double
    mu As Double
    tF1 As Double
    tF2 As Double
    l1 As Double
    l2 As Double
End Type

Type tAuxElmt
    rho1  As Double
    rho2 As Double
    sd1 As Double
    cd1 As Double
    sd1_d2 As Double
    cd1_d2 As Double
End Type

Type tDiffBess
    mu1 As Double
    x1 As Double
    y1 As Double
    l11 As Double
    l21 As Double
    a11 As Double
    a21 As Double
    b1 As Double
    c11 As Double
    c21 As Double
    d1 As Double
End Type

Type tPredData
    Psi As Double
    Psi1 As Double
    Xi As Double
    nu1 As Double
    Phi As Double
    Lambda As Double
    tPhi As Double
    tPhi1 As Double
    n As Double
    s As Double
    l2 As Double
    tQ0 As Double
    Del As Double
End Type

Type tEclipsePos
    lng As Double
    nb As Double
End Type
  
Type tLimits
    ULimN  As tEclipsePos
    ULimZ As tEclipsePos
    PLimN As tEclipsePos
    PLimZ As tEclipsePos
    l1  As Double
    l2 As Double
End Type

Type tOutCurve
    bQ As Double
    eQ As Double
    pos As tEclipsePos
End Type

Type tMaxEclCurve
    pos1  As tEclipsePos
    pos2 As tEclipsePos
    M1  As Double
    s1 As Double
    m2 As Double
    s2 As Double
End Type

Type tRiseSetCurve
    pos1 As tEclipsePos
    pos2 As tEclipsePos
End Type

Type tRSMaxCurve
    pos1 As tEclipsePos
    pos2 As tEclipsePos
End Type
  
Type tExtremePosTime
    lng As Double
    nb As Double
    time As Double
End Type

Type tExtremes
    fPenPos As tExtremePosTime
    lPenpos As tExtremePosTime
    fUmbPos As tExtremePosTime
    lUmbPos As tExtremePosTime
End Type

Type tGreatestEclipse
    T   As Double
    pos  As tEclipsePos
    magn As Double
End Type

Type tLocalEclipse
    Tb    As Double
    Tm    As Double
    Te    As Double
    Vb    As Double
    Ve    As Double
    Mag   As Double
    Ttotaalb As Double
    Ttotaale As Double
    Vtotaalb As Double
    Vtotaale As Double
    MagTotaal As Double
End Type
