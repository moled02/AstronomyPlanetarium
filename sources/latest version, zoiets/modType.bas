Attribute VB_Name = "modType"
Public Type tPlaneet_algemeen
    T As Double
    deltaT As Double
    T0 As Double
    JD_ZT As Double
    JD_WT As Double
    DtofUT As Double
    NutLon As Double
    NutObl As Double
    obl As Double
    LAST As Double
    ObsLat As Double
    ObsLon As Double
    height As Double
    RhoCosPhi As Double
    RhoSinPhi As Double
End Type

Public Type tPlaneet_Maan
    dist As Double
    dkm As Double
    phase As Double
    illum As Double
    diam As Double
    L As Double
    B As Double
    sAarde As TSVECTOR
    sMoon As TSVECTOR
    sHelio As TSVECTOR
    sGeo As TSVECTOR
    sSun As TSVECTOR
    RA_app As Double
    Decl_app As Double
    moonPhysData As TMOONPHYSDATA
    parAngle As Double
    sterbeeld As String
    parallax As Double
    moonHeight As Double
    riseSet As tRiseSetTran
End Type

Public Type tPlaneet_Zon
    sHelio As TSVECTOR
    sAarde As TSVECTOR
    sGeo As TSVECTOR
    RA2000 As Double
    Decl2000 As Double
    SunPhysData As TSUNPHYSDATA
    C As Long
    JDOfCarr As Double
    parAngle As Double
    sterbeeld As String
    RA_app As Double
    Decl_app As Double
    Azimuth As Double
    Hoogte As Double
    RTS As tRiseSetTran
    RTS6 As tRiseSetTran
    RTS18 As tRiseSetTran
End Type

Public Type tPlaneet_Mercurius
    sAarde As TSVECTOR
    sHelio As TSVECTOR
    sGeo As TSVECTOR
    RA2000 As Double
    Decl2000 As Double
    RA_app As Double
    Decl_app As Double
    Azimuth As Double
    Hoogte As Double
    phase As Double
    PhaseAngle As Double
    Elongation As Double
    Magnitude As Double
    Semidiameter As Double
    parAngle As Double
    sterbeeld As String
    RTS As tRiseSetTran
End Type

Public Type tPlaneet_Venus
    sAarde As TSVECTOR
    sHelio As TSVECTOR
    sGeo As TSVECTOR
    RA2000 As Double
    Decl2000 As Double
    RA_app As Double
    Decl_app As Double
    Azimuth As Double
    Hoogte As Double
    phase As Double
    PhaseAngle As Double
    Elongation As Double
    Magnitude As Double
    Semidiameter As Double
    parAngle As Double
    sterbeeld As String
    RTS As tRiseSetTran
End Type


Public Type tPlaneet_Mars
    sAarde As TSVECTOR
    sHelio As TSVECTOR
    sGeo As TSVECTOR
    RA2000 As Double
    Decl2000 As Double
    RA_app As Double
    Decl_app As Double
    Azimuth As Double
    Hoogte As Double
    phase As Double
    PhaseAngle As Double
    Elongation As Double
    Magnitude As Double
    Semidiameter As Double
    parAngle As Double
    sterbeeld As String
    RTS As tRiseSetTran
    MarsPhysData As TMARSPHYSDATA
End Type

Public Type tPlaneet_Jupiter
    sAarde As TSVECTOR
    sHelio As TSVECTOR
    sGeo As TSVECTOR
    RA2000 As Double
    Decl2000 As Double
    RA_app As Double
    Decl_app As Double
    Azimuth As Double
    Hoogte As Double
    phase As Double
    PhaseAngle As Double
    Elongation As Double
    Magnitude As Double
    Semidiameter As Double
    parAngle As Double
    PolarSemiDiameter As Double
    sterbeeld As String
    JupiterPhysData As TJUPITERPHYSDATA
    RTS As tRiseSetTran
'MoonName = Array("", "Io       ", "Europa   ", "Ganymedes", "Callisto ")
    vMaan(4) As TVECTOR
    vsMaan(4) As TVECTOR
    situatieMaan(4, 2) As String
End Type


Public Type tPlaneet_Saturnus
    sAarde As TSVECTOR
    sHelio As TSVECTOR
    sGeo As TSVECTOR
    RA2000 As Double
    Decl2000 As Double
    RA_app As Double
    Decl_app As Double
    Azimuth As Double
    Hoogte As Double
    phase As Double
    PhaseAngle As Double
    Elongation As Double
    Magnitude As Double
    Semidiameter As Double
    parAngle As Double
    PolarSemiDiameter As Double
    sterbeeld As String
    RTS As tRiseSetTran
    SaturnRingData As TSATURNRINGDATA
    AltSaturnRingData As TALTSATURNRINGDATA
    satmanen(9) As tMaan
End Type

Public Type tPlaneet_Uranus
    sAarde As TSVECTOR
    sHelio As TSVECTOR
    sGeo As TSVECTOR
    RA2000 As Double
    Decl2000 As Double
    RA_app As Double
    Decl_app As Double
    Azimuth As Double
    Hoogte As Double
    phase As Double
    PhaseAngle As Double
    Elongation As Double
    Magnitude As Double
    Semidiameter As Double
    parAngle As Double
    sterbeeld As String
    RTS As tRiseSetTran
End Type

Public Type tPlaneet_Neptunus
    sAarde As TSVECTOR
    sHelio As TSVECTOR
    sGeo As TSVECTOR
    RA2000 As Double
    Decl2000 As Double
    RA_app As Double
    Decl_app As Double
    Azimuth As Double
    Hoogte As Double
    phase As Double
    PhaseAngle As Double
    Elongation As Double
    Magnitude As Double
    Semidiameter As Double
    parAngle As Double
    sterbeeld As String
    RTS As tRiseSetTran
End Type

Public Type tPlaneet_Pluto
    sAarde As TSVECTOR
    sHelio As TSVECTOR
    sGeo As TSVECTOR
    RA2000 As Double
    Decl2000 As Double
    RA_app As Double
    Decl_app As Double
    Azimuth As Double
    Hoogte As Double
    phase As Double
    PhaseAngle As Double
    Elongation As Double
    Magnitude As Double
    Semidiameter As Double
    parAngle As Double
    sterbeeld As String
    RTS As tRiseSetTran
End Type


Public Type tPlaneet_Obl
    deltaT As Double
    NutLon As Double
    NutObl As Double
    obl As Double
    ObsLat As Double
    ObsLon As Double
    LAST As Double
    RhoCosPhi As Double
    RhoSinPhi As Double
End Type

