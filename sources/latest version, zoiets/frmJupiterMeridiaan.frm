VERSION 5.00
Begin VB.Form frmJupiterMeridiaan 
   Caption         =   "Meridian op Jupiter"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3870
   Icon            =   "frmJupiterMeridiaan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   3870
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Meridian calculation"
      Height          =   1695
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   3375
      Begin VB.TextBox txtMeridiaanII 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtMeridiaanI 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Meridian II"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Meridian I"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdVolgende 
      Caption         =   "Next"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdVorige 
      Caption         =   "Previous"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox txtMeridiaan 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Longitude"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmJupiterMeridiaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private JDI As Double
Private JDII As Double
Private Sub cmdVolgende_Click()
Dim Om1 As Double, Om2 As Double

    JDI = JDI + 360 / 877.9
    JDII = JDII + 360 / 870.27
    
    Call BerekenMeridiaan(JDI, Om1, Om2)
    Do While Abs(Om1 * RToD - Val(Me.txtMeridiaan)) > 0.01
          JDI = JDI + (Val(Me.txtMeridiaan) - Om1 * RToD) / 877.90003539
        Call BerekenMeridiaan(JDI, Om1, Om2)
    Loop

    Call BerekenMeridiaan(JDII, Om1, Om2)
    Do While Abs(Om2 * RToD - Val(Me.txtMeridiaan)) > 0.01
          JDII = JDII + (Val(Me.txtMeridiaan) - Om2 * RToD) / 870.27003539
        Call BerekenMeridiaan(JDII, Om1, Om2)
    Loop
    Me.txtMeridiaanI = StrDate(JDNaarKalender(JDI)) + " " + StrHMS(Frac(JDNaarKalender(JDI).DD) * Pi2, 2)
    Me.txtMeridiaanII = StrDate(JDNaarKalender(JDII)) + " " + StrHMS(Frac(JDNaarKalender(JDII).DD) * Pi2, 2)
End Sub

Private Sub cmdVorige_Click()
Dim Om1 As Double, Om2 As Double

    JDI = JDI - 360 / 877.9
    JDII = JDII - 360 / 870.27
    
    Call BerekenMeridiaan(JDI, Om1, Om2)
    Do While Abs(Om1 * RToD - Val(Me.txtMeridiaan)) > 0.01
          JDI = JDI + (Val(Me.txtMeridiaan) - Om1 * RToD) / 877.90003539
        Call BerekenMeridiaan(JDI, Om1, Om2)
    Loop

    Call BerekenMeridiaan(JDII, Om1, Om2)
    Do While Abs(Om2 * RToD - Val(Me.txtMeridiaan)) > 0.01
          JDII = JDII + (Val(Me.txtMeridiaan) - Om2 * RToD) / 870.27003539
        Call BerekenMeridiaan(JDII, Om1, Om2)
    Loop
    Me.txtMeridiaanI = StrDate(JDNaarKalender(JDI)) + " " + StrHMS(Frac(JDNaarKalender(JDI).DD) * Pi2, 2)
    Me.txtMeridiaanII = StrDate(JDNaarKalender(JDII)) + " " + StrHMS(Frac(JDNaarKalender(JDII).DD) * Pi2, 2)
End Sub

Private Sub Form_Load()
Dim ddatumI As tDatum
Dim tt As Double
#If FRANS Then
    Me.Caption = "Méridien de Jupiter"
    Me.Label1.Caption = "Longitude"
    Me.Label2.Caption = "Méridian I"
    Me.Label3.Caption = "Méridian II"
    Me.Frame1.Caption = "Calculation de Méridian"
    Me.cmdVorige.Caption = "Dernier"
    Me.cmdVolgende.Caption = "Suivant"
#End If
ddatumI.jj = frmPlanets.Year
ddatumI.mm = frmPlanets.MonthSelect.ListIndex + 1
ddatumI.DD = frmPlanets.DaySelect
tt = (Hrs + Min / 60 + Sec / 3600) / 24
ddatumI.DD = ddatumI.DD + tt
JDI = KalenderNaarJD(ddatumI)
JDII = JDI
End Sub

Private Sub BerekenMeridiaan(JD As Double, ByRef Om1 As Double, ByRef Om2 As Double)
Dim SHelio As TSVECTOR, SGeo As TSVECTOR, SSun As TSVECTOR
Dim SAarde As TSVECTOR
Dim t As Double
Dim Obl As Double
Dim NutLon As Double, NutObl As Double
Dim Parallax As Double, MoonHeight As Double
Dim JupiterPhysData As TJUPITERPHYSDATA
Dim deltaT As Double
Dim sLatitude As String, sLongitude As String
Dim LAST As Double
Dim JD0 As Double

t = JDToT(JD)  ' + TimeZone)
deltaT = ApproxDeltaT(t)
Call NutationConst(t, NutLon, NutObl)
Obl = Obliquity(t)
Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Latitude", _
            sLatitude)
ObsLat = Val(sLatitude) * Pi / 180
Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Longitude", _
            sLongitude)
ObsLon = Val(sLongitude) * Pi / 180

Call PlanetPosHi(0, t, SAarde, True)
Call PlanetPosHi(5, t, SHelio, True)
Call HelioToGeo(SHelio, SAarde, SGeo)
Call PlanetPosHi(5, t - SGeo.r * LightTimeConst, SHelio, True)
Call HelioToGeo(SHelio, SAarde, SGeo)
Call JupiterPhysEphemeris(t + deltaT / 36525 / 86400, SHelio, SAarde, SGeo, _
                               Obl, NutLon, NutObl, _
                              JupiterPhysData)
Om1 = JupiterPhysData.Om1
Om2 = JupiterPhysData.Om2
If Om1 > Pi Then
    Om1 = Om1 - Pi2
End If
If Om2 > Pi Then
    Om2 = Om2 - Pi2
End If
End Sub

