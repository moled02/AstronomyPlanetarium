VERSION 5.00
Begin VB.Form frmEphem 
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12450
   Icon            =   "frmEphem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   12450
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fmeOtherElements 
      Caption         =   "Other Elements"
      Height          =   2535
      Left            =   8400
      TabIndex        =   34
      Top             =   240
      Width           =   3495
      Begin VB.TextBox txtLongitudePerihelium 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   14
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtLongitudeNode 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtInclination 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Longitude Perihelium"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "Longitude Node"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "Inclination"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate"
      Height          =   375
      Left            =   5160
      TabIndex        =   31
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Frame fmeCalculation 
      Caption         =   "Calculation"
      Height          =   1335
      Left            =   240
      TabIndex        =   28
      Top             =   1440
      Width           =   3015
      Begin VB.TextBox txtStartingDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtIntervalInDays 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   17
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtNumberCalculations 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "Starting Date"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Interval in days"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Number tabulations"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.TextBox txtUitvoer 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   27
      Top             =   4200
      Width           =   12015
   End
   Begin VB.Frame fmeObject 
      Caption         =   "Type Object"
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      Begin VB.OptionButton optAsteroid 
         Alignment       =   1  'Right Justify
         Caption         =   "Asteroid"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   600
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton optComet 
         Alignment       =   1  'Right Justify
         Caption         =   "Comet"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame fmeAsteroid 
      Caption         =   "Elements Orbit Asteroid"
      Enabled         =   0   'False
      Height          =   2535
      Left            =   3360
      TabIndex        =   18
      Top             =   240
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox txtMagnitude_G 
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtMagnitude_H 
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtSemi_major_axis 
         Height          =   285
         Left            =   2400
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtEccentricity 
         Height          =   285
         Left            =   2400
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtMeanAnomaly 
         Height          =   285
         Left            =   2400
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtEpoch 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Magnitude G"
         Height          =   255
         Left            =   360
         TabIndex        =   41
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label15 
         Caption         =   "Magnitude H"
         Height          =   255
         Left            =   360
         TabIndex        =   40
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Semi-major axis"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Eccentricity"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Mean anomaly at epoch"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Epoch"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fmeComet 
      Caption         =   "Elements Orbit Comet"
      Height          =   2535
      Left            =   3360
      TabIndex        =   23
      Top             =   240
      Width           =   4815
      Begin VB.TextBox txtMagnitude_k 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   11
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtMagnitude_gc 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtDatePassagePerihelion 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1043
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtEccentricityComet 
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtPerihelionDistance 
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Magnitude k"
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label13 
         Caption         =   "Magnitude g"
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Date Passage Perihelion"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Eccentricity"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Perihelion Distance"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   1080
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmEphem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************)
'                                                                           *)
'                  Copyright (c) 1991-1992 by Jeffrey Sax                   *)
'                            All rights reserved                            *)
'                        Published and Distributed by                       *)
'                           Willmann-Bell, Inc.                             *)
'                             P.O. Box 35025                                *)
'                        Richmond, Virginia 23235                           *)
'                Voice (804) 320-7016 FAX (804) 272-5920                    *)
'                                                                           *)
'                                                                           *)
'                NOTICE TO COMMERCIAL SOFTWARE DEVELOPERS                   *)
'                                                                           *)
'        Prior to distributing software incorporating this code             *)
'        you MUST write Willmann-Bell, Inc. at the above address            *)
'        for validation of your book's (Astronomical Algorithms             *)
'        by Jean Meeus) and software Serial Numbers.  No additional         *)
'        fees will be required BUT you MUST have the following              *)
'        notice at the start of your program(s):                            *)
'                                                                           *)
'                    This program contains code                             *)
'              Copyright (c) 1991-1992 by Jeffrey Sax                       *)
'              and Distributed by Willmann-Bell, Inc.                       *)
'                         Serial #######                                    *)
'                                                                           *)
'****************************************************************************)
' Name:   EPHEM                                                             *)
' Module: EPHEM.PAS                                                         *)
' Type:   Main Program                                                      *)
' Purpose:calculate an ephemeris from orbital elements                      *)
' Version 2.0                                                               *)
' Last modified: October 1, 1992                                            *)
'****************************************************************************)
Private Sub cmdCalculate_Click()
Const COMET = 1
Const ASTEROID = 2

Dim NoIntervals As Long, I As Long, j As Long, Epoch As Long, ObjectType As Long
Dim JD As Double, dJD As Double, T As Double, Obl As Double
Dim RA As Double, Decl As Double
Dim RHelio As TVECTOR, RSun As TVECTOR
Dim SGeo As TSVECTOR
Dim RA1950 As Double, Decl1950 As Double, DistSun As Double
Dim tmp      As Double, PhaseAngle As Double, Magnitude As Double
Dim ddate As tDatum
Dim OrbitEl As TORBITEL
Dim OrbitCon As TORBITCON
Dim ch As String * 1
Dim sZon As TSVECTOR
Dim sAarde As TSVECTOR
Dim SHelio As TSVECTOR
Dim cWaarde As String
On Error GoTo fout_invoer

If optAsteroid.Value Then
    ObjectType = ASTEROID
    'writeln('Enter the object''s orbital elements :')
    'writeln('(All angles are degrees AND DECIMALS!)')
    ddate.jj = Year(Me.txtEpoch)
    ddate.mm = Month(Me.txtEpoch)
    ddate.DD = Day(Me.txtEpoch)
    OrbitEl.T0 = JDToT(KalenderNaarJD(ddate))
    OrbitEl.M0 = Val(Me.txtMeanAnomaly) * DToR
    OrbitEl.E = Val(Me.txtEccentricity)
    OrbitEl.A = Val(Me.txtSemi_major_axis)
    OrbitEl.n = GaussConstant / (OrbitEl.A * Sqr(OrbitEl.A))
  'Debug.Print OrbitEl.n * 180 / PI
Else
    ObjectType = COMET
      'ReadDate('Enter date of passage through perihelion : ', Date)
    ddate.jj = Year(Me.txtDatePassagePerihelion)
    ddate.mm = Month(Me.txtDatePassagePerihelion)
    ddate.DD = Day(Me.txtDatePassagePerihelion)
    OrbitEl.T0 = JDToT(KalenderNaarJD(ddate))
    OrbitEl.M0 = 0
    OrbitEl.E = Val(Me.txtEccentricityComet)
    OrbitEl.Q = Val(Me.txtPerihelionDistance)
    If OrbitEl.E < 1 Then
       OrbitEl.A = OrbitEl.Q / (1 - OrbitEl.E)
       OrbitEl.n = GaussConstant / (OrbitEl.A * Sqr(OrbitEl.A))
    End If
End If
OrbitEl.incl = Val(Me.txtInclination) * DToR
OrbitEl.LonNode = Val(Me.txtLongitudeNode) * DToR
OrbitEl.LonPeri = Val(Me.txtLongitudePerihelium) * DToR
If (ObjectType = COMET) Then
  OrbitEl.MagParam1 = Val(Me.txtMagnitude_gc)
  OrbitEl.MagParam2 = Val(Me.txtMagnitude_k)
Else
  OrbitEl.MagParam1 = Val(Me.txtMagnitude_H)
  OrbitEl.MagParam2 = Val(Me.txtMagnitude_G)
End If
  
ddate.jj = Year(Me.txtStartingDate)
ddate.mm = Month(Me.txtStartingDate)
ddate.DD = Day(Me.txtStartingDate)
NoIntervals = Val(Me.txtNumberCalculations)
JD = KalenderNaarJD(ddate)
dJD = Val(Me.txtIntervalInDays)

Me.txtUitvoer = "Date        R.A.       Decl.               Distance                Mag" & vbCrLf
Me.txtUitvoer = Me.txtUitvoer & "                 J2000.0             Earth           Sun           " & vbCrLf
Me.txtUitvoer = Me.txtUitvoer & "--------------------------------------------------------------------------" & vbCrLf

JD = KalenderNaarJD(ddate)
T = JDToT(JD)
Call CalcOrbitCon(OrbitEl, Obl2000, OrbitCon)

For I = 1 To NoIntervals
    T = JDToT(JD)
    SHelio.l = 0: SHelio.B = 0: SHelio.r = 0
    sAarde.l = 0: sAarde.B = 0: sAarde.r = 0
    Obl = Obliquity(T)
    Call PlanetPosHi(0, T, sAarde, False)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call SphToRect(SGeo, RHelio)
    Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
    Call PrecessFK5(T, 0, RA, Decl)
    Call EquToEcl(RA, Decl, Obliquity(0), SGeo.l, SGeo.B)
    Call SphToRect(SGeo, RSun)
    Call EclVSOP2000_equFK52000(RSun.x, RSun.Y, RSun.Z)

    RHelio.x = 0: RHelio.Y = 0: RHelio.Z = 0
    Call PosRectCo(T, OrbitEl, OrbitCon, RHelio)
    Call RectHelioToGeo(RHelio, RSun, SGeo)
      '{ Do just one light time iteration }
    Call PosRectCo(T - SGeo.r * LightTimeConst, OrbitEl, OrbitCon, RHelio)
 
    Call RectHelioToGeo(RHelio, RSun, SGeo)
  
  '' { Distance from the Sun }
    DistSun = RHelio.x * RHelio.x
    DistSun = DistSun + RHelio.Y * RHelio.Y
    DistSun = DistSun + RHelio.Z * RHelio.Z
    DistSun = Sqr(DistSun)
  
  '  { Magnitude }
    If ObjectType = COMET Then
       Magnitude = CometMagnitude(SGeo.r, DistSun, OrbitEl.MagParam1, OrbitEl.MagParam2)
    Else
       tmp = RSun.x * RSun.x
       tmp = tmp + RSun.Y * RSun.Y
       tmp = tmp + RSun.Z * RSun.Z
       tmp = Sqr(tmp)
       PhaseAngle = CalcPhaseAngle(DistSun, tmp, DistSun)
       Magnitude = AsteroidMagnitude(SGeo.r, DistSun, PhaseAngle, OrbitEl.MagParam1, OrbitEl.MagParam2)
    End If
  
    txtUitvoer = txtUitvoer & StrDate(JDNaarKalender(JD)) & "  " & StrHMS_DMS(SGeo.l * RToD, 3, 1, False, False, "h", 2) & "  " & StrHMS_DMS(SGeo.B * RToD, 3, 0, True, False, "g", 3) & "  "
    txtUitvoer = txtUitvoer & "   " & Format(SGeo.r, ZetFormat(7, 6)) & vbTab & Format(DistSun, ZetFormat(7, 6)) & vbTab '
  
    cWaarde = Format(Magnitude, "0.0")
    cWaarde = String(6 - Len(cWaarde), " ") + cWaarde
    txtUitvoer = txtUitvoer & cWaarde & vbCrLf
             
  'txtUitvoer = txtUitvoer + StrDate(dDate)
    JD = JD + dJD
Next

Exit Sub
fout_invoer:
    MsgBox "Foute invoer", vbCritical, "Ephemeriden"
End Sub



Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim stext As String
    If KeyCode = 67 And Shift = 2 Then
       stext = Me.txtUitvoer
       Clipboard.Clear
       Clipboard.SetText (stext)
    End If
End Sub

Private Sub Form_Load()
Me.fmeComet.Visible = False
Me.fmeComet.Enabled = False
Me.fmeAsteroid.Visible = True
Me.fmeAsteroid.Enabled = True
End Sub

Private Sub optAsteroid_Click()
Me.fmeComet.Visible = False
Me.fmeComet.Enabled = False
Me.fmeAsteroid.Visible = True
Me.fmeAsteroid.Enabled = True
End Sub

Private Sub optComet_Click()
Me.fmeComet.Enabled = True
Me.fmeComet.Visible = True
Me.fmeAsteroid.Visible = False
Me.fmeAsteroid.Enabled = False
End Sub
