VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSterren 
   AutoRedraw      =   -1  'True
   Caption         =   "Stars"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11895
   Icon            =   "frmSterren.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   11895
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraInputPositions 
      Caption         =   "Input positions (for plotting path of asteroids or comets)"
      Height          =   3015
      Left            =   120
      TabIndex        =   28
      Top             =   3360
      Width           =   11655
      Begin VB.TextBox txtMarkPerPositions 
         Height          =   285
         Left            =   1560
         TabIndex        =   29
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtInvoerPosities 
         Height          =   2655
         Left            =   4800
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   240
         Width           =   6735
      End
      Begin VB.Label Label1 
         Caption         =   "Mark per"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblInvoer 
         Caption         =   "Input positions (RA, Decl):"
         Height          =   615
         Left            =   3360
         TabIndex        =   31
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.ListBox lstPlaneten 
      Height          =   1860
      ItemData        =   "frmSterren.frx":030A
      Left            =   6600
      List            =   "frmSterren.frx":0326
      Style           =   1  'Checkbox
      TabIndex        =   27
      Top             =   1080
      Width           =   2175
   End
   Begin VB.ComboBox cmbHorizon 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmSterren.frx":0369
      Left            =   600
      List            =   "frmSterren.frx":0391
      TabIndex        =   26
      Top             =   1920
      Width           =   1575
   End
   Begin VB.OptionButton optHorizon 
      Caption         =   "Chart of Horizon"
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Frame fraKaart 
      Caption         =   "Chart"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   2880
      TabIndex        =   18
      Top             =   1080
      Width           =   2895
      Begin VB.TextBox txtStraal 
         Height          =   285
         Left            =   1440
         TabIndex        =   23
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtDeclinatie 
         Height          =   285
         Left            =   1440
         TabIndex        =   21
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtRechteKlimming 
         Height          =   285
         Left            =   1440
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblStraal 
         Caption         =   "Radius:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblDeclinatie 
         Caption         =   "Declination:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblRechteKlimming 
         Caption         =   "Right ascension:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CheckBox chkMetLijnen 
      Caption         =   "With lines"
      Height          =   255
      Left            =   8880
      TabIndex        =   17
      Top             =   240
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox chkMetBayer 
      Caption         =   "With Bayernumbers"
      Height          =   255
      Left            =   6360
      TabIndex        =   16
      Top             =   240
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.Frame fraPlaneten 
      Caption         =   "Planets"
      Height          =   2295
      Left            =   8880
      TabIndex        =   8
      Top             =   600
      Width           =   2895
      Begin VB.OptionButton optMerkPerMaand 
         Caption         =   "Mark per month"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtMerkTekenDagen 
         Height          =   285
         Left            =   1920
         TabIndex        =   14
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton optMerkPerDagen 
         Caption         =   "Mark per day"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtEindPeriode 
         Height          =   285
         Left            =   1200
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtStartPeriode 
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblEindPeriode 
         Caption         =   "Period end:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblPeriode 
         Caption         =   "Period start:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdGenereren 
      Caption         =   "Generate"
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   6840
      Width           =   1335
   End
   Begin VB.OptionButton optKaartje 
      Caption         =   "Small chart"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1200
      Width           =   2175
   End
   Begin VB.OptionButton optHuidigeSterrenhemel 
      Caption         =   "Current Sky"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.CheckBox chkMetPlaneten 
      Caption         =   "With planets"
      Height          =   255
      Left            =   6360
      TabIndex        =   4
      Top             =   720
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.TextBox txtGrensmagnitude 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Text            =   "5.5"
      Top             =   240
      Width           =   615
   End
   Begin MSComctlLib.ProgressBar pgbVoortgang 
      Height          =   255
      Left            =   5280
      TabIndex        =   1
      Top             =   7440
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   11880
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11880
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      X1              =   6000
      X2              =   6000
      Y1              =   0
      Y2              =   3120
   End
   Begin VB.Label lblGrensmagnitude 
      Caption         =   "Limiting magnitude:"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblgenerating 
      Caption         =   "Progress:"
      Height          =   255
      Left            =   4080
      TabIndex        =   0
      Top             =   7440
      Width           =   975
   End
End
Attribute VB_Name = "frmSterren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tSter
'    saonum As String * 6
'    sterbld As String * 3
    A As Double
    D As Double
    M As Integer
    flamsteed As Byte
    bayer As String * 2
End Type
Private Type tlijn
    ster1 As tSter
    ster2 As tSter
End Type
Private rtf_code(14) As String
Private Const schuif As Long = 1600
Private Const schaalfactor As Double = 2.6
Private objspecialfolder As New clsSpecialFolder
Private sTempName As String
Private nfile
Private Sub chkMetPlaneten_Click()
    fraPlaneten.Enabled = chkMetPlaneten
    lstPlaneten.Enabled = chkMetPlaneten
End Sub

Private Sub cmdGenereren_Click()
Dim sLatitude As String
Dim dat As tDatum
Dim dRechteKlimming As Double
Dim dDecinatie As Double
Dim dStraal As Double
Dim jdB As Double, jde As Double
Dim I As Long

dat.jj = frmPlanets.Year
dat.mm = frmPlanets.MonthSelect.ListIndex + 1
dat.DD = frmPlanets.DaySelect
tt = (frmPlanets.Hrs + frmPlanets.Min / 60 + frmPlanets.Sec / 3600) / 24
dat.DD = dat.DD + tt
Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Latitude", _
            sLatitude)
For I = 1 To UBound(rtf_code)
    rtf_code(I) = ""
Next

sTempName = objspecialfolder.TemporaryFolder + "\test_" + Format(Now(), "yyyy-mm-dd_hh.mm.ss") + ".rtf"

Open App.Path + "\rtf_standaards.txt" For Input As #2
Do While Not EOF(2)
  Line Input #2, sRegel
  Select Case sRegel
      Case Is = "[TEKSTVAK]"
          lrtf_code = 1
          sRegel = ""
      Case Is = "[CIRKEL-KLEUR]"
          lrtf_code = 2
          sRegel = ""
      Case Is = "[CIRKEL-GEWOON]"
          lrtf_code = 3
          sRegel = ""
      Case Is = "[EINDE]"
          lrtf_code = 4
          sRegel = ""
      Case Is = "[BEGIN]"
          lrtf_code = 5
          sRegel = ""
      Case Is = "[LIJN]"
          lrtf_code = 6
          sRegel = ""
      Case Is = "[GEBOGEN-LIJN]"
          lrtf_code = 7
          sRegel = ""
      Case Is = "[CIRKEL-GRIJS]"
          lrtf_code = 8
          sRegel = ""
      Case Is = "[BOOG]"
          lrtf_code = 9
          sRegel = ""
      Case Is = "[TEKSTVAK-MET-LIJN-ZWARTE-TEKST]"
          lrtf_code = 10
          sRegel = ""
      Case Is = "[NIEUWE-PAGINA]"
          lrtf_code = 11
          sRegel = ""
      Case Is = "[NIEUW-KLEUR-LIJN]]"
          lrtf_code = 12
          sRegel = ""
      Case Is = "[TEKSTVAK-ZONDER-LIJN-TEKST]"
          lrtf_code = 13
          sRegel = ""
      Case Is = "[JUPITER-IMAGE]"
          lrtf_code = 14
          sRegel = ""
  End Select
  If sRegel <> "" Then rtf_code(lrtf_code) = rtf_code(lrtf_code) + sRegel + vbCrLf
Loop
Close #2
nfile = 2
Open sTempName For Output As #nfile
Print #nfile, rtf_code(5);

pgbVoortgang.Value = 0
If optHorizon Then
    dRechteKlimming = 24# * PlaatselijkeSterrentijd(dat)
    dDecinatie = ReadDMS(sLatitude)
    dStraal = 90
    If cmbHorizon.ListIndex = -1 Then
        Close #2
        Exit Sub
    End If
    Call stertekHorizon(24# * PlaatselijkeSterrentijd(dat), ReadDMS(sLatitude), cmbHorizon.ItemData(cmbHorizon.ListIndex), Val(txtGrensmagnitude))
End If
If optHuidigeSterrenhemel Or optKaartje Then
    If Me.optHuidigeSterrenhemel Then
        dRechteKlimming = 24# * PlaatselijkeSterrentijd(dat)
        dDecinatie = ReadDMS(sLatitude)
        dStraal = 90
        Call stertek(24# * PlaatselijkeSterrentijd(dat), ReadDMS(sLatitude), 90, Val(txtGrensmagnitude))
    ElseIf Me.optKaartje Then
        dRechteKlimming = Val(txtRechteKlimming)
        dDecinatie = ReadDMS(txtDeclinatie)
        dStraal = Val(Me.txtStraal)
    End If
    Call stertek(dRechteKlimming, dDecinatie, dStraal, Val(txtGrensmagnitude))
End If
If chkMetPlaneten Then
    dat.DD = Val(Left(Me.txtStartPeriode, 2))
    dat.mm = Val(Mid(Me.txtStartPeriode, 4, 2))
    dat.jj = Val(Mid(Me.txtStartPeriode, 7, 4))
    jdB = KalenderNaarJD(dat)
    dat.DD = Val(Left(Me.txtEindPeriode, 2))
    dat.mm = Val(Mid(Me.txtEindPeriode, 4, 2))
    dat.jj = Val(Mid(Me.txtEindPeriode, 7, 4))
    jde = KalenderNaarJD(dat)
    For I = 1 To lstPlaneten.ListCount
        If lstPlaneten.Selected(I - 1) Then
            Call PlaneetTekenen(dRechteKlimming, dDecinatie, dStraal, Val(txtGrensmagnitude), _
                                 jdB, jde, _
                                 lstPlaneten.ItemData(I - 1), lstPlaneten.ItemData(I - 1), _
                                 optMerkPerDagen, Val(txtMerkTekenDagen), optHorizon)
        End If
    Next
End If
If Not Trim(Me.txtInvoerPosities) = vbNullString Then
    Call InvoerTekenen(dRechteKlimming, dDecinatie, dStraal, Val(txtGrensmagnitude), _
                  Val(Me.txtMarkPerPositions), optHorizon)
End If
Print #nfile, rtf_code(4); 'het einde toevoegen
Close #nfile

On Error GoTo word_open:
g_word.Visible = True
g_word.Documents.Open FileName:=Chr(34) & sTempName & Chr(34), ConfirmConversions:=False
' Shell "Winword " & Chr(34) & sTempName & Chr(34), vbNormalFocus
g_word.Activate
pgbVoortgang.Value = 0
Exit Sub

word_open:
    If Err.Number = 462 Then 'Word waarschijnlijk gesloten
        Set g_word = New Application
        Resume
    End If
End Sub

Sub stertek(ByVal RK As Double, ByVal Dec As Double, ByVal r As Double, ByVal mag As Double)

Dim PI_2 As Double, PI_180 As Double, sxdec As Double, cxdec As Double
Dim Az As Double, hg As Double, x As Long, Y As Long, midx As Long, midy As Long
Dim Q As Double
Dim ster As tSter
Dim sRegel As String
Dim lrtf_code As Long
Dim aantal_sterren As Long

  midx = 19140 / schaalfactor: midy = 14595 / schaalfactor
  Open App.Path + "\sterren.bin" For Random As #1 Len = LenB(ster)
  Call print_rtf_circle(2, rtf_code(8), midx, midy, midy, 15921906) 'grote cirkel
  If Me.chkMetLijnen Then Call sterlijn(RK, Dec, r, mag) 'alleen als de lijnen ook getekend moeten worden
  
  PI_2 = Pi / 2
  PI_180 = Pi / 180
  RK = RK * Pi / 12
  Dec = Dec * PI_180
  r = r * PI_180
  sxdec = Sin(Dec)
  cxdec = Cos(Dec)
  mag = mag * 10 '(* meegegeven magnitude is factor 10 te groot *)

  Get #1, , ster
  Do While (Not EOF(1)) And (ster.M <= mag)
    aantal_sterren = (aantal_sterren + 1) Mod 5000
     pgbVoortgang.Value = aantal_sterren / 50
      If (ster.D > Dec - r) And (ster.D < Dec + r) Then
            Q = RK - ster.A
            hg = asin(Sin(ster.D) * sxdec + Cos(ster.D) * Cos(Q) * cxdec)
            If (r > PI_2 - hg) Then Az = atan2(Sin(Q), Cos(Q) * sxdec - tan(ster.D) * cxdec)
     
            hg = PI_2 - hg
            If (r > hg) Then
                x = Int(midx + midy * Sin(Az) * hg / r)
                Y = Int(midy + midy * Cos(Az) * hg / r)
                If (x > 0) And (Y > 0) And (x < 2 * midx) And (Y < 2 * midy) Then
'                    Teken_ster x, Y, Straal(0.1 * ster.M, mag)
                    Call print_rtf_circle(2, rtf_code(2), x, Y, Straal(0.1 * ster.M, mag), 0)
                    If Me.chkMetBayer And ster.M < mag - 30 Then 'alleen als Bayer gevraagd
                        Call print_rtf_textbox(2, rtf_code(1), x + Straal(0.1 * ster.M, mag), Y, ster.bayer)
                    End If
                End If
            End If
        End If
        Get #1, , ster
        DoEvents
  Loop
  Close (1)
  pgbVoortgang.Value = 100
  DoEvents
 End Sub
Sub stertekHorizon(ByVal LST As Double, ByVal nb As Double, ByVal Az0 As Double, ByVal mag As Double)

Dim PI_2 As Double, PI_180 As Double, sxdec As Double, cxdec As Double
Dim Az As Double, Azt As Double, hg As Double, x As Long, Y As Long, midx As Long, midy As Long
Dim Q As Double
Dim ster As tSter
Dim sRegel As String
Dim lrtf_code As Long
Dim aantal_sterren As Long

  midx = 19140 / schaalfactor: midy = 14595 / schaalfactor
  Open App.Path + "\sterren.bin" For Random As #1 Len = LenB(ster)
  Call print_rtf_boog(2, rtf_code(9), midx - midy / 2, midy / 2, midy / 2, 1, 0) 'grote cirkel
  Call print_rtf_boog(2, rtf_code(9), midx + midy / 2, midy / 2, midy / 2, 0, 0) 'grote cirkel
  Call print_rtf_lijn(2, rtf_code(6), midx - midy, midy, midx + midy, midy)
  
'  Call print_rtf_circle(2, rtf_code(7), midx, midy, midy) 'grote cirkel
  ' dit even later
  If Me.chkMetLijnen Then Call sterlijnHorizon(LST, nb, Az0, mag) 'alleen als de lijnen ook getekend moeten worden
  
  PI_2 = Pi / 2
  PI_180 = Pi / 180
  nb = nb * PI_180
  r = r * PI_180
  mag = mag * 10 '(* meegegeven magnitude is factor 10 te groot *)
  Az0 = PI_180 * Az0
  LST = LST * Pi / 12

  Get #1, , ster
  Do While (Not EOF(1)) And (ster.M <= mag)
    aantal_sterren = (aantal_sterren + 1) Mod 5000
    pgbVoortgang.Value = aantal_sterren / 50
    Call EquToHor(ster.A, ster.D, LST, nb, Az, hg)
    Azt = Az - Az0
    If Azt > Pi Then Azt = Azt - 2 * Pi
    If Azt < -Pi Then Azt = Azt + 2 * Pi
    
    If (hg > 0) And (Abs(Azt) < PI_2) Then
        x = Int(midx + midy * Azt * Sqr(1 - hg / PI_2 * hg / PI_2) / PI_2)
        Y = Int(midy - midy * hg / PI_2)
    '                   Teken_ster x, Y, Straal(0.1 * ster.M, mag)
            Call print_rtf_circle(2, rtf_code(2), x, Y, Straal(0.1 * ster.M, mag), 0)
            If Me.chkMetBayer And ster.M < mag - 30 Then 'alleen als Bayer gevraagd
                Call print_rtf_textbox(2, rtf_code(1), x + Straal(0.1 * ster.M, mag), Y, ster.bayer)
            End If
    End If
    Get #1, , ster
    DoEvents
  Loop
  Close (1)
  pgbVoortgang.Value = 100
  DoEvents
 End Sub

Function Straal(ByVal SterMag As Double, ByVal mag As Double) As Long
    Straal = 20 * (mag / 10 - SterMag)
End Function

Sub print_rtf_circle(nfile As Long, ByVal srtf_code As String, ByVal x As Long, ByVal Y As Long, ByVal r As Long, _
                     Optional ByVal Vul_kleur As Long = 0)
Dim nPos As Long
    x = x - schuif
    Y = Y + schuif
    nPos = InStr(srtf_code, "<LEFT>")
    Do While InStr(srtf_code, "<LEFT>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(x - r, "0") + Mid(srtf_code, nPos + 6)
        nPos = InStr(srtf_code, "<LEFT>")
    Loop
    nPos = InStr(srtf_code, "<TOP>")
    Do While InStr(srtf_code, "<TOP>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(Y - r, "0") + Mid(srtf_code, nPos + 5)
        nPos = InStr(srtf_code, "<TOP>")
    Loop
    nPos = InStr(srtf_code, "<RIGHT>")
    Do While InStr(srtf_code, "<RIGHT>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(x + r, "0") + Mid(srtf_code, nPos + 7)
        nPos = InStr(srtf_code, "<RIGHT>")
    Loop
    nPos = InStr(srtf_code, "<BOTTOM>")
    Do While InStr(srtf_code, "<BOTTOM>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(Y + r, "0") + Mid(srtf_code, nPos + 8)
        nPos = InStr(srtf_code, "<BOTTOM>")
    Loop
    nPos = InStr(srtf_code, "<BOTTOM-TOP>")
    Do While InStr(srtf_code, "<BOTTOM-TOP>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(2 * r, "0") + Mid(srtf_code, nPos + 12)
        nPos = InStr(srtf_code, "<BOTTOM-TOP>")
    Loop
    nPos = InStr(srtf_code, "<RIGHT-LEFT>")
    Do While InStr(srtf_code, "<RIGHT-LEFT>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(2 * r, "0") + Mid(srtf_code, nPos + 12)
        nPos = InStr(srtf_code, "<RIGHT-LEFT>")
    Loop
    nPos = InStr(srtf_code, "<VUL-KLEUR>")
    Do While InStr(srtf_code, "<VUL-KLEUR>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(Vul_kleur, "0") + Mid(srtf_code, nPos + 11)
        nPos = InStr(srtf_code, "<VUL-KLEUR>")
    Loop
    Print #nfile, srtf_code;
End Sub
    
Sub print_rtf_boog(nfile As Long, ByVal srtf_code As String, ByVal x As Long, ByVal Y As Long, ByVal r As Long, ByVal FlipH As Long, FlipV As Long)
Dim nPos As Long
    x = x - schuif
    Y = Y + schuif
    nPos = InStr(srtf_code, "<LEFT>")
    Do While InStr(srtf_code, "<LEFT>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(x - r, "0") + Mid(srtf_code, nPos + 6)
        nPos = InStr(srtf_code, "<LEFT>")
    Loop
    nPos = InStr(srtf_code, "<TOP>")
    Do While InStr(srtf_code, "<TOP>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(Y - r, "0") + Mid(srtf_code, nPos + 5)
        nPos = InStr(srtf_code, "<TOP>")
    Loop
    nPos = InStr(srtf_code, "<RIGHT>")
    Do While InStr(srtf_code, "<RIGHT>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(x + r, "0") + Mid(srtf_code, nPos + 7)
        nPos = InStr(srtf_code, "<RIGHT>")
    Loop
    nPos = InStr(srtf_code, "<BOTTOM>")
    Do While InStr(srtf_code, "<BOTTOM>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(Y + r, "0") + Mid(srtf_code, nPos + 8)
        nPos = InStr(srtf_code, "<BOTTOM>")
    Loop
    nPos = InStr(srtf_code, "<BOTTOM-TOP>")
    Do While InStr(srtf_code, "<BOTTOM-TOP>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(2 * r, "0") + Mid(srtf_code, nPos + 12)
        nPos = InStr(srtf_code, "<BOTTOM-TOP>")
    Loop
    nPos = InStr(srtf_code, "<RIGHT-LEFT>")
    Do While InStr(srtf_code, "<RIGHT-LEFT>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(2 * r, "0") + Mid(srtf_code, nPos + 12)
        nPos = InStr(srtf_code, "<RIGHT-LEFT>")
    Loop
    nPos = InStr(srtf_code, "<FLIP-H>")
    Do While InStr(srtf_code, "<FLIP-H>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(FlipH, "0") + Mid(srtf_code, nPos + 8)
        nPos = InStr(srtf_code, "<FLIP-H>")
    Loop
    nPos = InStr(srtf_code, "<FLIP-V>")
    Do While InStr(srtf_code, "<FLIP-V>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(FlipV, "0") + Mid(srtf_code, nPos + 8)
        nPos = InStr(srtf_code, "<FLIP-V>")
    Loop
    Print #nfile, srtf_code;
End Sub
    
Sub print_rtf_lijn(nfile As Long, ByVal srtf_code As String, ByVal x As Long, ByVal Y As Long, ByVal x1 As Long, ByVal y1 As Long)
Dim nPos As Long
Dim links As Long, top As Long
Dim rechts As Long, onder As Long
Dim FlipV As Long, FlipH As Long
Dim tx1 As Long, tx2 As Long, ty1 As Long, ty2 As Long

x = x - schuif
Y = Y + schuif
x1 = x1 - schuif
y1 = y1 + schuif
If x < x1 Then tx1 = x: ty1 = Y: tx2 = x1: ty2 = y1 Else tx2 = x: ty2 = Y: tx1 = x1: ty1 = y1
    links = tx1
    rechts = tx2
    If ty1 < ty2 Then
        top = ty1
        onder = ty2
    Else
        top = ty2
        onder = ty1
    End If
    If tx1 - tx2 <> 0 Then
        If (ty2 - ty1) < 0 Then If (tx2 - tx1) > 0 Then FlipV = 1 'verticaal spiegelen, anders wordt de lijn verkeerd getekend
    End If
    nPos = InStr(srtf_code, "<LEFT>")
    Do While InStr(srtf_code, "<LEFT>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(links, "0") + Mid(srtf_code, nPos + 6)
        nPos = InStr(srtf_code, "<LEFT>")
    Loop
    nPos = InStr(srtf_code, "<TOP>")
    Do While InStr(srtf_code, "<TOP>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(top, "0") + Mid(srtf_code, nPos + 5)
        nPos = InStr(srtf_code, "<TOP>")
    Loop
    nPos = InStr(srtf_code, "<RIGHT>")
    Do While InStr(srtf_code, "<RIGHT>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(rechts, "0") + Mid(srtf_code, nPos + 7)
        nPos = InStr(srtf_code, "<RIGHT>")
    Loop
    nPos = InStr(srtf_code, "<BOTTOM>")
    Do While InStr(srtf_code, "<BOTTOM>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(onder, "0") + Mid(srtf_code, nPos + 8)
        nPos = InStr(srtf_code, "<BOTTOM>")
    Loop
    nPos = InStr(srtf_code, "<BOTTOM-TOP>")
    Do While InStr(srtf_code, "<BOTTOM-TOP>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(onder - top, "0") + Mid(srtf_code, nPos + 12)
        nPos = InStr(srtf_code, "<BOTTOM-TOP>")
    Loop
    nPos = InStr(srtf_code, "<RIGHT-LEFT>")
    Do While InStr(srtf_code, "<RIGHT-LEFT>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(rechts - links, "0") + Mid(srtf_code, nPos + 12)
        nPos = InStr(srtf_code, "<RIGHT-LEFT>")
    Loop
    nPos = InStr(srtf_code, "<FLIP-H>")
    Do While InStr(srtf_code, "<FLIP-H>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(FlipH, "0") + Mid(srtf_code, nPos + 8)
        nPos = InStr(srtf_code, "<FLIP-H>")
    Loop
    nPos = InStr(srtf_code, "<FLIP-V>")
    Do While InStr(srtf_code, "<FLIP-V>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(FlipV, "0") + Mid(srtf_code, nPos + 8)
        nPos = InStr(srtf_code, "<FLIP-V>")
    Loop
    Print #nfile, srtf_code;
End Sub
Sub print_rtf_textbox(nfile As Long, ByVal srtf_code As String, ByVal x As Long, ByVal Y As Long, ByVal stext As String)
Dim nPos As Long
Const nGroot As Long = 244
Dim sBayerPos2 As String
    
    x = x - schuif: Y = Y + schuif
    If Trim(stext) = vbNullString Then
        Exit Sub
    End If
    nPos = InStr(srtf_code, "<LEFT>")
    Do While InStr(srtf_code, "<LEFT>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(x, "0") + Mid(srtf_code, nPos + 6)
        nPos = InStr(srtf_code, "<LEFT>")
    Loop
    nPos = InStr(srtf_code, "<TOP>")
    Do While InStr(srtf_code, "<TOP>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(Y - nGroot, "0") + Mid(srtf_code, nPos + 5)
        nPos = InStr(srtf_code, "<TOP>")
    Loop
    nPos = InStr(srtf_code, "<RIGHT>")
    Do While InStr(srtf_code, "<RIGHT>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(x + nGroot, "0") + Mid(srtf_code, nPos + 7)
        nPos = InStr(srtf_code, "<RIGHT>")
    Loop
    nPos = InStr(srtf_code, "<BOTTOM>")
    Do While InStr(srtf_code, "<BOTTOM>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(Y, "0") + Mid(srtf_code, nPos + 8)
        nPos = InStr(srtf_code, "<BOTTOM>")
    Loop
    nPos = InStr(srtf_code, "<BOTTOM-TOP>")
    Do While InStr(srtf_code, "<BOTTOM-TOP>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(nGroot, "0") + Mid(srtf_code, nPos + 12)
        nPos = InStr(srtf_code, "<BOTTOM-TOP>")
    Loop
    nPos = InStr(srtf_code, "<RIGHT-LEFT>")
    Do While InStr(srtf_code, "<RIGHT-LEFT>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(nGroot, "0") + Mid(srtf_code, nPos + 12)
        nPos = InStr(srtf_code, "<RIGHT-LEFT>")
    Loop
    nPos = InStr(srtf_code, "<SYMBOOL>")
    Do While InStr(srtf_code, "<SYMBOOL>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(97 + Asc(Left(stext, 1)) - Asc("a"), "0") + Mid(srtf_code, nPos + 9)
        nPos = InStr(srtf_code, "<SYMBOOL>")
    Loop
    sBayerPos2 = Mid(stext, 2, 1)
    If sBayerPos2 = "0" Then sBayerPos2 = vbNullString
    nPos = InStr(srtf_code, "<SUPER>")
    Do While InStr(srtf_code, "<SUPER>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + sBayerPos2 + Mid(srtf_code, nPos + 7)
        nPos = InStr(srtf_code, "<SUPER>")
    Loop
    Print #nfile, srtf_code;
End Sub

Sub sterlijn(ByVal RK As Double, ByVal Dec As Double, ByVal r As Double, ByVal mag As Double)

Dim PI_2 As Double, PI_180 As Double, sxdec As Double, cxdec As Double
Dim Az As Double, hg As Double, x As Long, Y As Long, midx As Long, midy As Long
Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
Dim Q As Double, shg As Double
Dim lijn As tlijn
Dim aantal_lijnen As Long

  mag = mag * 10 '(* meegegeven magnitude is factor 10 te groot *)

'  midx = picCanvas.ScaleWidth / schaalfactor: midy = picCanvas.ScaleHeight / schaalfactor
  midx = 19140 / schaalfactor: midy = 14595 / schaalfactor
'  picCanvas.Circle (midx, midy), midy
  Open App.Path + "\sterlijn.bin" For Random As #3 Len = LenB(lijn)
  
  PI_2 = Pi / 2
  PI_180 = Pi / 180
  RK = RK * Pi / 12
  Dec = Dec * PI_180
  r = r * PI_180
  sxdec = Sin(Dec)
  cxdec = Cos(Dec)

  Get #3, , lijn
  
  Do While (Not EOF(3))
    aantal_lijnen = aantal_lijnen + 1
    pgbVoortgang.Value = aantal_lijnen * LenB(lijn) / LOF(1) * 100
      With lijn
      If ((.ster1.D > Dec - r) And (.ster1.D < Dec + r) And (.ster1.M <= mag)) And _
         ((.ster2.D > Dec - r) And (.ster2.D < Dec + r) And (.ster2.M <= mag)) Then
            Q = RK - .ster1.A
            shg = Sin(.ster1.D) * sxdec + Cos(.ster1.D) * Cos(Q) * cxdec
            If Abs(shg) <= 1 Then
                hg = asin(shg)
            Else
                hg = -99999
            End If
            If (r > PI_2 - hg) Then Az1 = atan2(Sin(Q), Cos(Q) * sxdec - tan(.ster1.D) * cxdec)
            hg1 = PI_2 - hg
            Q = RK - .ster2.A
            shg = Sin(.ster2.D) * sxdec + Cos(.ster2.D) * Cos(Q) * cxdec
            If Abs(shg) <= 1 Then
                hg = asin(shg)
            Else
                hg = -99999
            End If
            If (r > PI_2 - hg) Then Az2 = atan2(Sin(Q), Cos(Q) * sxdec - tan(.ster2.D) * cxdec)
            hg2 = PI_2 - hg
            
            If (r > hg1) And (r > hg2) Then
                x1 = Int(midx + midy * Sin(Az1) * hg1 / r)
                y1 = Int(midy + midy * Cos(Az1) * hg1 / r)
                x2 = Int(midx + midy * Sin(Az2) * hg2 / r)
                y2 = Int(midy + midy * Cos(Az2) * hg2 / r)
                If (x1 > 0) And (y1 > 0) And (x1 < 2 * midx) And (y1 < 2 * midy) And _
                   (x2 > 0) And (y2 > 0) And (x2 < 2 * midx) And (y2 < 2 * midy) Then
 '                  picCanvas.Line (x1, y1)-(x2, y2)
                   Call print_rtf_lijn(2, rtf_code(6), x1, y1, x2, y2)
                End If
            End If
        End If
        End With
        Get #3, , lijn
  Loop
  Close (3)
  DoEvents
 End Sub

Sub sterlijnHorizon(ByVal LST As Double, ByVal nb As Double, ByVal Az0 As Double, ByVal mag As Double)

Dim PI_2 As Double, PI_180 As Double, sxdec As Double, cxdec As Double
Dim Az As Double, hg As Double, x As Long, Y As Long, midx As Long, midy As Long
Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
Dim Q As Double, shg As Double
Dim lijn As tlijn
Dim aantal_lijnen As Long

  mag = mag * 10 '(* meegegeven magnitude is factor 10 te groot *)

'  midx = picCanvas.ScaleWidth / schaalfactor: midy = picCanvas.ScaleHeight / schaalfactor
  midx = 19140 / schaalfactor: midy = 14595 / schaalfactor
'  picCanvas.Circle (midx, midy), midy
  Open App.Path + "\sterlijn.bin" For Random As #3 Len = LenB(lijn)
  
  PI_2 = Pi / 2
  PI_180 = Pi / 180
  nb = nb * PI_180
  r = r * PI_180
  Az0 = PI_180 * Az0
  LST = LST * Pi / 12

  Get #3, , lijn
  
  Do While (Not EOF(3))
    aantal_lijnen = aantal_lijnen + 1
    pgbVoortgang.Value = aantal_lijnen * LenB(lijn) / LOF(1) * 100
      With lijn
      If (.ster1.M <= mag) And (.ster2.M <= mag) Then
            Call EquToHor(.ster1.A, .ster1.D, LST, nb, Az, hg)
            Azt = Az - Az0
            If Azt > Pi Then Azt = Azt - 2 * Pi
            If Azt < -Pi Then Azt = Azt + 2 * Pi
            Az1 = Azt: hg1 = hg
            
            Call EquToHor(.ster2.A, .ster2.D, LST, nb, Az, hg)
            Azt = Az - Az0
            If Azt > Pi Then Azt = Azt - 2 * Pi
            If Azt < -Pi Then Azt = Azt + 2 * Pi
            Az2 = Azt: hg2 = hg
                
            If (hg1 > 0) And (Abs(Az1) < PI_2) And _
            (hg2 > 0) And (Abs(Az2) < PI_2) Then
                x1 = Int(midx + midy * Az1 * Sqr(1 - hg1 / PI_2 * hg1 / PI_2) / PI_2)
                y1 = Int(midy - midy * hg1 / PI_2)
                x2 = Int(midx + midy * Az2 * Sqr(1 - hg2 / PI_2 * hg2 / PI_2) / PI_2)
                y2 = Int(midy - midy * hg2 / PI_2)
                If (x1 > 0) And (y1 > 0) And (x1 < 2 * midx) And (y1 < 2 * midy) And _
                   (x2 > 0) And (y2 > 0) And (x2 < 2 * midx) And (y2 < 2 * midy) Then
 '                  picCanvas.Line (x1, y1)-(x2, y2)
                   Call print_rtf_lijn(2, rtf_code(6), x1, y1, x2, y2)
                End If
            End If
        End If
        End With
        Get #3, , lijn
  Loop
  Close (3)
  DoEvents
 End Sub

Private Sub calcpos(ByVal Planet As Long, ByVal T As Double, RA As Double, Decl As Double)
Dim Obl As Double
Dim NutLon As Double
Dim NutObl As Double
Dim SHelio As TSVECTOR, SGeo As TSVECTOR, SSun As TSVECTOR
    'Q1 = SHelio, Q2 = SGeo
Dim sAarde As TSVECTOR
Dim TAarde As TVECTOR
Dim sZon As TSVECTOR
Dim TPluto As TVECTOR

    'berekenen posities voor epoch 2000
    Obl = Obliquity(T)
    Call NutationConst(T, NutLon, NutObl)
    
If Planet < 9 Then
    ' Main Calculations }
    Call PlanetPosHi(0, T, sAarde)
    Call PlanetPosHi(Planet, T, SHelio)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call PlanetPosHi(Planet, T - SGeo.r * LightTimeConst, SHelio)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
    Call PrecessFK5(T, 0#, RA, Decl)
Else 'Pluto, alhoewel achterhaald planeet
    Call PlanetPosHi(0, T, sAarde)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call SphToRect(SGeo, TAarde)
    Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
    ' Call Reduction2000(0, RA, Decl)
    'coordinaten omzetten naar J2000
    Call PrecessFK5(T, 0#, RA, Decl)
End If
End Sub

Private Sub Form_Load()
Dim JD As Double
Dim dat As tDatum

dat.DD = 1
dat.mm = frmPlanets.MonthSelect.ListIndex + 1
dat.jj = frmPlanets.Year
txtStartPeriode = "01-" & Format(dat.mm, "00") & "-" & Format(frmPlanets.Year, "0000")
dat = JDNaarKalender(KalenderNaarJD(dat) + 40)
dat.DD = 1
'dat = JDNaarKalender(KalenderNaarJD(dat) - 1)
txtEindPeriode = Format(dat.DD, "00") & "-" & Format(dat.mm, "00") & "-" & Format(dat.jj, "0000")
optMerkPerMaand = True
#If FRANS Then
    Me.Caption = "Carte du Ciel"
    Me.lblDeclinatie.Caption = "Declination"
    Me.lblEindPeriode.Caption = "Période Fin"
    Me.lblgenerating.Caption = "Progression"
    Me.lblGrensmagnitude.Caption = "Magnitude Limite"
    Me.lblPeriode.Caption = "Période"
    Me.lblRechteKlimming.Caption = "Ascension Droite"
    Me.lblStraal.Caption = "Rayon"
    Me.chkMetBayer.Caption = "Avec désignation Bayer"
    Me.chkMetLijnen.Caption = "Avec ligne"
    Me.chkMetPlaneten.Caption = "Avec planète"
    Me.cmdGenereren.Caption = "Faire"
    Me.optHorizon.Caption = "Horizon"
    Me.optHuidigeSterrenhemel.Caption = "Ciel actuel"
    Me.optKaartje.Caption = "Carte"
    Me.optMerkPerDagen.Caption = "Marque par Jour"
    Me.optMerkPerMaand.Caption = "Marque par Mois"
    Me.lstPlaneten.Clear
    lstPlaneten.AddItem "Mercure"
    lstPlaneten.AddItem "Vénus"
    lstPlaneten.AddItem "Mars"
    lstPlaneten.AddItem "Jupiter"
    lstPlaneten.AddItem "Saturne"
    lstPlaneten.AddItem "Uranus"
    lstPlaneten.AddItem "Neptune"
    lstPlaneten.AddItem "Pluton"
    lstPlaneten.ItemData(0) = 1
    lstPlaneten.ItemData(1) = 2
    lstPlaneten.ItemData(2) = 4
    lstPlaneten.ItemData(3) = 5
    lstPlaneten.ItemData(4) = 6
    lstPlaneten.ItemData(5) = 7
    lstPlaneten.ItemData(6) = 8
    lstPlaneten.ItemData(7) = 9
#End If
End Sub

Private Sub optHorizon_Click()
    cmbHorizon.Enabled = True
End Sub

Private Sub optHuidigeSterrenhemel_Click()
Dim JD As Double
Dim dat As tDatum
fraKaart.Enabled = optKaartje.Value
txtGrensmagnitude = "5.5"

dat.DD = 1
dat.mm = frmPlanets.MonthSelect.ListIndex + 1
dat.jj = frmPlanets.Year
txtStartPeriode = "01-" & Format(dat.mm, "00") & "-" & Format(frmPlanets.Year, "0000")
dat = JDNaarKalender(KalenderNaarJD(dat) + 40)
dat.DD = 1
'dat = JDNaarKalender(KalenderNaarJD(dat) - 1)
txtEindPeriode = Format(dat.DD, "00") & "-" & Format(dat.mm, "00") & "-" & Format(frmPlanets.Year, "0000")
optMerkPerMaand = True: cmbHorizon.Enabled = False
End Sub

Private Sub optKaartje_Click()
fraKaart.Enabled = optKaartje.Value

txtStartPeriode = "01-01-" & Format(frmPlanets.Year, "0000")
txtEindPeriode = "01-01-" & Format(frmPlanets.Year + 1, "0000")
cmbHorizon.Enabled = False
End Sub

Private Sub PlaneetTekenen(RK As Double, delta As Double, radius As Double, Maxmag As Double, _
                         jdB As Double, jde As Double, _
                         Dpb As Long, Dpe As Long, _
                         MerkDagen As Boolean, MerkTekenDagen As Long, blnHorizon As Boolean)

Dim plNr As Long, k As Long
Dim epunt As Boolean
Dim JD As Double, t1 As Double
Dim rkp     As Double, decp As Double, magp As Double
Dim cRegel As String
Dim ddate As tDatum
    
    For plNr = Dpb To Dpe

      If plNr <> 3 Then
        k = 0

        Do While k <= Int(jde - jdB + 0.3)
            pgbVoortgang.Value = k / (jde - jdB + 1.3) * 100
            JD = jdB + k
            t1 = JDToT(JD)
            Call calcpos(plNr, t1, rkp, decp)
            If (MerkDagen) Then
                If k Mod MerkTekenDagen = 0 Then
                    Call tekpunt(RK, delta, radius, rkp, decp, Maxmag, 20, blnHorizon)
                Else
                    Call tekpunt(RK, delta, radius, rkp, decp, Maxmag, 10, blnHorizon)
                End If
            Else
                ddate = JDNaarKalender(JD)
                If Abs(ddate.DD - 1#) < 0.000001 Then
                     Call tekpunt(RK, delta, radius, rkp, decp, Maxmag, 20, blnHorizon)
                Else
                     Call tekpunt(RK, delta, radius, rkp, decp, Maxmag, 10, blnHorizon)
                End If
            End If
            k = k + 1
        Loop
      End If
    Next
End Sub

Sub InvoerTekenen(ByVal RK As Double, ByVal delta As Double, ByVal radius As Double, ByVal Maxmag As Double, ByVal MarkPerPositions As Double, optHorizon As Boolean)
Dim rkp     As Double, decp As Double
Dim cRegel As String
Dim nAantal As Long
    
    k = 0
    ctext = Me.txtInvoerPosities
    nPos = 1
    nPos = InStr(nPos, ctext, vbCrLf)
    Do While nPos > 0
        k = k + 1
        nPos = InStr(nPos + 1, ctext, vbCrLf)
    Loop
    nAantal = k
    
    nPos = 1
    nPos = InStr(nPos, ctext, ",")
    Do While nPos > 0
        ctext = Left(ctext, nPos - 1) + "." + Mid(ctext, nPos + 1)
        nPos = InStr(nPos, ctext, ",")
    Loop
    k = 0
    Do While k < nAantal
        nPos = 1
        pgbVoortgang.Value = k / nAantal * 100
        If InStr(nPos, ctext, vbCrLf) > 0 Then
            nPos = InStr(nPos, ctext, vbCrLf)
            sPositie = Left(ctext, nPos)
            ctext = Mid(ctext, nPos + 2)
        Else
            nPos = InStr(nPos, ctext, vbCr)
            sPositie = Left(ctext, nPos)
            ctext = Mid(ctext, nPos + 2)
        End If
        nPos = InStr(sPositie, "h")
        If nPos = 0 Then Exit Do
        Do While Mid(sPositie, nPos, 1) <> " "
            nPos = nPos - 1
            If nPos = 0 Then Exit Do
        Loop
        sPositie = Mid(sPositie, nPos + 1)
        
        srkp = Left(sPositie, InStr(sPositie, " "))
        sdecl = Mid(sPositie, InStr(sPositie, " ") + 1)
        
        nPos = InStr(srkp, "h")
        rkp = Val(Left(srkp, nPos - 1))
        srkp = Mid(srkp, nPos + 1)
        
        nPos = InStr(srkp, "m")
        If nPos > 0 Then
            rkp = rkp + Val(Left(srkp, nPos - 1)) / 60
            srkp = Mid(srkp, nPos + 1)
        End If
        
        nPos = InStr(srkp, "s")
        If nPos > 0 Then
            rkp = rkp + Val(Left(srkp, nPos - 1)) / 3600
            srkp = Mid(srkp, nPos + 1)
        End If
        
        nPos = InStr(sdecl, "°")
        decp = Val(Left(sdecl, nPos - 1))
        tekenDecp = sign(decp)
        decp = decp * tekenDecp
        sdecl = Mid(sdecl, nPos + 1)
        
        nPos = InStr(sdecl, "'")
        If nPos > 0 Then
            decp = decp + Val(Left(sdecl, nPos - 1)) / 60
            sdecl = Mid(sdecl, nPos + 1)
        End If
        
        nPos = InStr(sdecl, """")
        If nPos > 0 Then
            decp = decp + Val(Left(sdecl, nPos - 1)) / 3600
            sdecl = Mid(sdecl, nPos + 1)
        End If
        If MarkPerPositions = 0 Then MarkPerPositions = 1
        Debug.Print rkp & vbTab & decp
        If k Mod MarkPerPositions = 0 Then
            Call tekpunt(RK, delta, radius, rkp * Pi / 12, tekenDecp * decp * Pi / 180, Maxmag, 20, optHorizon)
        Else
            Call tekpunt(RK, delta, radius, rkp * Pi / 12, tekenDecp * decp * Pi / 180, Maxmag, 10, optHorizon)
        End If
        k = k + 1
    Loop
End Sub

Sub tekpunt(ByVal RK As Double, ByVal Dec As Double, ByVal r As Double, ByVal rkp As Double, ByVal decp As Double, ByVal mag As Double, ByVal r_planeet As Double, ByVal blnHorizon As Boolean)

Dim PI_2 As Double, PI_180 As Double, sxdec As Double, cxdec As Double
Dim Az As Double, hg As Double, x As Long, Y As Long, midx As Long, midy As Long
Dim Q As Double
Dim ster As tSter
Dim sRegel As String
Dim lrtf_code As Long
Dim aantal_sterren As Long
Dim Azt As Double

  midx = 19140 / schaalfactor: midy = 14595 / schaalfactor
  
  PI_2 = Pi / 2
  PI_180 = Pi / 180
  RK = RK * Pi / 12
  Dec = Dec * PI_180
  r = r * PI_180
  sxdec = Sin(Dec)
  cxdec = Cos(Dec)
  mag = mag * 10 '(* meegegeven magnitude is factor 10 te groot *)

If Not blnHorizon Then
    If (decp > Dec - r) And (decp < Dec + r) Then
          Q = RK - rkp
          hg = asin(Sin(decp) * sxdec + Cos(decp) * Cos(Q) * cxdec)
          If (r > PI_2 - hg) Then Az = atan2(Sin(Q), Cos(Q) * sxdec - tan(decp) * cxdec)
    
          hg = PI_2 - hg
          If (r > hg) Then
              x = Int(midx + midy * Sin(Az) * hg / r)
              Y = Int(midy + midy * Cos(Az) * hg / r)
              If (x > 0) And (Y > 0) And (x < 2 * midx) And (Y < 2 * midy) Then
                  Call print_rtf_circle(2, rtf_code(8), x, Y, r_planeet, RGB(255, 0, 0))
              End If
          End If
    End If
Else
    Call EquToHor(rkp, decp, RK, Dec, Az, hg)
    Azt = Az - Pi / 180 * cmbHorizon.ItemData(cmbHorizon.ListIndex) 'Az0
    If Azt > Pi Then Azt = Azt - 2 * Pi
    If Azt < -Pi Then Azt = Azt + 2 * Pi
    
    If (hg > 0) And (Abs(Azt) < PI_2) Then
        x = Int(midx + midy * Azt * Sqr(1 - hg / PI_2 * hg / PI_2) / PI_2)
        Y = Int(midy - midy * hg / PI_2)
    '                   Teken_ster x, Y, Straal(0.1 * ster.M, mag)
        Call print_rtf_circle(2, rtf_code(8), x, Y, r_planeet, RGB(255, 0, 0))
    End If
End If
  DoEvents
 End Sub

Private Sub optMerkPerDagen_Click()
If Val(txtMerkTekenDagen) = 0 Then
    txtMerkTekenDagen = "10"
End If
End Sub






