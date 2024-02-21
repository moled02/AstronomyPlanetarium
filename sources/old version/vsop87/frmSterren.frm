VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSterren 
   AutoRedraw      =   -1  'True
   Caption         =   "Stars"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11895
   Icon            =   "frmSterren.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   611
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   793
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picVolleMaan 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   2880
      Picture         =   "frmSterren.frx":030A
      ScaleHeight     =   217
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   54
      Top             =   480
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.PictureBox picCroppedMaan 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3315
      Left            =   6240
      ScaleHeight     =   225.074
      ScaleMode       =   0  'User
      ScaleWidth      =   221
      TabIndex        =   55
      Top             =   480
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   360
      TabIndex        =   52
      Top             =   8280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame fraMoon 
      Caption         =   "Moon"
      Height          =   1695
      Left            =   8880
      TabIndex        =   44
      Top             =   2760
      Width           =   2895
      Begin VB.TextBox txtStartPeriodeMoon 
         Height          =   285
         Left            =   1200
         TabIndex        =   25
         Top             =   240
         Width           =   1575
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   120
         TabIndex        =   45
         Top             =   960
         Width           =   2655
         Begin VB.OptionButton optMerkPerDagenMoon 
            Caption         =   "Mark per day"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.TextBox txtMerkTekenDagenMoon 
            Height          =   285
            Left            =   1680
            TabIndex        =   28
            Text            =   "1"
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.TextBox txtEindPeriodeMoon 
         Height          =   285
         Left            =   1200
         TabIndex        =   26
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Period start:"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Period end:"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.CheckBox chkEcliptica 
      Caption         =   "With ecliptica"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   3000
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.Frame fraInputPositions 
      Caption         =   "Input positions (for plotting path of asteroids or comets)"
      Height          =   3015
      Left            =   120
      TabIndex        =   39
      Top             =   4680
      Width           =   11655
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   1695
         Left            =   5520
         TabIndex        =   53
         Top             =   720
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   2990
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmSterren.frx":22BF8
      End
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
         TabIndex        =   41
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblInvoer 
         Caption         =   "Input positions (RA, Decl):"
         Height          =   615
         Left            =   3360
         TabIndex        =   40
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.ListBox lstPlaneten 
      Height          =   2085
      ItemData        =   "frmSterren.frx":22C7A
      Left            =   6360
      List            =   "frmSterren.frx":22C99
      Style           =   1  'Checkbox
      TabIndex        =   17
      Top             =   360
      Width           =   2175
   End
   Begin VB.ComboBox cmbHorizon 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmSterren.frx":22CE2
      Left            =   600
      List            =   "frmSterren.frx":22D30
      TabIndex        =   6
      Top             =   1920
      Width           =   1575
   End
   Begin VB.OptionButton optHorizon 
      Caption         =   "Chart of Horizon"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Frame fraKaart 
      Caption         =   "Chart"
      Enabled         =   0   'False
      Height          =   1335
      Left            =   2880
      TabIndex        =   35
      Top             =   1080
      Width           =   2895
      Begin VB.TextBox txtStraal 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtDeclinatie 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtRechteKlimming 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblStraal 
         Caption         =   "Radius:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblDeclinatie 
         Caption         =   "Declination:"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblRechteKlimming 
         Caption         =   "Right ascension:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CheckBox chkMetLijnen 
      Caption         =   "With lines"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   3720
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox chkMetBayer 
      Caption         =   "With Bayernumbers"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   3360
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.Frame fraPlaneten 
      Caption         =   "Planets"
      Height          =   2535
      Left            =   8880
      TabIndex        =   31
      Top             =   120
      Width           =   2895
      Begin VB.TextBox txtEindPeriode 
         Height          =   285
         Left            =   1200
         TabIndex        =   19
         Top             =   600
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   120
         TabIndex        =   42
         Top             =   840
         Width           =   2655
         Begin VB.TextBox txtMerkTekenDagen 
            Height          =   285
            Left            =   1680
            TabIndex        =   21
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optMerkPerMaand 
            Caption         =   "Mark per month"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   600
            Width           =   1815
         End
         Begin VB.OptionButton optMerkPerDagen 
            Caption         =   "Mark per day"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.TextBox txtStartPeriode 
         Height          =   285
         Left            =   1200
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   120
         TabIndex        =   43
         Top             =   1800
         Width           =   2655
         Begin VB.OptionButton optLines 
            Caption         =   "Lines"
            Height          =   375
            Left            =   1680
            TabIndex        =   24
            Top             =   150
            Width           =   735
         End
         Begin VB.OptionButton optDots 
            Caption         =   "Dots"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Label lblEindPeriode 
         Caption         =   "Period end:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblPeriode 
         Caption         =   "Period start:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdGenereren 
      Caption         =   "Generate"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   32
      Top             =   8160
      Width           =   1335
   End
   Begin VB.OptionButton optKaartje 
      Caption         =   "Small chart"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
   Begin VB.OptionButton optHuidigeSterrenhemel 
      Caption         =   "Current Sky"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.CheckBox chkMetPlaneten 
      Caption         =   "With planets"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   4080
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
      Top             =   8760
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame fraStarSchart 
      Caption         =   "Starchart"
      Height          =   2535
      Left            =   120
      TabIndex        =   48
      Top             =   0
      Width           =   5775
      Begin VB.Label lblGrensmagnitude 
         Caption         =   "Limiting magnitude"
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraOptionsChartchart 
      Caption         =   "Options starchart"
      Height          =   1695
      Left            =   120
      TabIndex        =   50
      Top             =   2760
      Width           =   5775
      Begin VB.TextBox txtGridDec 
         Height          =   285
         Left            =   4200
         TabIndex        =   16
         Text            =   "10"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtGridRk 
         Height          =   285
         Left            =   4200
         TabIndex        =   15
         Text            =   "60"
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox chkGrid 
         Caption         =   "With grid"
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   240
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Grid Dec (per º):"
         Height          =   255
         Left            =   2880
         TabIndex        =   57
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblGridRk 
         Caption         =   "Grid RK (per min):"
         Height          =   255
         Left            =   2880
         TabIndex        =   56
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame fraPlanets 
      Caption         =   "Selection Planets"
      Height          =   2535
      Left            =   6240
      TabIndex        =   51
      Top             =   120
      Width           =   2415
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   400
      Y1              =   176
      Y2              =   176
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   792
      Y1              =   528
      Y2              =   528
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   792
      Y1              =   304
      Y2              =   304
   End
   Begin VB.Line Line1 
      X1              =   400
      X2              =   400
      Y1              =   0
      Y2              =   304
   End
   Begin VB.Label lblgenerating 
      Caption         =   "Progress:"
      Height          =   255
      Left            =   4080
      TabIndex        =   0
      Top             =   8760
      Width           =   975
   End
End
Attribute VB_Name = "frmSterren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================== begin private uit configuratie bestand ===================
Private g_lijnkleur_grid As Long
Private g_Grote_cirkel_vulkleur As Long
'==================== einde private uit configuratie bestand ===================
Private Const MERGEPAINT = &HBB0226
Private Const SRCAND = &H8800C6
Private Const SRCCOPY = &HCC0020
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
' Private Declare Function RotateBitmap Lib "gdiplus" (ByVal hBitmapDC As Long, _
'      ByVal lWidth As Long, _
'      ByVal lHeight As Long, _
'  ByVal lRadians As Long)
      
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
lParam As Any) As Long
Private Const WM_PASTE = &H302


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
Private rtf_code(15) As String
Private Const schuif As Long = 1600
Private Const schaalfactor As Double = 2.6
Private objspecialfolder As New clsSpecialFolder
Private sTempName As String
Private nfile
Private nShplId As Long
Private Sub chkMetPlaneten_Click()
    fraPlaneten.Enabled = chkMetPlaneten
    lstPlaneten.Enabled = chkMetPlaneten
    fraMoon.Enabled = chkMetPlaneten
End Sub

Private Sub cmdGenereren_Click()
Dim sLatitude As String
Dim dat As tDatum
Dim dRechteKlimming As Double
Dim dDecinatie As Double
Dim dStraal As Double
Dim jdB As Double, jde As Double, JD_ZT As Double, JD_WT As Double
Dim JD0 As Double
Dim i As Long
Dim jdBM As Double, jdeM As Double
dat.jj = frmPlanets.Year
dat.mm = frmPlanets.MonthSelect.ListIndex + 1
dat.DD = frmPlanets.DaySelect
tt = (frmPlanets.Hrs + frmPlanets.Min / 60 + frmPlanets.Sec / 3600) / 24
dat.DD = dat.DD + tt
JD0 = KalenderNaarJD(dat)
Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Latitude", _
            sLatitude)
For i = 1 To UBound(rtf_code)
    rtf_code(i) = ""
Next

sTempName = objspecialfolder.TemporaryFolder + "\test_" + Format(Now(), "yyyy-mm-dd_hh.mm.ss") + ".rtf"

nShplId = 1025 'startwaarde
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
      Case Is = "[NIEUW-KLEUR-LIJN]"
          lrtf_code = 12
          sRegel = ""
      Case Is = "[TEKSTVAK-ZONDER-LIJN-TEKST]"
          lrtf_code = 13
          sRegel = ""
      Case Is = "[JUPITER-IMAGE]"
          lrtf_code = 14
          sRegel = ""
      Case Is = "[POLYGON]"
          lrtf_code = 15
          sRegel = ""
  End Select
  If sRegel <> "" Then rtf_code(lrtf_code) = rtf_code(lrtf_code) + sRegel + vbCrLf
Loop
Close #2
nfile = 2
Open sTempName For Output As #nfile
Print #nfile, rtf_code(5);

pgbVoortgang.value = 0
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
 '       Call TekenenGrid(24# * PlaatselijkeSterrentijd(dat), ReadDMS(sLatitude), 90, Val(txtGrensmagnitude))
    ElseIf Me.optKaartje Then
        dRechteKlimming = Val(txtRechteKlimming)
        dDecinatie = ReadDMS(txtDeclinatie)
        dStraal = Val(Me.txtStraal)
        Call stertek(dRechteKlimming, dDecinatie, dStraal, Val(txtGrensmagnitude))
'        Call TekenenGrid(dRechteKlimming, dDecinatie, dStraal, Val(txtGrensmagnitude))
    End If
End If

' {ecliptica}

If chkEcliptica Then
    JD = KalenderNaarJD(dat)
    'correctie voor zomertijd/wintertijd
    Call Zomertijd_Wintertijd(dat.jj, JD_ZT, JD_WT)
    If JD >= JD_ZT And JD < JD_WT Then
        JD = JD - 2 / 24
    Else
        JD = JD - 1 / 24
    End If

    'Bereken T (juliaanse eeuwen)
    
    Call tekEcliptica(dRechteKlimming, dDecinatie, dStraal, (JD - 2451545) / 36525, optHorizon)
End If
'   {ecliptica}

If chkMetPlaneten Then
    dat.DD = Val(Left(Me.txtStartPeriode, 2))
    dat.mm = Val(Mid(Me.txtStartPeriode, 4, 2))
    dat.jj = Val(Mid(Me.txtStartPeriode, 7, 4))
    jdB = KalenderNaarJD(dat)
    dat.DD = Val(Left(Me.txtEindPeriode, 2))
    dat.mm = Val(Mid(Me.txtEindPeriode, 4, 2))
    dat.jj = Val(Mid(Me.txtEindPeriode, 7, 4))
    jde = KalenderNaarJD(dat)
    
    dat.DD = Val(Left(Me.txtStartPeriodeMoon, 2))
    dat.mm = Val(Mid(Me.txtStartPeriodeMoon, 4, 2))
    dat.jj = Val(Mid(Me.txtStartPeriodeMoon, 7, 4))
    jdBM = KalenderNaarJD(dat) + tt
    dat.DD = Val(Left(Me.txtEindPeriodeMoon, 2))
    dat.mm = Val(Mid(Me.txtEindPeriodeMoon, 4, 2))
    dat.jj = Val(Mid(Me.txtEindPeriodeMoon, 7, 4))
    jdeM = KalenderNaarJD(dat) + tt
    
    For i = 1 To lstPlaneten.ListCount
        If lstPlaneten.Selected(i - 1) Then
            If lstPlaneten.ItemData(i - 1) > 0 Then
                Call PlaneetTekenen(dRechteKlimming, dDecinatie, dStraal, Val(txtGrensmagnitude), _
                                 jdB, jde, _
                                 lstPlaneten.ItemData(i - 1), lstPlaneten.ItemData(i - 1), _
                                 optMerkPerDagen, Val(txtMerkTekenDagen), optHorizon)
            Else
                Call MaanTekenen(dRechteKlimming, dDecinatie, dStraal, Val(txtGrensmagnitude), _
                                 jdBM, jdeM, _
                                 lstPlaneten.ItemData(i - 1), lstPlaneten.ItemData(i - 1), _
                                 Val(txtMerkTekenDagenMoon), optHorizon)
            End If
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
pgbVoortgang.value = 0
Exit Sub

word_open:
    If Err.Number = 462 Then 'Word waarschijnlijk gesloten
        Set g_word = New Application
        Resume
    End If
End Sub

Private Sub Lezen_configuratie()
Dim sRegel As String
Dim nfile As Long
    nfile = FreeFile()
    Open App.Path & "\Astronomie.cfg" For Input As #nfile
    Do While Not EOF(nfile)
        Line Input #nfile, sRegel
        nPosIs = InStr(sRegel, "=")
        If nPosIs > 0 Then
            sKeuze = Left(sRegel, nPosIs - 1)
            Select Case sKeuze
                Case Is = "LIJNKLEUR_GRID"
                    g_lijnkleur_grid = Val(Mid(sRegel, nPosIs + 1))
                Case Is = "GROTE_CIRKEL_VULKLEUR"
                    g_Grote_cirkel_vulkleur = Val(Mid(sRegel, nPosIs + 1))
            End Select
        End If
    Loop
    Close (nfile)
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
  Call print_rtf_circle(2, rtf_code(8), midx, midy, midy, g_Grote_cirkel_vulkleur) 'grote cirkel
'          Call TekenenGrid(dRechteKlimming, dDecinatie, dStraal, Val(txtGrensmagnitude))
  If Me.chkGrid Then
    Call TekenenGrid(RK, Dec, r, mag)
  End If
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
     pgbVoortgang.value = aantal_sterren / 50
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
  pgbVoortgang.value = 100
  DoEvents
 End Sub
 Sub tekEcliptica(ByVal RK As Double, ByVal Dec As Double, ByVal r As Double, ByVal T As Double, ByVal optHorizon As Boolean)

Dim PI_2 As Double, PI_180 As Double, sxdec As Double, cxdec As Double
Dim Az As Double, hg As Double, x As Long, Y As Long, midx As Long, midy As Long
Dim Q As Double
Dim ster As tSter
Dim sRegel As String
Dim lrtf_code As Long
Dim aantal_sterren As Long
Const nAantPunten = 50#
Dim posX1 As Double, posX2 As Double, posY1 As Double, posY2 As Double

  midx = 19140 / schaalfactor: midy = 14595 / schaalfactor
  
  PI_2 = Pi / 2
  PI_180 = Pi / 180
  sxdec = Sin(Dec)
  cxdec = Cos(Dec)
  mag = mag * 10 '(* meegegeven magnitude is factor 10 te groot *)

  eps = Obliquity(T)
  For k = 0 To 24 * nAantPunten - 1
    'Debug.Print eps * Sin(k / 500 * 12 / Pi)
        Call CalcXY(RK, Dec, r, k / nAantPunten * Pi / 12, eps * Sin(k / nAantPunten * Pi / 12), optHorizon, posX1, posY1)
        Call CalcXY(RK, Dec, r, (k + 1) / nAantPunten * Pi / 12, eps * Sin((k + 1) / nAantPunten * Pi / 12), optHorizon, posX2, posY2)
        If Not (posX1 = -1 Or posY1 = -1 Or posX2 = -1 Or posY2 = -1) Then
            Call print_rtf_lijn(2, rtf_code(6), posX1, posY1, posX2, posY2)
        End If
'          Call tekpunt(RK, Dec, r, k / nAantPunten * Pi / 12, eps * Sin(k / nAantPunten * Pi / 12), 10, 1, optHorizon)
  Next
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
    pgbVoortgang.value = aantal_sterren / 50
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
  pgbVoortgang.value = 100
  DoEvents
 End Sub

Function Straal(ByVal SterMag As Double, ByVal mag As Double) As Long
    Straal = 20 * (mag / 10 - SterMag)
End Function

Sub print_rtf_circle(nfile As Long, ByVal srtf_code As String, ByVal x As Long, ByVal Y As Long, ByVal r As Long, _
                     Optional ByVal Vul_kleur As Long = 0)
Dim nPos As Long
    nShplId = nShplId + 1
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
    nPos = InStr(srtf_code, "<SHPLID>")
    Do While InStr(srtf_code, "<SHPLID>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(nShplId, "0") + Mid(srtf_code, nPos + 8)
        nPos = InStr(srtf_code, "<SHPLID>")
    Loop
    Print #nfile, srtf_code;
End Sub
Sub print_rtf_maanschijf(nfile As Long, ByVal srtf_code As String, ByVal x As Long, ByVal Y As Long, ByVal r_maan As Long, _
                     ByVal PosAngle As Double, parAngle As Double, fase As Double, Optional ByVal Vul_kleur As Long = 0)
Dim nPos As Long
    nShplId = nShplId + 1
Dim i As Long
Dim XH As Long
    
    Const nStep = 1
    
'aanmaken van een maanschijf
   angle = PosAngle - parAngle '- Pi / 2
  
'   angle = modpi2(PosAngle - parAngle)
   'angle = modpi2(PosAngle)
   
'    angle = -Pi / 4
    Dim aPoints(360, 2) As Double
    For i = 0 To 360 Step nStep
        'test of deze aan de verlichte kant zit
        If i <= 180 Then
            PosX = Cos(i * Pi / 180): PosY = Sin(i * Pi / 180)
        Else
                If fase > 0.5 Then
                    PosX = Cos(i * Pi / 180): PosY = (2 * fase - 1) * Sin(i * Pi / 180)
                Else
                    PosX = Cos(i * Pi / 180): PosY = -(1 - 2 * fase) * Sin(i * Pi / 180)
                End If
        End If
        posX1 = PosX * Cos(angle) - PosY * Sin(angle)
        posY1 = PosX * Sin(angle) + PosY * Cos(angle)
        aPoints(i, 1) = Round(r_maan * posX1) + x - schuif
        aPoints(i, 2) = Round(r_maan * posY1) + Y + schuif
    Next
    
    Dim MidPoint(1, 2)   As Double
    MidPoint(1, 1) = (aPoints(0, 1) + aPoints(180, 1)) / 2
    MidPoint(1, 2) = (aPoints(0, 2) + aPoints(180, 2)) / 2
    
''     For i = 0 To 360 Step 2
'        aPoints(i, 1) = aPoints(i, 1) - MidPoint(1, 1) - schuif
'        aPoints(i, 2) = aPoints(i, 2) - MidPoint(1, 2) + schuif
'    Next
    
    Dim xMax As Double, yMax As Double
    xMax = -1: xmin = 999999
    yMax = -1: ymin = 999999
    For i = 0 To 360 Step nStep
        If aPoints(i, 1) > xMax Then xMax = aPoints(i, 1)
        If aPoints(i, 2) > yMax Then yMax = aPoints(i, 2)
        If aPoints(i, 1) < xmin Then xmin = aPoints(i, 1)
        If aPoints(i, 2) < ymin Then ymin = aPoints(i, 2)
    Next
    

    
'    x = x - schuif - MidPoint(1, 1)
'    Y = Y + schuif - MidPoint(1, 2)
    
    XH = xmin
    sC = "<LEFT>": nC = Len(sC)
    nPos = InStr(srtf_code, sC)
    Do While InStr(srtf_code, sC) > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(XH, "0") + Mid(srtf_code, nPos + nC)
        nPos = InStr(srtf_code, sC)
    Loop
    XH = ymin
    sC = "<TOP>": nC = Len(sC)
    nPos = InStr(srtf_code, sC)
    Do While InStr(srtf_code, sC) > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(XH, "0") + Mid(srtf_code, nPos + nC)
        nPos = InStr(srtf_code, sC)
    Loop
    
    XH = xMax
    sC = "<RIGHT>": nC = Len(sC)
    nPos = InStr(srtf_code, sC)
    Do While InStr(srtf_code, sC) > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(XH, "0") + Mid(srtf_code, nPos + nC)
        nPos = InStr(srtf_code, sC)
    Loop
    
    XH = yMax
    sC = "<BOTTOM>": nC = Len(sC)
    nPos = InStr(srtf_code, sC)
    Do While InStr(srtf_code, sC) > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(XH, "0") + Mid(srtf_code, nPos + nC)
        nPos = InStr(srtf_code, sC)
    Loop
    
    XH = xMax - xmin
    sC = "<GEO-RIGHT>": nC = Len(sC)
    nPos = InStr(srtf_code, sC)
    Do While InStr(srtf_code, sC) > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(XH, "0") + Mid(srtf_code, nPos + nC)
        nPos = InStr(srtf_code, sC)
    Loop
    sC = "<RIGHT-LEFT>": nC = Len(sC)
    nPos = InStr(srtf_code, sC)
    Do While InStr(srtf_code, sC) > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(XH, "0") + Mid(srtf_code, nPos + nC)
        nPos = InStr(srtf_code, sC)
    Loop
    
    XH = yMax - ymin
    sC = "<GEO-BOTTOM>": nC = Len(sC)
    nPos = InStr(srtf_code, sC)
    Do While InStr(srtf_code, sC) > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(XH, "0") + Mid(srtf_code, nPos + nC)
        nPos = InStr(srtf_code, sC)
    Loop
    sC = "<BOTTOM-TOP>": nC = Len(sC)
    nPos = InStr(srtf_code, sC)
    Do While InStr(srtf_code, sC) > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(XH, "0") + Mid(srtf_code, nPos + nC)
        nPos = InStr(srtf_code, sC)
    Loop
    
    XH = 360 / nStep + 1
    sC = "<VERT-PUNTEN>": nC = Len(sC)
    nPos = InStr(srtf_code, sC)
    Do While InStr(srtf_code, sC) > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(XH, "0") + Mid(srtf_code, nPos + nC)
        nPos = InStr(srtf_code, sC)
    Loop
    
    sARR_Punten = ""
    For XH = 0 To 360 Step nStep
        sARR_Punten = sARR_Punten + "(" + Format(aPoints(XH, 1) - xmin) + "," + _
                      Format(aPoints(XH, 2) - ymin) + ");"
    Next
    sARR_Punten = Left(sARR_Punten, Len(sARR_Punten) - 1)
    
    sC = "<VERT-ARRAY>": nC = Len(sC)
    nPos = InStr(srtf_code, sC)
    Do While InStr(srtf_code, sC) > 0
        srtf_code = Left(srtf_code, nPos - 1) + sARR_Punten + Mid(srtf_code, nPos + nC)
        nPos = InStr(srtf_code, sC)
    Loop
    
    XH = 2 * (360 / nStep + 1) + 2
    sC = "<VERT-PUNTEN*2+2>": nC = Len(sC)
    nPos = InStr(srtf_code, sC)
    Do While InStr(srtf_code, sC) > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(XH, "0") + Mid(srtf_code, nPos + nC)
        nPos = InStr(srtf_code, sC)
    Loop
    
    sARR_Segment = "16384;44032"
    For XH = 1 To 360 / nStep
        sARR_Segment = sARR_Segment + ";1;44032"
    Next
    sARR_Segment = sARR_Segment + ";24577;32768"
    sC = "<SEGMENT-ARRAY>": nC = Len(sC)
    nPos = InStr(srtf_code, sC)
    Do While InStr(srtf_code, sC) > 0
        srtf_code = Left(srtf_code, nPos - 1) + sARR_Segment + Mid(srtf_code, nPos + nC)
        nPos = InStr(srtf_code, sC)
    Loop
    
    XH = 360 / nStep
    sC = "<POLY-PUNTEN>": nC = Len(sC)
    nPos = InStr(srtf_code, sC)
    Do While InStr(srtf_code, sC) > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(XH, "0") + Mid(srtf_code, nPos + nC)
        nPos = InStr(srtf_code, sC)
    Loop
    
    sARR_Poly = ""
    For XH = 0 To 360 - 2 Step 2
        sARR_Poly = sARR_Poly + "\dpptx" + Format(aPoints(XH, 1) - xmin) + "\dppty" + _
                      Format(aPoints(XH, 2) - ymin)
    Next
    sC = "<POLY-ARRAY>": nC = Len(sC)
    nPos = InStr(srtf_code, sC)
    Do While InStr(srtf_code, sC) > 0
        srtf_code = Left(srtf_code, nPos - 1) + sARR_Poly + Mid(srtf_code, nPos + nC)
        nPos = InStr(srtf_code, sC)
    Loop

'    Me.picCroppedMaan.Visible = True
    Me.picVolleMaan.Visible = True
    Me.picCroppedMaan.Height = Me.picVolleMaan.Height
    Me.picCroppedMaan.Width = Me.picVolleMaan.Width
    If PosAngle > Pi Then
        Call RotateMoon(Me.picVolleMaan, Me.picVolleMaan, PosAngle - parAngle + Pi / 4)
    Else
        Call RotateMoon(Me.picVolleMaan, Me.picVolleMaan, -PosAngle - parAngle + Pi / 4)
    End If
    Me.picCroppedMaan.ScaleMode = 1
    Me.picCroppedMaan.Width = Int((xMax - xmin) / (r_maan * 2) * 217)
    Me.picCroppedMaan.Height = Int((yMax - ymin) / (r_maan * 2) * 217)
'    Me.picCroppedMaan.Width = picCroppedMaan.Width * (xMax - xmin) / (yMax - ymin)
    Me.picCroppedMaan.Cls
    Dim xMin1 As Long, yMin1 As Long
    xMin1 = (xmin - x + schuif + r_maan) / (r_maan * 2) * 217
    yMin1 = (ymin - Y - schuif + r_maan) / (r_maan * 2) * 217
    
    BitBlt Me.picCroppedMaan.hdc, _
        0, 0, Me.picCroppedMaan.Width, Me.picCroppedMaan.Height, _
        Me.picVolleMaan.hdc, xMin1, yMin1, SRCAND

'    Select Case modpi2(-angle)
'    Case Is > 3 * Pi / 2
'        BitBlt Me.picCroppedMaan.hdc, _
'        0, 0, Me.picCroppedMaan.Width, Me.picCroppedMaan.Height, _
'        Me.picVolleMaan.hdc, (217 - Me.picCroppedMaan.Width), (217 - Me.picCroppedMaan.Height), SRCAND
'    Case Is > Pi
'        BitBlt Me.picCroppedMaan.hdc, _
'        0, 0, Me.picCroppedMaan.Width, Me.picCroppedMaan.Height, _
'        Me.picVolleMaan.hdc, 0, (217 - Me.picCroppedMaan.Height), SRCAND
'    Case Is > Pi / 2
'        BitBlt Me.picCroppedMaan.hdc, _
'        0, 0, Me.picCroppedMaan.Width, Me.picCroppedMaan.Height, _
'        Me.picVolleMaan.hdc, 0, 0, SRCAND
'    Case Else
'        BitBlt Me.picCroppedMaan.hdc, _
'        0, 0, Me.picCroppedMaan.Width, Me.picCroppedMaan.Height, _
'        Me.picVolleMaan.hdc, (217 - Me.picCroppedMaan.Width), 0, SRCAND
'    End Select
'    If angle > Pi Then
'        BitBlt Me.picCroppedMaan.hdc, _
'        0, 0, Me.picCroppedMaan.Width, Me.picCroppedMaan.Height, _
'        Me.picVolleMaan.hdc, (217 - Me.picCroppedMaan.Width), 0, SRCAND
'        BitBlt Me.picCroppedMaan.hdc, _
'        0, 0, Me.picCroppedMaan.Width, Me.picCroppedMaan.Height, _
'        Me.picVolleMaan.hdc, (217 - Me.picCroppedMaan.Width), (217 - Me.picCroppedMaan.Height), SRCAND
'    Else
'        BitBlt Me.picCroppedMaan.hdc, _
'        0, 0, Me.picCroppedMaan.Width, Me.picCroppedMaan.Height, _
'        Me.picVolleMaan.hdc, 0, 0, SRCAND
'    End If
    picCroppedMaan.Refresh
    Call SavePicture(picCroppedMaan.Image, "c:\mijn.bmp")
    'Clipboard.SetData Screen.ActiveControl.Picture
    Me.RichTextBox1.TextRTF = ""
    Call InsertPictureInRichTextBox(Me.RichTextBox1, LoadPicture("c:\mijn.bmp"))
    Me.picVolleMaan.Visible = False

'    Debug.Print Me.RichTextBox1.TextRTF
    Dim sString As String
    sString = Me.RichTextBox1.TextRTF
    'Debug.Print Left(sString, 1000)
    Me.RichTextBox1.Text = sString
    
    Dim str As String
    str = Me.RichTextBox1.Text
    nPos = InStr(str, "{\pict")
    str = Mid(str, nPos)
    str = Mid(str, 1, Len(str) - 8)
    sC = "<MAAN-PICT>": nC = Len(sC)
    nPos = InStr(srtf_code, sC)
    Do While InStr(srtf_code, sC) > 0
        srtf_code = Left(srtf_code, nPos - 1) + str + Mid(srtf_code, nPos + nC)
        nPos = InStr(srtf_code, sC)
    Loop
    
    nPos = InStr(srtf_code, "<SHPLID>")
    Do While InStr(srtf_code, "<SHPLID>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(nShplId, "0") + Mid(srtf_code, nPos + 8)
        nPos = InStr(srtf_code, "<SHPLID>")
    Loop
    Print #nfile, srtf_code;
End Sub
Sub print_rtf_boog(nfile As Long, ByVal srtf_code As String, ByVal x As Long, ByVal Y As Long, ByVal r As Long, ByVal FlipH As Long, FlipV As Long)
Dim nPos As Long
    nShplId = nShplId + 1
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
    nPos = InStr(srtf_code, "<SHPLID>")
    Do While InStr(srtf_code, "<SHPLID>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(nShplId, "0") + Mid(srtf_code, nPos + 8)
        nPos = InStr(srtf_code, "<SHPLID>")
    Loop
    Print #nfile, srtf_code;
End Sub
    
Sub print_rtf_lijn(nfile As Long, ByVal srtf_code As String, ByVal x As Long, ByVal Y As Long, ByVal x1 As Long, ByVal y1 As Long, Optional ByVal kleur As Long = 0)
Dim nPos As Long
Dim links As Long, top As Long
Dim rechts As Long, onder As Long
Dim FlipV As Long, FlipH As Long
Dim tx1 As Long, tx2 As Long, ty1 As Long, ty2 As Long
    
nShplId = nShplId + 1

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
    nPos = InStr(srtf_code, "<SHPLID>")
    Do While InStr(srtf_code, "<SHPLID>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(nShplId, "0") + Mid(srtf_code, nPos + 8)
        nPos = InStr(srtf_code, "<SHPLID>")
    Loop
    nPos = InStr(srtf_code, "<LINE-COLOR>")
    Do While InStr(srtf_code, "<LINE-COLOR>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(kleur, "0") + Mid(srtf_code, nPos + 12)
        nPos = InStr(srtf_code, "<LINE-COLOR>")
    Loop
    Print #nfile, srtf_code;
End Sub
Sub print_rtf_textbox(nfile As Long, ByVal srtf_code As String, ByVal x As Long, ByVal Y As Long, ByVal stext As String)
Dim nPos As Long
Const nGroot As Long = 244
Dim sBayerPos2 As String
    
    nShplId = nShplId + 1
    x = x - schuif - nGroot / 2 + 110: Y = Y + schuif + nGroot / 2 - 50
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
    nPos = InStr(srtf_code, "<SHPLID>")
    Do While InStr(srtf_code, "<SHPLID>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(nShplId, "0") + Mid(srtf_code, nPos + 8)
        nPos = InStr(srtf_code, "<SHPLID>")
    Loop
    sBayerPos2 = Mid(stext, 2, 1)
    If sBayerPos2 = "0" Then sBayerPos2 = vbNullString
    nPos = InStr(srtf_code, "<SUPER>")
    Do While InStr(srtf_code, "<SUPER>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + sBayerPos2 + Mid(srtf_code, nPos + 7)
        If sBayerPos2 <> "" Then
        x = x
        End If
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
    pgbVoortgang.value = aantal_lijnen * LenB(lijn) / LOF(1) * 100
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
    pgbVoortgang.value = aantal_lijnen * LenB(lijn) / LOF(1) * 100
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
ElseIf Planet = 9 Then 'Pluto, alhoewel achterhaald planeet
    Call PlanetPosHi(0, T, sAarde)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call SphToRect(SGeo, TAarde)
    Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
    ' Call Reduction2000(0, RA, Decl)
    'coordinaten omzetten naar J2000
    Call PrecessFK5(T, 0#, RA, Decl)
End If
End Sub
Private Sub calcposMaan(ByVal Planet As Long, ByVal T As Double, RA As Double, Decl As Double, r As Double, ByRef PosAngle As Double, ByRef parAngle As Double, ByRef fase As Double)
Dim Obl As Double
Dim NutLon As Double
Dim NutObl As Double
Dim SHelio As TSVECTOR, SGeo As TSVECTOR, SSun As TSVECTOR
    'Q1 = SHelio, Q2 = SGeo
Dim sAarde As TSVECTOR
Dim TAarde As TVECTOR
Dim sZon As TSVECTOR
Dim TPluto As TVECTOR
Dim Dist As Double
Dim sLatitude As String, sLongitude As String
Dim MoonPhysData As TMOONPHYSDATA
Dim deltaT As Double
Dim JD_ZT As Double, JD_WT As Double

Dim LAST As Double, ObsLat As Double, ObsLon As Double, Height As Double
Dim RhoCosPhi As Double, RhoSinPhi As Double
Dim sMoon As TSVECTOR

Call Zomertijd_Wintertijd(frmPlanets.Year, JD_ZT, JD_WT)

Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Latitude", _
            sLatitude)
ObsLat = Val(sLatitude) * Pi / 180
Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Longitude", _
            sLongitude)
ObsLon = Val(sLongitude) * Pi / 180
deltaT = ApproxDeltaT(T)
T = T + TijdCorrectie(TToJD(T) + 0.2, JD_ZT, JD_WT) / 36525# + deltaT * secToT
LAST = SiderealTime(T) + NutLon * Cos(Obl) - ObsLon
Call ObserverCoord(ObsLat, Height, RhoCosPhi, RhoSinPhi)

    'berekenen posities voor epoch 2000
Obl = Obliquity(T)
Call NutationConst(T, NutLon, NutObl)

Call Lune(TToJD(T), RA, Decl, Dist, dkm, diam, phase, illum)
Call Lune(TToJD(T - Dist * LightTimeConst), RA, Decl, Dist, dkm, diam, phase, illum)

RA = RA * Pi / 12
Decl = Decl * Pi / 180
'coordinaten zijn voor J2000. Omzetten naar huidige dag, en daarna appearent berekenen
Call PrecessFK5(0, T, RA, Decl)

r = diam / 3600
'coordinaten zijn voor J2000. Omzetten naar huidige dag, en daarna appearent berekenen
Call Nutation(NutLon, NutObl, Obl, RA, Decl)

Call EquToEcl(RA, Decl, Obl, SGeo.l, SGeo.B)
SGeo.r = sMoon.r
SHelio.l = 0: SHelio.B = 0: SHelio.r = 0
Call PlanetPosHi(0, T, sAarde, True)
Call HelioToGeo(SHelio, sAarde, SSun)
Call PlanetPosHi(0, T - SSun.r * LightTimeConst, sAarde, True)
Call HelioToGeo(SHelio, sAarde, SSun)
Call MoonPhysEphemeris(T, SGeo, SSun, Obl, NutLon, NutObl, MoonPhysData)
parAngle = ParallacticAngle(RA, Decl, ObsLat, LAST)
PosAngle = MoonPhysData.x
fase = MoonPhysData.k
End Sub

Private Sub command1_click()
    Me.picCroppedMaan.Height = Me.picVolleMaan.Height
    Me.picCroppedMaan.Width = Me.picVolleMaan.Width
'    Call bmp_rotate(picVolleMaan, Me.picCroppedMaan, 90)
'    Me.picCroppedMaan.Cls
'    BitBlt Me.picCroppedMaan.hDC, _
'    0, 0, 500, 500, _
'    Me.picVolleMaan.hDC, 0, 0, SRCAND
    picCroppedMaan.Refresh
    'RotateBitmap picCroppedMaan, 500, 500, 0.2
    Call RotateMoon(Me.picVolleMaan, Me.picCroppedMaan, 0.5)
'    Call RotatePicture(Me.picVolleMaan, Me.picCroppedMaan, 0.5)
'    Call SavePicture(picCroppedMaan.Image, "c:\mijn.bmp")
'    Clipboard.SetData Screen.ActiveControl.Picture
    'RotateBitmap picCroppedMaan.hDC, 500, 500, 0.2
'    Call InsertPictureInRichTextBox(Me.RichTextBox1, LoadPicture("c:\mijn.bmp"))
'    Debug.Print Me.RichTextBox1.TextRTF
'    Dim sString As String
'    sString = Me.RichTextBox1.TextRTF
'    Debug.Print Left(sString, 1000)
'    Me.RichTextBox1.Text = sString
End Sub

Sub InsertPictureInRichTextBox(RTB As RichTextBox, Picture As StdPicture)
' copy into the clipboard
' Copy the picture into the clipboard.
Clipboard.Clear
Clipboard.SetData Picture
' paste into the RichTextBox control
SendMessage RTB.hwnd, WM_PASTE, 0, 0
End Sub

Private Sub Form_Load()
Dim JD As Double
Dim dat As tDatum
 '=============configuratie
Call Lezen_configuratie
'=============configuratie

dat.DD = 1
dat.mm = frmPlanets.MonthSelect.ListIndex + 1
dat.jj = frmPlanets.Year
txtStartPeriode = "01-" & Format(dat.mm, "00") & "-" & Format(frmPlanets.Year, "0000")
dat = JDNaarKalender(KalenderNaarJD(dat) + 40)
dat.DD = 1
'dat = JDNaarKalender(KalenderNaarJD(dat) - 1)
txtEindPeriode = Format(dat.DD, "00") & "-" & Format(dat.mm, "00") & "-" & Format(dat.jj, "0000")

dat.DD = frmPlanets.DaySelect.ListIndex + 1
dat.mm = frmPlanets.MonthSelect.ListIndex + 1
dat.jj = frmPlanets.Year

txtStartPeriodeMoon = Format(dat.DD, "00") & "-" & Format(dat.mm, "00") & "-" & Format(frmPlanets.Year, "0000")
txtEindPeriodeMoon = Format(dat.DD, "00") & "-" & Format(dat.mm, "00") & "-" & Format(frmPlanets.Year, "0000")
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
lstPlaneten.Selected(8) = True
lstPlaneten.ListIndex = -1

End Sub

Private Sub optHorizon_Click()
    cmbHorizon.Enabled = True
    fraKaart.Enabled = False
End Sub

Private Sub optHuidigeSterrenhemel_Click()
Dim JD As Double
Dim dat As tDatum
fraKaart.Enabled = optKaartje.value
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
fraKaart.Enabled = optKaartje.value

txtStartPeriode = "01-01-" & Format(frmPlanets.Year, "0000")
txtEindPeriode = "01-01-" & Format(frmPlanets.Year + 1, "0000")
cmbHorizon.Enabled = False
End Sub

Private Sub PlaneetTekenen(RK As Double, delta As Double, radius As Double, maxmag As Double, _
                         jdB As Double, jde As Double, _
                         Dpb As Long, Dpe As Long, _
                         MerkDagen As Boolean, MerkTekenDagen As Long, blnHorizon As Boolean)

Dim plNr As Long, k As Long
Dim epunt As Boolean
Dim JD As Double, T1 As Double
Dim rkp     As Double, decp As Double, magp As Double
Dim cRegel As String
Dim ddate As tDatum
    
    For plNr = Dpb To Dpe

      If plNr <> 3 Then
        k = 0

        Do While k <= Int(jde - jdB + 0.3)
            pgbVoortgang.value = k / (jde - jdB + 1.3) * 100
            JD = jdB + k
            T1 = JDToT(JD)
            Call calcpos(plNr, T1, rkp, decp)
            If (MerkDagen) Then
                If k Mod MerkTekenDagen = 0 Then
                    Call tekpunt(RK, delta, radius, rkp, decp, maxmag, 20, blnHorizon)
                Else
                    Call tekpunt(RK, delta, radius, rkp, decp, maxmag, 10, blnHorizon)
                End If
            Else
                ddate = JDNaarKalender(JD)
                If Abs(ddate.DD - 1#) < 0.000001 Then
                     Call tekpunt(RK, delta, radius, rkp, decp, maxmag, 20, blnHorizon)
                Else
                     Call tekpunt(RK, delta, radius, rkp, decp, maxmag, 10, blnHorizon)
                End If
            End If
            k = k + 1
        Loop
      End If
    Next
End Sub
Private Sub MaanTekenen(RK As Double, delta As Double, radius As Double, maxmag As Double, _
                         jdBM As Double, jdeM As Double, _
                         Dpb As Long, Dpe As Long, _
                         MerkTekenDagenMoon As Double, blnHorizon As Boolean)

Dim plNr As Long, k As Long
Dim epunt As Boolean
Dim T1 As Double
Dim rkp     As Double, decp As Double, magp As Double, semidiam As Double, r As Double
Dim cRegel As String
Dim ddate As tDatum
Dim PosAngle As Double, parAngle As Double, fase As Double
Dim JD As Double
Dim AantalStappen As Long

    AantalStappen = Int((jdeM - jdBM) / MerkTekenDagenMoon) + 1
    k = 0

    Do While k < AantalStappen
        pgbVoortgang.value = k / AantalStappen * 100
        JD = jdBM + MerkTekenDagenMoon * k
        T1 = JDToT(JD)
        Call calcposMaan(plNr, T1, rkp, decp, r, PosAngle, parAngle, fase)
        If radius = 90 Then 'tbv overzicht kaarten wordt de Maan iets te groot getekend om zodoende de fase goed te kunnen zien
            Call tekMaanSchijf(RK, delta, radius, rkp, decp, maxmag, 2 * r / radius * 14595 / schaalfactor, blnHorizon, PosAngle, parAngle, fase)
        Else
            Call tekMaanSchijf(RK, delta, radius, rkp, decp, maxmag, r / radius * 14595 / schaalfactor, blnHorizon, PosAngle, parAngle, fase)
        End If
        k = k + 1
    Loop
    pgbVoortgang.value = 100
End Sub




Sub InvoerTekenen(ByVal RK As Double, ByVal delta As Double, ByVal radius As Double, ByVal maxmag As Double, ByVal MarkPerPositions As Double, optHorizon As Boolean)
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
        pgbVoortgang.value = k / nAantal * 100
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
            Call tekpunt(RK, delta, radius, rkp * Pi / 12, tekenDecp * decp * Pi / 180, maxmag, 20, optHorizon)
        Else
            Call tekpunt(RK, delta, radius, rkp * Pi / 12, tekenDecp * decp * Pi / 180, maxmag, 10, optHorizon)
        End If
        k = k + 1
    Loop
End Sub
Sub tekMaanSchijf(ByVal RK As Double, ByVal Dec As Double, ByVal r As Double, ByVal rkp As Double, ByVal decp As Double, ByVal mag As Double, ByVal r_maan As Double, ByVal blnHorizon As Boolean, ByVal PosAngle As Double, ByVal parAngle As Double, fase As Double)

Dim PI_2 As Double, PI_180 As Double, sxdec As Double, cxdec As Double
Dim Az As Double, hg As Double, x As Long, Y As Long, midx As Long, midy As Long
Dim Q As Double
Dim ster As tSter
Dim sRegel As String
Dim lrtf_code As Long
Dim aantal_sterren As Long
Dim Azt As Double
Dim radius As Double
  midx = 19140 / schaalfactor: midy = 14595 / schaalfactor
  
  PI_2 = Pi / 2
  PI_180 = Pi / 180
  RK = RK * Pi / 12
  Dec = Dec * PI_180
  radius = r
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
                If radius = 90 Then
                    Call print_rtf_maanschijf(2, rtf_code(15), x, Y, r_maan, PosAngle, parAngle, fase, RGB(255, 0, 0))
                Else
                    Call print_rtf_maanschijf(2, rtf_code(15), x, Y, r_maan, PosAngle, parAngle, fase, RGB(255, 0, 0))
                End If
                  'Call print_rtf_circle(2, rtf_code(8), x, Y, 5, RGB(255, 0, 0))
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
        Call print_rtf_maanschijf(2, rtf_code(15), x, Y, r_maan, PosAngle + Az, parAngle, fase, RGB(255, 0, 0))
'        Call print_rtf_maanschijf(2, rtf_code(15), x, Y, r_maan, PosAngle + Az, parAngle, fase, RGB(255, 0, 0))
        'Call print_rtf_circle(2, rtf_code(8), x, Y, 5, RGB(255, 0, 0))
    End If
End If
  DoEvents
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
Sub CalcXY(ByVal RK As Double, ByVal Dec As Double, ByVal r As Double, ByVal rkp As Double, ByVal decp As Double, ByVal blnHorizon As Boolean, ByRef PosX As Double, ByRef PosY As Double)

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

PosX = -1: PosY = -1
If Not blnHorizon Then
    If (decp >= Dec - r) And (decp <= Dec + r) Then
          Q = RK - rkp
          hg = asin(Sin(decp) * sxdec + Cos(decp) * Cos(Q) * cxdec)
          If (r >= PI_2 - hg) Then Az = atan2(Sin(Q), Cos(Q) * sxdec - tan(decp) * cxdec)
    
          hg = PI_2 - hg
          If (r >= hg) Then
              x = Int(midx + midy * Sin(Az) * hg / r)
              Y = Int(midy + midy * Cos(Az) * hg / r)
              If (x >= 0) And (Y >= 0) And (x <= 2 * midx) And (Y <= 2 * midy) Then
                PosX = x: PosY = Y
'                  Call print_rtf_circle(2, rtf_code(8), x, Y, r_planeet, RGB(255, 0, 0))
              End If
          End If
    End If
Else
    Call EquToHor(rkp, decp, RK, Dec, Az, hg)
    Azt = Az - Pi / 180 * cmbHorizon.ItemData(cmbHorizon.ListIndex) 'Az0
    If Azt > Pi Then Azt = Azt - 2 * Pi
    If Azt < -Pi Then Azt = Azt + 2 * Pi
    
    If (hg >= 0) And (Abs(Azt) <= PI_2) Then
        x = Int(midx + midy * Azt * Sqr(1 - hg / PI_2 * hg / PI_2) / PI_2)
        Y = Int(midy - midy * hg / PI_2)
    '                   Teken_ster x, Y, Straal(0.1 * ster.M, mag)
        PosX = x: PosY = Y
        'Call print_rtf_circle(2, rtf_code(8), x, Y, r_planeet, RGB(255, 0, 0))
    End If
End If
  DoEvents
 End Sub
Private Sub optMerkPerDagen_Click()
If Val(txtMerkTekenDagen) = 0 Then
    txtMerkTekenDagen = "10"
End If
End Sub




     Sub bmp_rotate(pic1 As PictureBox, pic2 As PictureBox, ByVal Theta!)
       Const Pi = 3.14159265359
       Dim c1x As Integer  ' Center of pic1.
       Dim c1y As Integer  '   "
       Dim c2x As Integer  ' Center of pic2.
       Dim c2y As Integer  '   "
       Dim A As Single     ' Angle of c2 to p2.
       Dim r As Integer    ' Radius from c2 to p2.
       Dim p1x As Integer  ' Position on pic1.
       Dim p1y As Integer  '   "
       Dim p2x As Integer  ' Position on pic2.
       Dim p2y As Integer  '   "
       Dim n As Integer    ' Max width or height of pic2.

       ' Compute the centers.
       c1x = pic1.ScaleWidth / 2
       c1y = pic1.ScaleHeight / 2
       c2x = pic2.ScaleWidth / 2
       c2y = pic2.ScaleHeight / 2

       ' Compute the image size.
       n = pic2.ScaleWidth
       If n < pic2.ScaleHeight Then n = pic2.ScaleHeight
       n = n / 2 - 1
       ' For each pixel position on pic2.
       For p2x = 0 To n
          For p2y = 0 To n
             ' Compute polar coordinate of p2.
             If p2x = 0 Then
               A = Pi / 2
             Else
               A = Atn(p2y / p2x)
             End If
             r = Sqr(1& * p2x * p2x + 1& * p2y * p2y)

             ' Compute rotated position of p1.
             p1x = r * Cos(A + Theta)
             p1y = r * Sin(A + Theta)

             ' Copy pixels, 4 quadrants at once.
             c0& = pic1.Point(c1x + p1x, c1y + p1y)
             c1& = pic1.Point(c1x - p1x, c1y - p1y)
             c2& = pic1.Point(c1x + p1y, c1y - p1x)
             c3& = pic1.Point(c1x - p1y, c1y + p1x)
             If c0& <> -1 Then pic2.PSet (c2x + p2x, c2y + p2y), c0&
             If c1& <> -1 Then pic2.PSet (c2x - p2x, c2y - p2y), c1&
             If c2& <> -1 Then pic2.PSet (c2x + p2y, c2y - p2x), c2&
             If c3& <> -1 Then pic2.PSet (c2x - p2y, c2y + p2x), c3&
          Next
          ' Allow pending Windows messages to be processed.
          DoEvents
       Next
     End Sub

' Rotate fr_pic by a multiple of 90 degrees
' and place the result in to_pic. Both PictureBoxes
' should have AutoRedraw = True.
Public Sub RotateMoon(fr_pic As PictureBox, to_pic As PictureBox, ByVal angle As Double)
Dim fr_pixels() As RGBTriplet
Dim c0 As RGBTriplet
Dim c1 As RGBTriplet
Dim c2 As RGBTriplet
Dim c3 As RGBTriplet

Dim to_pixels() As RGBTriplet
Dim bits_per_pixel As Integer
Dim fr_wid As Long
Dim fr_hgt As Long
Dim to_wid As Long
Dim to_hgt As Long
Dim x As Integer
Dim Y As Integer
Dim A As Double, r As Double
Dim p1x As Long, p1y As Long
Dim x1 As Integer, y1 As Integer

Const Pi = 3.1415926536
    ' Get the picture's image.
    GetBitmapPixels fr_pic, fr_pixels, bits_per_pixel

    ' Get the picture's size.
    fr_wid = UBound(fr_pixels, 1) + 1
    fr_hgt = UBound(fr_pixels, 2) + 1
    If angle = 0 Or angle = 180 Then
        to_wid = fr_wid
        to_hgt = fr_hgt
    Else
        to_wid = fr_hgt
        to_hgt = fr_wid
    End If

    ' Size the output picture to fit.
    to_pic.Width = to_pic.Parent.ScaleX(to_wid, vbPixels, to_pic.Parent.ScaleMode) + _
        to_pic.Width - to_pic.ScaleWidth
    to_pic.Height = to_pic.Parent.ScaleY(to_hgt, vbPixels, to_pic.Parent.ScaleMode) + _
        to_pic.Height - to_pic.ScaleHeight

    to_pic.Cls
    Dim cx As Long
    Dim cy As Long
    Dim px As Long
    Dim py As Long
    cx = to_wid / 2
    cy = to_hgt / 2
    
    ' Copy the rotated pixels.
    ReDim to_pixels(0 To to_wid - 1, 0 To to_hgt - 1)
    For x = 0 To fr_wid - 1
        For Y = 0 To fr_hgt - 1
            to_pixels(x, Y) = fr_pixels(1, 1)
        Next
    Next

             Dim c1x As Integer  ' Center of pic1.
            Dim c1y As Integer  '   "
            Dim c2x As Integer  ' Center of pic2.
            Dim c2y As Integer  '   "
            Dim p2x As Integer  ' Position on pic2.
            Dim p2y As Integer  '   "
            Dim n As Integer    ' Max width or height of pic2.

       
       ' Compute the centers.
       c1x = fr_pic.ScaleWidth / 2
       
       c1y = fr_pic.ScaleHeight / 2
       c2x = to_pic.ScaleWidth / 2
       c2y = to_pic.ScaleHeight / 2

       ' Compute the image size.
       n = to_pic.ScaleWidth
       If n < to_pic.ScaleHeight Then n = to_pic.ScaleHeight
       n = n / 2 - 1

            For p2x = 0 To fr_wid - 1
 
                For p2y = 0 To fr_hgt - 1
                 ' Compute polar coordinate of p2.
                 If p2x = 0 Then
                   A = Pi / 2
                 Else
                   A = Atn(p2y / p2x)
                 End If
                 r = Sqr(1& * p2x * p2x + 1& * p2y * p2y)
    
                 ' Compute rotated position of p1.
                 p1x = r * Cos(A + angle)
                 p1y = r * Sin(A + angle)
    
                 ' Copy pixels, 4 quadrants at once.
                 On Error Resume Next
                 c0 = fr_pixels(c1x + p1x, c1y + p1y)
                  c1 = fr_pixels(c1x - p1x, c1y - p1y)
                 c2 = fr_pixels(c1x + p1y, c1y - p1x)
                 c3 = fr_pixels(c1x - p1y, c1y + p1x)
                 to_pixels(c2x + p2x, c2y + p2y) = c0
                 to_pixels(c2x - p2x, c2y - p2y) = c1
                 to_pixels(c2x + p2y, c2y - p2x) = c2
                 to_pixels(c2x - p2y, c2y + p2x) = c3
              Next
            Next
   

    ' Display the result.
    SetBitmapPixels to_pic, bits_per_pixel, to_pixels

    ' Make the image permanent.
    to_pic.Refresh
    to_pic.Picture = to_pic.Image
End Sub

' ======================================  TEKENEN GRID ====================================

Public Sub BepaalGrenzenDecl(ByVal hg As Double, ByVal nb As Double, ByVal u As Double, ByRef d1 As Double, ByRef d2 As Double)

'{formule :  sin(h) = sin(p)*sin(d)+cos(p)*cos(d)*cos(u)
'           A = Sin(H) + B * Cos(H)
' bekenden u, H, P
'}
Dim Az As Double, A As Double, B As Double
Dim sind_1 As Double, sind1 As Double, sind2 As Double
Dim T As Double

A = Sin(hg) / Sin(nb)
B = Cos(u) * Cos(nb) / Sin(nb)
sind_1 = Sqr(A * A - (1 + B * B) * (A * A - B * B))

sind1 = (A + sind_1) / (1 + B * B)
d1 = asin(sind1)

sind2 = (A - sind_1) / (1 + B * B)
d2 = asin(sind2)
'{d2 = Arctan(Sind2 / (Sqrt(1 - Sind2 * Sind2)))}
If d2 < d1 Then
   T = d1
   d1 = d2
   d2 = T
End If
End Sub

Public Sub BepaalGrenzenU(ByVal hg As Double, ByVal nb As Double, ByVal D As Double, ByRef u1 As Double, ByRef u2 As Double)
'{formule :  sin(h) = sin(p)*sin(d)+cos(p)*cos(d)*cos(u)
'           a = Sin(H) + b * Cos(H)
' bekenden u, H, P
'}
Dim u As Double, cosu As Double, T As Double

cosu = (Sin(hg) - Sin(nb) * Sin(D)) / (Cos(nb) * Cos(D))
If (cosu = 0) Then
     u1 = 0
     u2 = 2 * Pi
ElseIf Abs(cosu) > 1 Then
     u1 = 0
     u2 = 2 * Pi
Else
     u = acos(cosu)
     If u > 2 * Pi Then u = u - 2 * Pi
     If u < 0 Then u = u + 2 * Pi
     If u < Pi Then
          u2 = u
          u1 = -u2
     Else
          u1 = u
          u2 = 2 * Pi - u1
     End If
End If
If u2 < u1 Then
   u2 = u2 + 2 * Pi
End If
End Sub

Public Sub BepaalGrenzen(ByVal x As Double, ByVal dx As Double, ByRef g1 As Double, ByRef g2 As Double)

'   IF (x-dx>=g1) AND (x+dx<=g2) THEN
'    BEGIN
'        g1 = x - dx
'        g2 = x + dx
'    End
'    ELSE IF (x-dx<g1) THEN
'        g2 = x + dx
'    Else
'        g1=x-dx}
    g1 = x - dx
    g2 = x + dx
End Sub

Public Sub TekenenGrid(ByVal GSTL As Double, ByVal nb As Double, ByVal r As Double, ByVal maxmag As Double)
'(24# * PlaatselijkeSterrentijd(dat), ReadDMS(sLatitude), 90, Val(txtGrensmagnitude))
Dim g1a As Double, g2a As Double, g1b As Double, g2b As Double
Dim d1 As Double, d2 As Double
Dim u1 As Double, u2 As Double
Dim GridRKstap As Double
Dim GridDecstap As Double
Dim posX1 As Double, posY1 As Double
Dim posX2 As Double, posY2 As Double

  GSTL = GSTL * Pi / 12
  If Abs(nb) <= 0.01 Then nb = sign(nb) * 0.01
  nb = nb * Pi / 180
GridRKstap = Int(r / 18)
If GridRKstap = 0 Then GridRKstap = 1
GridDecstap = Int(r / 90)
If GridDecstap = 0 Then GridDecstap = 1

'==============
GridRk = Val(Me.txtGridRk) 'elk uur een lijn
gridDec = Val(Me.txtGridDec) 'om de 10 graden lijn
'================================
Dim dp As String
dp = "D"

g2a = 24 * 60
g1a = 0
i = g1a
delta = nb * 180 / Pi
    
g1a = 0
g2a = 24
g1b = -80
g2b = 80
    
' ================== teken declinatie cirkels ==================
    If dp = "D" Then
          delta = nb * 180 / Pi
          Call BepaalGrenzen(delta, 90, g1b, g2b)
          If g2b > 80 Then g2b = 80
          If g1b < -80 Then g1b = -80
          g2a = g2a * 60
      Else
'{          BepaalGrenzen(RK,2*radius/15,g1a,g2a)}
          g1a = g1a * 60
          g2a = g2a * 60
'{          BepaalGrenzen(Delta,90,g1b,g2b) }
          g1b = -90
          g2b = 90
          If g2b > 90 Then g2b = 90
          If g1b < -90 Then g1b = -90
      End If
      stapA = (g2a - g1a) * 0.05
      stapb = (g2b - g1b) * 0.1
 '=====================================
    j = Int(g1b / gridDec) * gridDec
    stapb = gridDec '(d2 - d1) * 0.1
    Do While j <= g2b
          If dp = "D" Then
              Call BepaalGrenzenU(0, nb, j * Pi / 180, u1, u2)
          Else
              u1 = 0
              u2 = 2 * Pi
          End If
        
              u1 = u1 * 12 / Pi * 60
          u2 = u2 * 12 / Pi * 60
      i = Int(u1 / GridRKstap) * GridRKstap
      Do While i <= u2
        
       ' If dp = "D" Then
            Call CalcXY(GSTL * 12 / Pi, delta, r, GSTL - i / 60 * Pi / 12, j * Pi / 180, False, posX1, posY1)
            Call CalcXY(GSTL * 12 / Pi, delta, r, GSTL - (i + GridRKstap) / 60 * Pi / 12, j * Pi / 180, False, posX2, posY2)
            If Not (posX1 = -1 Or posY1 = -1 Or posX2 = -1 Or posY2 = -1) Then
                Call print_rtf_lijn(2, rtf_code(6), posX1, posY1, posX2, posY2, g_lijnkleur_grid)
            End If
    
        ' Else
        '    Call tekpunt(RK, delta, radius, I / 60 * Pi / 12, j * Pi / 180, 5, 10, False)
        ' End If
         i = i + GridRKstap
        Loop
         j = j + stapb
    Loop

' ================== teken uur (rechte klimming) cirkels ==================
    u1 = 0
    u2 = 2 * Pi
    u1 = u1 * 12 / Pi * 60
    u2 = u2 * 12 / Pi * 60
    i = Int(u1 / GridRk) * GridRk
    stapA = GridRk
    Do While i < u2
           If dp = "D" Then
               Call BepaalGrenzenDecl(0, nb, GSTL - i * Pi / 12 / 60, d1, d2)
               If Sin(GSTL - i * Pi / 12 / 60 - Pi / 2) <= 0 Then
                   d1 = d1 * 180 / Pi
               Else
                   d1 = d2 * 180 / Pi
                End If
               d2 = 90
           Else
               d1 = -90
               d2 = 90
           End If
           j = d1
           stapb = GridDecstap '(d2 - d1) * 0.1
           Do While j < d2
                Call CalcXY(GSTL * 12 / Pi, delta, r, i / 60 * Pi / 12, j * Pi / 180, False, posX1, posY1)
                j1 = j + stapb
                If j1 > 90 Then j1 = 90
                Call CalcXY(GSTL * 12 / Pi, delta, r, i / 60 * Pi / 12, (j1) * Pi / 180, False, posX2, posY2)
                If Not (posX1 = -1 Or posY1 = -1 Or posX2 = -1 Or posY2 = -1) Then
                    Call print_rtf_lijn(2, rtf_code(6), posX1, posY1, posX2, posY2, g_lijnkleur_grid)
                End If
             j = j + stapb
          Loop
          j1 = j + stapb
          If j1 > 90 Then j1 = 90
          Call CalcXY(GSTL * 12 / Pi, delta, r, i / 60 * Pi / 12, j1 * Pi / 180, False, posX1, posY1)
          Call CalcXY(GSTL * 12 / Pi, delta, r, i / 60 * Pi / 12, d2 * Pi / 180, False, posX2, posY2)
          If Not (posX1 = -1 Or posY1 = -1 Or posX2 = -1 Or posY2 = -1) Then
            Call print_rtf_lijn(2, rtf_code(6), posX1, posY1, posX2, posY2, g_lijnkleur_grid)
          End If
          i = i + stapA
    Loop
      
'===================================
      
      
End Sub

