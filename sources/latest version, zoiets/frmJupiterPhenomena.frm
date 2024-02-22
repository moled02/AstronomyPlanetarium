VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmJupiterPhenomena 
   Caption         =   "Phenomena of Jupitermoons"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9390
   Icon            =   "frmJupiterPhenomena.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   9390
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox txtWordMacro 
      Height          =   3495
      Left            =   240
      TabIndex        =   4
      Top             =   4680
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6165
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmJupiterPhenomena.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdToonGegevens 
      Caption         =   "Show data"
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdNaarWord 
      Caption         =   "Show in Word"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   3960
      Width           =   1935
   End
   Begin MSComctlLib.ProgressBar PBVoortgang 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   8280
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtVerschijnselen 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   8895
   End
End
Attribute VB_Name = "frmJupiterPhenomena"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const JUP_NORMAAL = 0
Private Const JUP_BEDEKT = 1
Private Const JUP_SCHADUW = 2
Private Const JUP_VERDUISTERD = 3

Private Const IO = 1
Private Const Europa = 2
Private Const GANYMEDES = 3
Private Const Callisto = 4
Private Const Schaal = 0.2  ' {straal Jupiter}

Private rMaan As Variant
Private oMaan As Variant
Private dMaan As Variant
Private rMaanJup As Variant

Private JD0 As Double, jde As Double, JD_ZT As Double, JD_WT As Double
Private ep           As Double
Private ObsLon As Double, ObsLat As Double
Private blnDoorgaan As Boolean

Private rtf_code(14) As String
Private Const schuif As Long = 1600
Private Const schaalfactor As Double = 2.6
Private objspecialfolder As New clsSpecialFolder
Private sTempName As String
Private nfile
Private blnZwart As Boolean
Private nShplId As Long
Private Sub SchrijfVak2(ByRef stext As String, vector As TVECTOR, regel As Long, maan As Long)
stext = stext + "TekenTekstVak2 """
stext = stext + FormatX(10 - Schaal * vector.x - 0.15, "##0.00", True) + ""","""
'    { 0.30 / 2 = 0.15 ---> 0.30 is getal wat in word-macro staat }
stext = stext + FormatX(1 + regel + Schaal * vector.Y - 0.5, "##0.00", True) + ""","""
'    write(uitvoer,1 + regel + Schaal*vector[2] - 0.50:6:2);
'{       write(uitvoer,1 + regel -0.50:6:2);
'        hiermee worden de tekstvak op een horizontale regel geplaatst }
stext = stext + FormatX(maan, "0") + """" + vbCrLf
End Sub

Private Sub SchrijfCirkel(ByRef stext As String, vector As TVECTOR, r1 As Double, r2 As Double, kleur As Long, regel As Long, maan As Long)
stext = stext + "TekenCirkel """
stext = stext + FormatX(10 - Schaal * vector.x - r1, "##0.00", True) + ""","""
stext = stext + FormatX(1 + regel + Schaal * vector.Y - r2, "##0.00", True) + ""","""
stext = stext + FormatX(r1 * 2, "#0.00", True) + ""","""
stext = stext + FormatX(r2 * 2, "#0.00", True) + ""","""
stext = stext + FormatX(kleur, "0", True) + """" + vbCrLf
If maan > 0 Then Call SchrijfVak2(stext, vector, regel, maan)
End Sub

Private Sub SchrijfTekstVak(ByRef stext As String, x As Double, Y As Double, B As Double, H As Double, T As String)
stext = stext + "TekenTekstVak """
stext = stext + FormatX(x - B / 2, "##0.00", True) + ""","""
stext = stext + FormatX(1 + Y - H / 2, "##0.00", True) + ""","""
stext = stext + FormatX(B, "#0.00", True) + ""","""
stext = stext + FormatX(H, "#0.00", True) + ""","""
stext = stext + T + """" + vbCrLf
End Sub

Private Sub cmdNaarWord_Click()
Dim objspecialfolder As New clsSpecialFolder
Dim stext As String
Dim sString As String
Dim sTmp As String
Dim nPos As Long
Dim ddate As tDatum
Dim i As Long, nAantVbCrLf As Long, nVbCrLf As Long
Dim JD As Double
Dim rJup As Double
Dim vMaan As TVECTOR
Dim vMaanShadow As TVECTOR
Dim T As Double
Dim regel As Long
Dim j As Long
Dim x As Double
Dim sTempFile As String
Dim nfile As Long
cmdNaarWord.Enabled = False
For j = 1 To UBound(rtf_code()): rtf_code(j) = "": Next
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
  End Select
  If sRegel <> "" Then rtf_code(lrtf_code) = rtf_code(lrtf_code) + sRegel + vbCrLf
Loop
Close #2
nfile = 2


nfile = FreeFile
sTempFile = objspecialfolder.TemporaryFolder + "\jupmanen_" + Format(Now(), "yyyy-mm-dd_hh.mm.ss") + ".rtf"
Open sTempFile For Output As #nfile
Print #nfile, rtf_code(5);

blnZwart = MsgBox("Do you want a black background?", vbYesNo, "Question") = vbYes
If blnZwart Then
    Print #nfile, "{\shp{\*\shpinst\shpleft100\shptop400\shpright11000\shpbottom15859\shpfhdr0\shpbxcolumn\shpbypara\shpwr3\shpwrk0\shpfblwtxt0\shpz0\shplid1026{\sp{\sn shapeType}{\sv 1}}{\sp{\sn fFlipH}{\sv 0}}{\sp{\sn fFlipV}{\sv 0}}" & _
                  "{\sp{\sn fLockAspectRatio}{\sv 0}}{\sp{\sn fillColor}{\sv 0}}{\sp{\sn fFilled}{\sv 1}}{\sp{\sn fLine}{\sv 1}}}{\shprslt{\*\do\dobxcolumn\dobypara\dodhgt8192\dprect\dpx100\dpy400\dpxsize10900\dpysize158194" & _
                  "\dpfillfgcr255\dpfillfgcg255\dpfillfgcb255\dpfillbgcr0\dpfillbgcg0\dpfillbgcb0\dpfillpat1\dplinew15\dplinecor0\dplinecog0\dplinecob0}}}"
End If
'    Print #nfile, "AchterGrond=Zwart"
'Else
'    Print #nfile, "AchterGrond=Wit"
'End If

'Print #nfile, "AchtergrondTekstVak" + vbCrLf

sString = txtVerschijnselen.Text
nPos = -1
nAantVbCrLf = 0
nPos = InStr(sString, vbCrLf)
Do While nPos <> 0
    nAantVbCrLf = nAantVbCrLf + 1
    nPos = InStr(nPos + 1, sString, vbCrLf)
Loop
regel = -1
pagina = 1
totregel = 1
aantsub = 0
blnDoorgaan = True
Do While (sString <> vbNullString Or nPos = 0) And blnDoorgaan
    nVbCrLf = nVbCrLf + 1
    DoEvents
    stext = ""
    regel = regel + 1
    totregel = totregel + 1
    nPos = InStr(sString, vbCrLf)
    sTmp = Left(sString, nPos - 1)
    ddate.DD = Val(Left(sTmp, 2))
    ddate.MM = Val(Mid(sTmp, 4, 2))
    ddate.jj = Val(Mid(sTmp, 7, 4))
    
    ddate.DD = ddate.DD + Val(Mid(sTmp, 12, 2)) / 24 + Val(Mid(sTmp, 15, 2)) / 1440 '+ Val(Mid(sTmp, 18, 2)) / 86400
    Call Zomertijd_Wintertijd(ddate.jj, JD_ZT, JD_WT)
    '{JD_ZT/JD_WT zijn berekend voor 0h UT}
    JD_ZT = JD_ZT + 2 / 24 '{= 3h WT}
    JD_WT = JD_WT + 1 / 24 '{= 3h ZT}
    JD = KalenderNaarJD(ddate) + TijdCorrectie(JD, JD_ZT, JD_WT)
    vMaan.x = 0: vMaan.Y = 0: vMaan.Z = 0
    Call SchrijfCirkel(stext, vMaan, Schaal, Schaal / 1.071374, 8, regel, 0)
    Call print_rtf_circle(nfile, rtf_code(2), 8200 - 105 * vMaan.x, regel * 567 + 105 * vMaan.Y, 105, JUP_NORMAAL, 0)

    For i = 1 To 4
        Call BerekenPositieMaan(i, JD, False, vMaan, rJup)
        vMaan.Y = vMaan.Y * 1.071374
        Call BerekenPositieMaan(i, JD, True, vMaanShadow, rJup)
        vMaanShadow.Y = vMaanShadow.Y * 1.071374
        If (vMaan.Z > 0) And (Abs((vMaan.x * vMaan.x + vMaan.Y * vMaan.Y)) < 1) Then
    '        {bedekt}
            Call SchrijfCirkel(stext, vMaan, Schaal * 0.2, Schaal * 0.2, JUP_BEDEKT, regel, i)
            Call print_rtf_circle(nfile, rtf_code(2), 8200 - 105 * vMaan.x, regel * 567 + 105 * vMaan.Y, 20, JUP_BEDEKT, i)
        Else   '{niet bedekt}
            If (vMaan.Z > 0) And (Abs((vMaanShadow.x * vMaanShadow.x + vMaanShadow.Y * vMaanShadow.Y)) < 1) Then
                '{verduisterd (schaduw jupiter op maantje)}
                Call SchrijfCirkel(stext, vMaan, Schaal * 0.2, Schaal * 0.2, JUP_VERDUISTERD, regel, i)
                Call print_rtf_circle(nfile, rtf_code(2), 8200 - 105 * vMaan.x, regel * 567 + 105 * vMaan.Y, 23, JUP_VERDUISTERD, i)
            Else  '{niet bedekt en niet verduisterd}
                '{zichtbaar}
                Call SchrijfCirkel(stext, vMaan, Schaal * 0.2, Schaal * 0.2, JUP_NORMAAL, regel, i)
                Call print_rtf_circle(nfile, rtf_code(2), 8200 - 105 * vMaan.x, regel * 567 + 105 * vMaan.Y, 23, JUP_NORMAAL, i)
            End If
        End If
        If (vMaan.Z < 0) And (Abs((vMaanShadow.x * vMaanShadow.x + vMaanShadow.Y * vMaanShadow.Y)) < 1) Then
            '{schaduw op jupiter}
            Call SchrijfCirkel(stext, vMaanShadow, Schaal * 0.2, Schaal * 0.2, JUP_SCHADUW, regel, 0)
            Call print_rtf_circle(nfile, rtf_code(2), 8200 - 105 * vMaanShadow.x, regel * 567 + 105 * vMaanShadow.Y, 15, JUP_SCHADUW, 0)
        End If
    Next
    x = regel
    If blnZwart Then
        Call print_rtf_textbox(nfile, rtf_code(13), 2200, 567 * regel, 1871, 312, StrDate(ddate) + " " + StrHMS(Frac(ddate.DD) * Pi2, 2), "8")
    Else
        Call print_rtf_textbox(nfile, rtf_code(10), 2200, 567 * regel, 1871, 312, StrDate(ddate) + " " + StrHMS(Frac(ddate.DD) * Pi2, 2), "8")
    End If
    'Call SchrijfTekstVak(stext, 1, X, 3.3, 0.4, StrDate(ddate) + " " + StrHMS(Frac(ddate.DD) * Pi2, 2))
    sString = Mid(sString, nPos + 2)
    If regel > 22 Then
'        stext = stext + "NieuwePagina" + vbCrLf
'        stext = stext + "AchtergrondTekstVak" + vbCrLf
'        pagina = pagina + 1
        Print #nfile, rtf_code(11)
        If blnZwart Then
            Print #nfile, "{\shp{\*\shpinst\shpleft100\shptop400\shpright11000\shpbottom15859\shpfhdr0\shpbxcolumn\shpbypara\shpwr3\shpwrk0\shpfblwtxt0\shpz0\shplid1026{\sp{\sn shapeType}{\sv 1}}{\sp{\sn fFlipH}{\sv 0}}{\sp{\sn fFlipV}{\sv 0}}" & _
                          "{\sp{\sn fLockAspectRatio}{\sv 0}}{\sp{\sn fillColor}{\sv 0}}{\sp{\sn fFilled}{\sv 1}}{\sp{\sn fLine}{\sv 1}}}{\shprslt{\*\do\dobxcolumn\dobypara\dodhgt8192\dprect\dpx100\dpy400\dpxsize10900\dpysize158194" & _
                          "\dpfillfgcr255\dpfillfgcg255\dpfillfgcb255\dpfillbgcr0\dpfillbgcg0\dpfillbgcb0\dpfillpat1\dplinew15\dplinecor0\dplinecog0\dplinecob0}}}"
        End If
        regel = -1
    End If
    'Print #nfile, stext
'    If totregel > 100 Then
'        aantsub = aantsub + 1
'        totregel = 0
'        If aantsub < 10 Then
'            sText = sText + "Call x" + Format(aantsub, "0") + vbCrLf
'        Else
'            sText = sText + "Call x" + Format(aantsub, "#0") + vbCrLf
'        End If
'        sText = sText + "End Sub" + vbCrLf
'        If aantsub < 10 Then
'           sText = sText + "Sub x" + Format(aantsub, "0") + vbCrLf
'        Else
'           sText = sText + "Sub x" + Format(aantsub, "#0")
'        End If
'    End If
    PBVoortgang.value = nVbCrLf / nAantVbCrLf * 100
Loop
'If regel > 0 Then
'    sText = sText + "End Sub" + vbCrLf
'End If
'nfile = FreeFile
'Open App.Path + "\rtf_wordcode.txt" For Binary As #nfile
'sTmp = Space(LOF(nfile))
'Get #nfile, , sTmp
'sText = sText + sTmp

Print #nfile, rtf_code(4); 'het einde toevoegen
Close (nfile)
'Me.txtWordMacro = sText
On Error GoTo word_open
g_word.Visible = True
g_word.Documents.Open FileName:=sTempFile, ConfirmConversions:=False
' Shell "Winword " & Chr(34) & sTempName & Chr(34), vbNormalFocus
g_word.Activate
cmdNaarWord.Enabled = True
Exit Sub

word_open:
    If Err.Number = 462 Then 'Word waarschijnlijk gesloten
        Set g_word = New Application
        Resume
    End If
    cmdNaarWord.Enabled = True
End Sub

Private Sub cmdToonGegevens_Click()
Dim sString As String
Dim stext As String
Dim sTmp As String
Dim nPos As Long
Dim ddate As tDatum
Dim i As Long
Dim JD As Double
Dim rJup As Double
Dim vMaan As TVECTOR
Dim MoonName As Variant
Dim T As Double
Me.cmdToonGegevens.Enabled = False
MoonName = Array("", "Io       ", "Europa   ", "Ganymedes", "Callisto ")
sString = txtVerschijnselen.Text
nPos = -1
nAantVbCrLf = 0
nPos = InStr(sString, vbCrLf)
Do While nPos <> 0
    nAantVbCrLf = nAantVbCrLf + 1
    nPos = InStr(nPos + 1, sString, vbCrLf)
Loop
blnDoorgaan = True
Do While (sString <> vbNullString Or nPos = 0) And blnDoorgaan
    nVbCrLf = nVbCrLf + 1
    DoEvents
    nPos = InStr(sString, vbCrLf)
    sTmp = Left(sString, nPos - 1)
    ddate.DD = Val(Left(sTmp, 2))
    ddate.MM = Val(Mid(sTmp, 4, 2))
    ddate.jj = Val(Mid(sTmp, 7, 4))
    
    ddate.DD = ddate.DD + Val(Mid(sTmp, 12, 2)) / 24 + Val(Mid(sTmp, 15, 2)) / 1440 + Val(Mid(sTmp, 18, 2)) / 86400
    Call Zomertijd_Wintertijd(ddate.jj, JD_ZT, JD_WT)
    '{JD_ZT/JD_WT zijn berekend voor 0h UT}
    JD_ZT = JD_ZT + 2 / 24 '{= 3h WT}
    JD_WT = JD_WT + 1 / 24 '{= 3h ZT}
    JD = KalenderNaarJD(ddate) + TijdCorrectie(JD, JD_ZT, JD_WT)
    
    stext = stext + "Ephemeris for Jupitermoons on "
    stext = stext + StrDate(ddate) + " at " + StrHMS(Frac(ddate.DD) * Pi2, 4) + vbCrLf
    
    T = JDToT(KalenderNaarJD(ddate))
    stext = stext + "JD = " + FormatX(JD, "######0.000") + ", DeltaT : " + FormatX(ApproxDeltaT(T), "##0.0")
    stext = stext + vbCrLf + vbCrLf

    For i = 1 To 4
        Call BerekenPositieMaan(i, JD, False, vMaan, rJup)
        stext = stext + "M " + MoonName(i) + " " + FormatX(vMaan.x, "##0.00000") + " " + _
                        FormatX(vMaan.Y, "##0.00000") + vbTab + FormatX(vMaan.Y, "##0.00000") + " "
        
        vMaan.Y = vMaan.Y * 1.071374
        stext = stext + " " + FormatX(Sqr(vMaan.x * vMaan.x + vMaan.Y * vMaan.Y), "##0.00000") + vbCrLf
        
        Call BerekenPositieMaan(i, JD, True, vMaan, rJup)
        stext = stext + "S " + MoonName(i) + " " + FormatX(vMaan.x, "##0.00000") + " " + _
                        FormatX(vMaan.Y, "##0.00000") + vbTab + FormatX(vMaan.Y, "##0.00000") + vbTab
        
        vMaan.Y = vMaan.Y * 1.071374
        stext = stext + " " + FormatX(Sqr(vMaan.x * vMaan.x + vMaan.Y * vMaan.Y), "##0.00000") + vbCrLf
    Next
    stext = stext + "----------------------------------------------------------" + vbCrLf
    sString = Mid(sString, nPos + 2)
    PBVoortgang.value = nVbCrLf / nAantVbCrLf * 100
Loop
Me.txtWordMacro = stext
Me.cmdToonGegevens.Enabled = True
End Sub
Private Function FormatX(expression, fformat, Optional blnMetPunt As Boolean = False)
Dim sFormat As String
    sFormat = Format(expression, fformat)
    sFormat = Space(Len(fformat) - Len(sFormat)) + sFormat
    If blnMetPunt Then
        nPos = InStr(sFormat, ",")
        Do While nPos > 0
            sFormat = Left(sFormat, nPos - 1) + "." + Mid(sFormat, nPos + 1)
            nPos = InStr(sFormat, ",")
        Loop
    End If
    FormatX = sFormat
End Function

Private Sub Form_Activate()
      Me.cmdNaarWord.Enabled = False
      Me.cmdToonGegevens.Enabled = False
      Call bereken
      Me.cmdNaarWord.Enabled = True
      Me.cmdToonGegevens.Enabled = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim stext As String
    If KeyCode = 67 And Shift = 2 Then
       If ActiveControl.Name = "txtWordMacro" Then
          stext = Me.txtWordMacro.Text
       Else
          stext = Me.txtVerschijnselen
       End If
       Clipboard.Clear
       Clipboard.SetText (stext)
    ElseIf KeyCode = 17 And Shift = 0 Then
        stext = Me.txtWordMacro.Text
       Clipboard.Clear
       Clipboard.SetText (stext)
    End If
    Debug.Print KeyCode & Shift
End Sub

Private Sub Form_Load()
      rMaan = Array(0, 5.8995412, 9.3865048, 14.972724, 26.334835)
      oMaan = Array(0, 203.4058643, 101.2916334, 50.2345179, 21.4879801)
      dMaan = Array(0, 1.769860477, 3.5504094133, 7.166387079, 16.753552373)
      rMaanJup = Array(0, 0.025471381, 0.021890561, 0.036843283, 0.033612152)   ' {straal jupitermaan in jupstralen}
#If FRANS Then
    Me.Caption = "Phénomène de satellite de Jupiter"
    Me.cmdNaarWord.Caption = "Montrer in Word"
    Me.cmdToonGegevens.Caption = "Montrer données"
#End If
End Sub


Private Sub ov(ByVal Planet As Long, ByRef ddate As tDatum, _
             ByVal ObsLon As Double, ByVal ObsLat As Double, ByVal TimeZone As Double, ByVal Height As Double, ByVal StartHeight As Double, _
             ByRef Opk As Double, ByRef Ond As Double)

Dim sAarde As TSVECTOR, SHelio As TSVECTOR, SGeo As TSVECTOR
Dim Obl As Double
Dim T0 As Double, RA As Double, Decl As Double, RA1 As Double, Decl1 As Double, RA2 As Double, Decl2 As Double
Dim deltaT As Double
Dim RTS As tRiseSetTran
Dim sLatitude As String, sLongitude As String

Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Latitude", _
            sLatitude)
ObsLat = Val(sLatitude) * Pi / 180
Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Longitude", _
            sLongitude)
ObsLon = Val(sLongitude) * Pi / 180
    
T0 = JDToT(KalenderNaarJD(ddate))
Obl = Obliquity(T0)
If Planet = 0 Then 'Zon
    Call PlanetPosHi(0, T0 - 1 / 36525, sAarde, True)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call EclToEqu(SGeo.l, SGeo.B, Obl, RA1, Decl1)
    
    Call PlanetPosHi(0, T0, sAarde, True)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
    
    Call PlanetPosHi(0, T0 + 1 / 36525, sAarde, True)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call EclToEqu(SGeo.l, SGeo.B, Obl, RA2, Decl2)
    
    Call RiseSet(T0, deltaT, RA1, Decl1, RA, Decl, RA2, Decl2, Height, ObsLon, ObsLat, RTS)
Else 'Jupiter
    Call PlanetPosHi(0, T0 - 1 / 36525, sAarde, True)
    Call PlanetPosHi(5, T0 - 1 / 36525, SHelio, True)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call PlanetPosHi(5, T0 - 1 / 36525 - SGeo.r * LightTimeConst, SHelio, True)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call EclToEqu(SGeo.l, SGeo.B, Obl, RA1, Decl1)
    
    Call PlanetPosHi(0, T0, sAarde, True)
    Call PlanetPosHi(5, T0, SHelio, True)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call PlanetPosHi(5, T0 - SGeo.r * LightTimeConst, SHelio, True)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
    
    Call PlanetPosHi(0, T0 + 1 / 36525, sAarde, True)
    Call PlanetPosHi(5, T0 + 1 / 36525, SHelio, True)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call PlanetPosHi(5, T0 + 1 / 36525 - SGeo.r * LightTimeConst, SHelio, True)
    Call HelioToGeo(SHelio, sAarde, SGeo)
    Call EclToEqu(SGeo.l, SGeo.B, Obl, RA2, Decl2)
    
    Call RiseSet(T0, deltaT, RA1, Decl1, RA, Decl, RA2, Decl2, Height, ObsLon, ObsLat, RTS)
    
End If
Opk = RTS.Rise * 12 / Pi
Ond = RTS.Setting * 12 / Pi
End Sub
             

Function Zichtbaar(JD As Double) As Boolean
Dim ddate As tDatum
Dim Opk As Double, Ond As Double
Dim tijd As Double
Dim bTmp As Boolean
Dim hoogte As Double

    tijd = Frac(JD + 0.5) * 24
    ddate = JDNaarKalender(JD)
    ddate.DD = Int(ddate.DD)
    hoogte = 0
    
    Call ov(0, ddate, ObsLon, ObsLat, 0, hoogte, -6 * DToR, Opk, Ond)
    If Opk <= 0 Then
        Opk = 0
    End If
    If Ond <= 0 Then
        Ond = 24
    End If
    bTmp = (tijd < Opk) Or (tijd > Ond)
    Call ov(5, ddate, ObsLon, ObsLat, 0, hoogte, 5 * DToR, Opk, Ond)
    If Ond > Opk Then
      bTmp = (bTmp) And (tijd >= Opk) And (tijd <= Ond)
    Else
      bTmp = (bTmp) And ((tijd <= Ond) Or (tijd >= Opk))
    End If
    Zichtbaar = bTmp
End Function

Private Sub BerekenPositieMaan(maan As Long, JD As Double, bShadow As Boolean, _
                             ByRef vMaan As TVECTOR, ByRef rJup As Double)

Dim T As Double, deltaT As Double, DtofUT As Double, Obl As Double
Dim NutLon As Double, NutObl As Double, TimeZone As Double
Dim SHelio As TSVECTOR, ShelioJ As TSVECTOR, SNiks As TSVECTOR, SEarth As TSVECTOR, SGeo As TSVECTOR
Dim i As Long
Dim vDummy As TVECTOR, vdummy2 As TVECTOR
Dim nArg As Double

T = JDToT(JD)
deltaT = ApproxDeltaT(T)
DtofUT = T + deltaT * secToT

Obl = Obliquity(T)

'{ Main Calculations }
Call PlanetPosHi(5, T, SHelio, True)
Call PlanetPosHi(EPDATE, T, SEarth, True)
Call HelioToGeo(SHelio, SEarth, SGeo)

'{ Do just one light time iteration }

For i = 1 To 1
    Call PlanetPosHi(5, DtofUT - SGeo.r * LightTimeConst, SHelio, True)
    Call HelioToGeo(SHelio, SEarth, SGeo)
Next
'{---------------------------------------------------------------------------}
  T = DtofUT - SGeo.r * LightTimeConst
  If Not bShadow Then
      Call JSatEclipticPosition(DUMMY_SATELLITE, T, vDummy)
      Call JSatViewFrom(DUMMY_SATELLITE, vDummy, SGeo, vDummy, False, vDummy, True)
      Call JSatEclipticPosition(maan, T, vMaan)
      Call JSatViewFrom(maan, vMaan, SGeo, vDummy, False, vMaan, True)
      rJup = 1
  Else
      Call JSatEclipticPosition(DUMMY_SATELLITE, T, vdummy2)
      Call JSatViewFrom(DUMMY_SATELLITE, vdummy2, SHelio, vdummy2, False, vdummy2, False)
      Call JSatEclipticPosition(maan, T, vMaan)
      Call JSatViewFrom(maan, vMaan, SHelio, vdummy2, False, vMaan, False)
      rJup = 1
      If vMaan.Z < 0 Then
'        rJup = 1 + (-8.74789916 * vMaan.Z) / (2095.20826331 * sGeo.r) - rMaanJup(maan)
      End If
'rJup = 1 + (-8.74789916 * vMaan.Z) / (2095.20826331 * sGeo.r) - rMaanJup(maan)
  End If
End Sub

Private Function BerekenTijdstip(maan As Long, bShadow As Boolean, _
                          ByRef JD As Double, ByRef vMaan As TVECTOR, _
                          richting As Long) As Boolean

Dim change As Boolean
Dim A As Double, r As Double, dT As Double, ep As Double
Dim rJup As Double     '{de schaduw/lichtkegel op afstand r van jupiter is kleiner/groter
                       ' dan 1, als je niet op jupiter bivakkeert}
Dim stap As Long
Dim JDoud As Double
     JDoud = JD
     stap = 0
     change = True
     Do While (change) And (stap < 51)
         stap = stap + 1
         Call BerekenPositieMaan(maan, JD, bShadow, vMaan, rJup)
         ep = asin(rJup / rMaan(maan)) * RToD
         r = Sqr(vMaan.x * vMaan.x + 1.071374 * 1.071374 * vMaan.Y * vMaan.Y)
'{         if vmaan(1)<0 then
'           r =-r}
         A = r / rMaan(maan)
         If A > 1 Then A = 1
         If A < -1 Then A = -1
         A = asin(A) * RToD
         If A > 180 Then A = A - 360
         If A < -180 Then A = A + 360
         dT = (ep - A) / oMaan(maan)
'{         if ep<0 then
'           dtas =-dt   }
         JD = JD + dT * richting
         change = Abs(dT) > 0.00001 '{3s grens}
     Loop
     If stap >= 50 Then JD = JDoud
     BerekenTijdstip = stap < 50
End Function

Function Bedekking(maan As Long, JD As Double) As Boolean

Dim vMaan As TVECTOR
Dim rJup As Double, r As Double

     Call BerekenPositieMaan(maan, JD, False, vMaan, rJup)
     r = Sqr(vMaan.x * vMaan.x + 1.071374 * 1.071374 * vMaan.Y * vMaan.Y) + rMaanJup(maan)
     If (r <= rJup) And (vMaan.Z > 0) Then
         Bedekking = True
     Else
         Bedekking = False
     End If
End Function

Function VERDUISTERD(maan As Long, JD As Double) As Boolean

Dim vMaan As TVECTOR
Dim rJup As Double, r As Double

     Call BerekenPositieMaan(maan, JD, True, vMaan, rJup)
     r = Sqr(vMaan.x * vMaan.x + 1.071374 * 1.071374 * vMaan.Y * vMaan.Y) + rMaanJup(maan)
     If (r <= rJup) And (vMaan.Z > 0) Then
         VERDUISTERD = True
     Else
         VERDUISTERD = False
    End If
End Function


Private Sub bereken()

Dim maan As Long
Dim JD As Double
Dim JDV As Double
Dim vMaan As TVECTOR
Dim ddate      As tDatum
Dim s As String, s1 As String
Dim hJaar As Integer
Dim nStap As Double
Dim aResult()
Dim nAantal As Long

nAantal = 0
ReDim Preserve aResult(nAantal)

'{Schaduw overgangen}
ddate.jj = frmPlanets.Year
Call WeekDate(ddate.jj * 100 + 1, ddate)
'dDate.MM = 1
'dDate.DD = 1
JD0 = KalenderNaarJD(ddate)
ddate.jj = frmPlanets.Year + 1
Call WeekDate(ddate.jj * 100 + 1, ddate)
jde = KalenderNaarJD(ddate)
Call Zomertijd_Wintertijd(frmPlanets.Year, JD_ZT, JD_WT)
'{JD_ZT/JD_WT zijn berekend voor 0h UT}
JD_ZT = JD_ZT + 2 / 24 '{= 3h WT}
JD_WT = JD_WT + 1 / 24 '{= 3h ZT}

'{intrede}
nStap = 0
blnDoorgaan = True
For maan = 1 To 4
    JD = JD0
    Do Until JD >= jde + 0.3 Or Not blnDoorgaan
        JDV = JD
        PBVoortgang.value = nStap + ((JD - JD0) / (jde - JD0) * 100) / 16
        DoEvents
        If BerekenTijdstip(maan, True, JD, vMaan, -1) Then
          ddate = JDNaarKalender(JD)
          If (Zichtbaar(JD)) And (JD >= JD0) And (JD <= jde) Then
              ddate = JDNaarKalender(JD - TijdCorrectie(JD, JD_ZT, JD_WT))
              s = StrDate(ddate)
              s1 = StrHMS(Frac(ddate.DD) * Pi2, 4)
              If vMaan.Z > 0 Then
                 If Not (Bedekking(maan, JD)) Then
                    nAantal = nAantal + 1
                    ReDim Preserve aResult(nAantal)
                    aResult(nAantal) = s & " " & s1 & " VI " & maan
                    'Me.txtVerschijnselen = Me.txtVerschijnselen & vbCrLf & _
                        s & " " & s1 & " VI " & maan
'                      write(uit,s,' ',s1)
'                      WRITE(s,' ',s1)
'                      writeln(uit,' VI ',Maan)
'                      writeln(' VI ',Maan)
                 End If
              Else
                    nAantal = nAantal + 1
                    ReDim Preserve aResult(nAantal)
                    aResult(nAantal) = s & " " & s1 & " SI " & maan
'                    Me.txtVerschijnselen = Me.txtVerschijnselen & vbCrLf & _
                        s & " " & s1 & " SI " & maan
'                  write(uit,s,' ',s1)
'                  WRITE(s,' ',s1)
'                  writeln(uit,' SI ',Maan)
'                  writeln(' SI ',Maan)
              End If
          End If
        End If
        JD = JD + (dMaan(maan) / 2)
        If Abs(JD - JDV) < 0.00001 Then
            JD = JD + dMaan(maan) / 2
        End If
    Loop

    nStap = nStap + 6.25
'    {uittrede}
    JD = JD0
    Do Until JD >= jde + 0.3 Or Not blnDoorgaan
        JDV = JD
        PBVoortgang.value = nStap + ((JD - JD0) / (jde - JD0) * 100) / 16
        DoEvents
        If BerekenTijdstip(maan, True, JD, vMaan, 1) Then
          ddate = JDNaarKalender(JD)
          If (Zichtbaar(JD)) And (JD >= JD0) And (JD <= jde) Then
              ddate = JDNaarKalender(JD - TijdCorrectie(JD, JD_ZT, JD_WT))
              s = StrDate(ddate)
              s1 = StrHMS(Frac(ddate.DD) * Pi2, 4)
              If vMaan.Z > 0 Then
                   If Not (Bedekking(maan, JD)) Then
                    nAantal = nAantal + 1
                    ReDim Preserve aResult(nAantal)
                    aResult(nAantal) = s & " " & s1 & " VU " & maan
'                    Me.txtVerschijnselen = Me.txtVerschijnselen & vbCrLf & _
                        s & " " & s1 & " VU " & maan
'                        write(uit,s,' ',s1)
'                        WRITE(s,' ',s1)
'                        writeln(uit,' VU ',Maan)
'                        writeln(' VU ',Maan)
                   End If
              Else
                    nAantal = nAantal + 1
                    ReDim Preserve aResult(nAantal)
                    aResult(nAantal) = s & " " & s1 & " SU " & maan
'                    Me.txtVerschijnselen = Me.txtVerschijnselen & vbCrLf & _
                        s & " " & s1 & " SU " & maan
'                   write(uit,s,' ',s1)
'                   WRITE(s,' ',s1)
'                   writeln(uit,' SU ',Maan)
'                   writeln(' SU ',Maan)
              End If
          End If
        End If
        JD = JD + dMaan(maan) / 2
        If Abs(JD - JDV) < 0.00001 Then
            JD = JD + dMaan(maan) / 2
        End If
    Loop
    nStap = nStap + 6.25
Next
'{Bedekkingen overgangen}

'{intrede}
For maan = 1 To 4
    hJaar = 0
    JD = JD0
    Do Until JD >= jde + 1 Or Not blnDoorgaan
        JDV = JD
        PBVoortgang.value = nStap + ((JD - JD0) / (jde - JD0) * 100) / 16
        DoEvents
        If BerekenTijdstip(maan, False, JD, vMaan, -1) Then
          ddate = JDNaarKalender(JD)
          If (Zichtbaar(JD)) And (JD >= JD0) And (JD <= jde) Then
              ddate = JDNaarKalender(JD - TijdCorrectie(JD, JD_ZT, JD_WT))
              s = StrDate(ddate)
              s1 = StrHMS(Frac(ddate.DD) * Pi2, 4)
              If vMaan.Z > 0 Then
                   If Not (VERDUISTERD(maan, JD)) Then
                    nAantal = nAantal + 1
                    ReDim Preserve aResult(nAantal)
                    aResult(nAantal) = s & " " & s1 & " BI " & maan
'                    Me.txtVerschijnselen = Me.txtVerschijnselen & vbCrLf & _
                        s & " " & s1 & " BI " & maan
'                        write(uit,s,' ',s1)
'                        WRITE(s,' ',s1)
'                        writeln(uit,' BI ',Maan)
'                        writeln(' BI ',Maan)
                   End If
              Else
                    nAantal = nAantal + 1
                    ReDim Preserve aResult(nAantal)
                    aResult(nAantal) = s & " " & s1 & " OI " & maan
'                    Me.txtVerschijnselen = Me.txtVerschijnselen & vbCrLf & _
                        s & " " & s1 & " OI " & maan
'                   write(uit,s,' ',s1)
'                   WRITE(s,' ',s1)
'                   writeln(uit,' OI ',Maan)
'                   writeln(' OI ',Maan)
              End If
          End If
        End If
        JD = JD + dMaan(maan) / 2
        If Abs(JD - JDV) < 0.00001 Then
            JD = JD + dMaan(maan) / 2
        End If
    Loop

'    {uittrede}
    nStap = nStap + 6.25
    hJaar = 0
    JD = JD0
    Do Until JD >= jde + 1 Or Not blnDoorgaan
        JDV = JD
        PBVoortgang.value = nStap + ((JD - JD0) / (jde - JD0) * 100) / 16
        DoEvents
        If BerekenTijdstip(maan, False, JD, vMaan, 1) Then
          ddate = JDNaarKalender(JD)
          If (Zichtbaar(JD)) And (JD >= JD0) And (JD <= jde) Then
              ddate = JDNaarKalender(JD - TijdCorrectie(JD, JD_ZT, JD_WT))
              s = StrDate(ddate)
              s1 = StrHMS(Frac(ddate.DD) * Pi2, 4)
              If vMaan.Z > 0 Then
                   If Not (VERDUISTERD(maan, JD)) Then
                    nAantal = nAantal + 1
                    ReDim Preserve aResult(nAantal)
                    aResult(nAantal) = s & " " & s1 & " BU " & maan
'                    Me.txtVerschijnselen = Me.txtVerschijnselen & vbCrLf & _
                        s & " " & s1 & " BU " & maan
'                       write(uit,s,' ',s1)
'                       WRITE(s,' ',s1)
'                       writeln(uit,' BU ',Maan)
'                       writeln(' BU ',Maan)
                   End If
              Else
                    nAantal = nAantal + 1
                    ReDim Preserve aResult(nAantal)
                    aResult(nAantal) = s & " " & s1 & " OU " & maan
'                    Me.txtVerschijnselen = Me.txtVerschijnselen & vbCrLf & _
                        s & " " & s1 & " OU " & maan
'                   write(uit,s,' ',s1)
'                   WRITE(s,' ',s1)
'                   writeln(uit,' OU ',Maan)
'                   writeln(' OU ',Maan)
              End If
          End If
        End If
        JD = JD + dMaan(maan) / 2
        If Abs(JD - JDV) < 0.00001 Then
            JD = JD + dMaan(maan) / 2
        End If
    Loop
    nStap = nStap + 6.25
Next

For i = 1 To nAantal
    aResult(i) = Mid(aResult(i), 7, 4) & "-" & Mid(aResult(i), 4, 2) & "-" & Left(aResult(i), 2) & _
                        Mid(aResult(i), 11)
Next

Call QuickSort(aResult, 0, nAantal)
For i = 1 To nAantal
    Me.txtVerschijnselen = Me.txtVerschijnselen & _
        Mid(aResult(i), 9, 2) & "-" & Mid(aResult(i), 6, 2) & "-" & Left(aResult(i), 4) & Mid(aResult(i), 11) & vbCrLf
Next
End Sub


Private Sub rtfOverigeCode_Change()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
blnDoorgaan = False
End Sub



Sub print_rtf_textbox(nfile As Long, ByVal srtf_code As String, ByVal x As Long, ByVal Y As Long, ByVal B As Double, H As Double, T As String, kleur As String)
Dim nPos As Long
    nShplId = nShplId + 1
    x = x - schuif: Y = Y + schuif - H / 2
    If Trim(T) = vbNullString Then
        Exit Sub
    End If
    nPos = InStr(srtf_code, "<LEFT>")
    Do While InStr(srtf_code, "<LEFT>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(x, "0") + Mid(srtf_code, nPos + 6)
        nPos = InStr(srtf_code, "<LEFT>")
    Loop
    nPos = InStr(srtf_code, "<TOP>")
    Do While InStr(srtf_code, "<TOP>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(Y, "0") + Mid(srtf_code, nPos + 5)
        nPos = InStr(srtf_code, "<TOP>")
    Loop
    nPos = InStr(srtf_code, "<RIGHT>")
    Do While InStr(srtf_code, "<RIGHT>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(x + B, "0") + Mid(srtf_code, nPos + 7)
        nPos = InStr(srtf_code, "<RIGHT>")
    Loop
    nPos = InStr(srtf_code, "<BOTTOM>")
    Do While InStr(srtf_code, "<BOTTOM>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(Y + H, "0") + Mid(srtf_code, nPos + 8)
        nPos = InStr(srtf_code, "<BOTTOM>")
    Loop
    nPos = InStr(srtf_code, "<BOTTOM-TOP>")
    Do While InStr(srtf_code, "<BOTTOM-TOP>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(H, "0") + Mid(srtf_code, nPos + 12)
        nPos = InStr(srtf_code, "<BOTTOM-TOP>")
    Loop
    nPos = InStr(srtf_code, "<RIGHT-LEFT>")
    Do While InStr(srtf_code, "<RIGHT-LEFT>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(B, "0") + Mid(srtf_code, nPos + 12)
        nPos = InStr(srtf_code, "<RIGHT-LEFT>")
    Loop
    nPos = InStr(srtf_code, "<TEKST>")
    Do While InStr(srtf_code, "<TEKST>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + T + Mid(srtf_code, nPos + 7)
        nPos = InStr(srtf_code, "<TEKST>")
    Loop
    nPos = InStr(srtf_code, "<KLEUR>")
    Do While InStr(srtf_code, "<KLEUR>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(kleur, "0") + Mid(srtf_code, nPos + 7)
        nPos = InStr(srtf_code, "<KLEUR>")
    Loop
    nPos = InStr(srtf_code, "<SHPLID>")
    Do While InStr(srtf_code, "<SHPLID>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(nShplId, "0") + Mid(srtf_code, nPos + 8)
        nPos = InStr(srtf_code, "<SHPLID>")
    Loop
    Print #nfile, srtf_code;
End Sub

Sub print_rtf_circle(nfile As Long, ByVal srtf_code As String, ByVal x As Long, ByVal Y As Long, ByVal r As Long, _
                     ByVal Vul_kleur As Long, maan As Long)
Dim nPos As Long
    nShplId = nShplId + 1
    If Vul_kleur = JUP_NORMAAL Then Vul_kleur = 16777215
    If Vul_kleur = JUP_BEDEKT Then Vul_kleur = 16776960
    If Vul_kleur = JUP_SCHADUW Then Vul_kleur = 0
    If Vul_kleur = JUP_VERDUISTERD Then Vul_kleur = 16711680
    x = x - schuif
    Y = Y + schuif
    If blnZwart And r = 105 And maan = 0 Then
        'srtf_code = jupimage(x - R, y - R)
        Print #nfile, jupimage(x - r, Y - r);
        Exit Sub
    End If
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
    If maan > 0 Then
        If blnZwart Then
            Call print_rtf_textbox(nfile, rtf_code(13), x + schuif, Y - schuif - 170, 170, 227, Format(maan, "0"), "16")
        Else
            Call print_rtf_textbox(nfile, rtf_code(13), x + schuif, Y - schuif - 170, 170, 227, Format(maan, "0"), "1")
        End If
    End If
End Sub

Sub x()
Dim l As Long: l = 255 * 255 * 255: Debug.Print l
End Sub

Function jupimage(x As Long, Y As Long) As String
Dim srtf_code As String
nShplId = nShplId + 1
srtf_code = rtf_code(14)
nPos = InStr(srtf_code, "<LEFT>")
Do While InStr(srtf_code, "<LEFT>") > 0
    srtf_code = Left(srtf_code, nPos - 1) + Format(x, "0") + Mid(srtf_code, nPos + 6)
    nPos = InStr(srtf_code, "<LEFT>")
Loop
nPos = InStr(srtf_code, "<RIGHT>")
Do While InStr(srtf_code, "<RIGHT>") > 0
    srtf_code = Left(srtf_code, nPos - 1) + Format(x + 238, "0") + Mid(srtf_code, nPos + 7)
    nPos = InStr(srtf_code, "<RIGHT>")
Loop
nPos = InStr(srtf_code, "<BOTTOM>")
Do While InStr(srtf_code, "<BOTTOM>") > 0
    srtf_code = Left(srtf_code, nPos - 1) + Format(Y + 238, "0") + Mid(srtf_code, nPos + 8)
    nPos = InStr(srtf_code, "<BOTTOM>")
Loop
nPos = InStr(srtf_code, "<TOP>")
Do While InStr(srtf_code, "<TOP>") > 0
    srtf_code = Left(srtf_code, nPos - 1) + Format(Y, "0") + Mid(srtf_code, nPos + 5)
    nPos = InStr(srtf_code, "<TOP>")
Loop
nPos = InStr(srtf_code, "<LEFT_1>")
Do While InStr(srtf_code, "<LEFT_1>") > 0
    srtf_code = Left(srtf_code, nPos - 1) + Format(x - 1, "0") + Mid(srtf_code, nPos + 8)
    nPos = InStr(srtf_code, "<LEFT_1>")
Loop
nPos = InStr(srtf_code, "<TOP_1>")
Do While InStr(srtf_code, "<TOP_1>") > 0
    srtf_code = Left(srtf_code, nPos - 1) + Format(Y - 1, "0") + Mid(srtf_code, nPos + 7)
    nPos = InStr(srtf_code, "<TOP_1>")
Loop
    nPos = InStr(srtf_code, "<SHPLID>")
    Do While InStr(srtf_code, "<SHPLID>") > 0
        srtf_code = Left(srtf_code, nPos - 1) + Format(nShplId, "0") + Mid(srtf_code, nPos + 8)
        nPos = InStr(srtf_code, "<SHPLID>")
    Loop
jupimage = srtf_code
End Function
