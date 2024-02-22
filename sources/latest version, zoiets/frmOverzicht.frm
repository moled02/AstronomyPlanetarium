VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmOverzicht 
   BackColor       =   &H80000005&
   Caption         =   "Zichtbaarheidsdiagram"
   ClientHeight    =   10350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10575
   DrawStyle       =   1  'Dash
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOverzicht.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10350
   ScaleWidth      =   10575
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox rtfResultaat 
      Height          =   3735
      Left            =   240
      TabIndex        =   1
      Top             =   6480
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   6588
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmOverzicht.frx":030A
   End
   Begin VB.PictureBox picDiagram 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6445
      Left            =   0
      ScaleHeight     =   426
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   699
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   10545
   End
   Begin VB.Label Labelv1 
      BackColor       =   &H80000005&
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   14
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lbPlaneten 
      BackColor       =   &H80000005&
      Caption         =   "Zichtbaarheidsdiagram"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3000
      TabIndex        =   41
      Top             =   6000
      Width           =   6935
   End
   Begin VB.Label Labelv1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   13
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelv1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   12
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelv1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   11
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelv1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   10
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelv1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelv1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelv1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelv1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelv1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelv1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelv1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelv1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelv1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelv1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   24
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   23
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   22
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   21
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   20
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   19
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   18
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   17
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   16
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   15
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   14
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   13
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   12
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   11
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   10
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Labelh1 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   400
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmOverzicht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MERGEPAINT = &HBB0226
Private Const SRCAND = &H8800C6
Private Const SRCCOPY = &HCC0020

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private nPlaneet As Long
Private Type tgebied
    aantgeb As Long
    Y(5) As Boolean
    x(5) As Double
End Type
Private Type tZonGebied
    gevuld As Boolean
    gebied As tgebied
    zonopk As Double
    zonond As Double
    Opk As Double
    Ond As Double
End Type
Const nC = 400
Const x0 = 500
Const y0 = 400
Private Sub ResGebieden(geb1 As tgebied, geb2 As tgebied, resgeb As tgebied)


Dim res As Long, g1 As Long, g2 As Long

    g1 = 1
    g2 = 1
    res = 0
    Do While g1 + g2 <= geb1.aantgeb + geb2.aantgeb
         res = res + 1
         resgeb.Y(res) = (geb1.Y(g1)) And (geb2.Y(g2))
         If geb1.x(g1) = geb2.x(g2) Then
              resgeb.x(res) = geb1.x(g1)
              g1 = g1 + 1
              g2 = g2 + 1
         Else
             If geb1.x(g1) < geb2.x(g2) Then
                  resgeb.x(res) = geb1.x(g1)
                  g1 = g1 + 1
             Else
                  resgeb.x(res) = geb2.x(g2)
                  g2 = g2 + 1
             End If
        End If
    Loop
    resgeb.aantgeb = res
End Sub

Private Sub MaakGebied(T0 As Double, t1 As Double, ByRef geb As tgebied)

     If T0 < t1 Then  '{|   ---   |}
          geb.x(1) = T0
          geb.Y(1) = False
          geb.x(2) = t1
          geb.Y(2) = True
          geb.x(3) = 24
          geb.Y(3) = False
          geb.aantgeb = 3
     Else
          geb.x(1) = t1
          geb.Y(1) = True
          geb.x(2) = T0
          geb.Y(2) = False
          geb.x(3) = 24
          geb.Y(3) = True
          geb.aantgeb = 3
     End If
End Sub

Private Sub Inverse(ByRef geb1 As tgebied)

Dim i As Long

    For i = 1 To geb1.aantgeb
        geb1.Y(i) = Not (geb1.Y(i))
    Next
        
End Sub

Private Sub ov(Planet As Long, ddate As tDatum, _
                   ObsLon As Double, ObsLat As Double, TimeZone As Double, Height As Double, _
                   ByRef Opk As Double, ByRef Ond As Double)

Dim SHelio As TSVECTOR, SGeo As TSVECTOR, SSun As TSVECTOR
    'Q1 = SHelio, Q2 = SGeo
Dim sAarde As TSVECTOR
Dim RA As Double
Dim Decl As Double
Dim RA1 As Double
Dim Decl1 As Double
Dim RA2 As Double
Dim Decl2 As Double
Dim dat As tDatum
Dim tt As Double
Dim t As Double
Dim T0 As Double 'tijdstip op 0h
Dim DtofUT As Double
Dim Obl As Double
Dim phase As Double
Dim PhaseAngle As Double
Dim Elongation As Double
Dim Magnitude As Double
Dim Semidiameter As Double
Dim PolarSemiDiameter As Double
Dim NutLon As Double, NutObl As Double
Dim Parallax As Double, MoonHeight As Double
Dim JupiterPhysData As TJUPITERPHYSDATA
Dim MarsPhysData As TMARSPHYSDATA
Dim SunPhysData As TSUNPHYSDATA
Dim SaturnRingData As TSATURNRINGDATA
Dim AltSaturnRingData As TALTSATURNRINGDATA
Dim MoonPhysData As TMOONPHYSDATA
Dim C As Long
Dim JDOfCarr As Double
Dim deltaT As Double
Dim RTS As tRiseSetTran, RTS1 As tRiseSetTran, RTS2 As tRiseSetTran
Dim sLatitude As String, sLongitude As String
Dim LAST As Double
Dim RhoCosPhi As Double, RhoSinPhi As Double
Dim JD0 As Double
Dim sSterbeeld As String
Dim Az As Double, hg As Double, Alt As Double, dAlt As Double, maxhoogte As Double
dat.jj = ddate.jj
dat.mm = ddate.mm
dat.DD = ddate.DD
'tt = (Hrs + Min / 60 + Sec / 3600) / 24
'dat.DD = dat.DD + tt


JD0 = KalenderNaarJD(dat)
    t = JDToT(JD0 + TimeZone) '+ i * Interval_dagen)
    deltaT = ApproxDeltaT(t)
    T0 = (floor(t * 36525 + 0.50001 - TimeZone) - 0.5 + TimeZone) / 36525#
    DtofUT = T0 + secToT * deltaT
    't = DtofUT '+ i * Interval_dagen / 36525#
    't = t + secToT * deltaT
    Call NutationConst(t, NutLon, NutObl)
    Obl = Obliquity(t)
    
    LAST = SiderealTime(t) + NutLon * Cos(Obl) - ObsLon
    Call ObserverCoord(ObsLat, 0, RhoCosPhi, RhoSinPhi)
'    Call ObserverCoord(ObsLat, Height, RhoCosPhi, RhoSinPhi)
   
    '======================== ZON ========================
    If Planet = 0 Then
        SHelio.l = 0: SHelio.B = 0: SHelio.r = 0
    
        ' bepalen opkomst e.d.
        Call PlanetPosHi(0, T0 - 1 / 36525, sAarde, False)
        Call HelioToGeo(SHelio, sAarde, SGeo)
        Call EclToEqu(SGeo.l, SGeo.B, Obl, RA1, Decl1)
        
        Call PlanetPosHi(0, T0, sAarde, chkGrootstePrecisie = False)
        Call HelioToGeo(SHelio, sAarde, SGeo)
        Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
        
        Call PlanetPosHi(0, T0 + 1 / 36525, sAarde, False)
        Call HelioToGeo(SHelio, sAarde, SGeo)
        Call EclToEqu(SGeo.l, SGeo.B, Obl, RA2, Decl2)
            
        Call RiseSet(T0, deltaT, RA1, Decl1, RA, Decl, RA2, Decl2, Height, ObsLon, ObsLat, RTS)
    
    ElseIf Planet > 0 And Planet < 9 Then
      
        ' bepalen opkomst e.d.
        Call PlanetPosHi(0, T0 - 1 / 36525, sAarde, False)
        Call PlanetPosHi(Planet, T0 - 1 / 36525, SHelio, False)
        Call HelioToGeo(SHelio, sAarde, SGeo)
        Call PlanetPosHi(Planet, T0 - 1 / 36525 - SGeo.r * LightTimeConst, SHelio, False)
        Call HelioToGeo(SHelio, sAarde, SGeo)
        Call EclToEqu(SGeo.l, SGeo.B, Obl, RA1, Decl1)
        
        Call PlanetPosHi(0, T0, sAarde, False)
        Call PlanetPosHi(Planet, T0, SHelio, False)
        Call HelioToGeo(SHelio, sAarde, SGeo)
        Call PlanetPosHi(Planet, T0 - SGeo.r * LightTimeConst, SHelio, False)
        Call HelioToGeo(SHelio, sAarde, SGeo)
        Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
        
        Call PlanetPosHi(0, T0 + 1 / 36525, sAarde, False)
        Call PlanetPosHi(Planet, T0 + 1 / 36525, SHelio, False)
        Call HelioToGeo(SHelio, sAarde, SGeo)
        Call PlanetPosHi(Planet, T0 + 1 / 36525 - SGeo.r * LightTimeConst, SHelio, False)
        Call HelioToGeo(SHelio, sAarde, SGeo)
        Call EclToEqu(SGeo.l, SGeo.B, Obl, RA2, Decl2)
        
        Call RiseSet(T0, deltaT, RA1, Decl1, RA, Decl, RA2, Decl2, Height, ObsLon, ObsLat, RTS)
    ElseIf Planet = 9 Then ', dus Pluto
        'Dit is een speciaal geval. Coordinaten zijn voor 2000. Deze moeten voor de Zon worden berekend.
        'Dat is: bereken positie voor vandaag. De coordinaten omzetten naar J2000
        'dit is niet voldoende. Er moet nog een correctie plaatsvinden van Ecl. VSOP -> equ FK5-2000
        'vervolgens de positie van pluto berekenen
        Dim TAarde As TVECTOR
        Dim sZon As TSVECTOR
        Dim TPluto As TVECTOR
        SHelio.l = 0: SHelio.B = 0: SHelio.r = 0
        Call PlanetPosHi(0, t, sAarde, False)
        Call HelioToGeo(SHelio, sAarde, SGeo)
        Call SphToRect(SGeo, TAarde)
        Call EclToEqu(SGeo.l, SGeo.B, Obl, RA, Decl)
        ' Call Reduction2000(0, RA, Decl)
        'coordinaten omzetten naar J2000
        Call PrecessFK5(t, 0, RA, Decl)
        Call EquToEcl(RA, Decl, Obliquity(0), SGeo.l, SGeo.B)
        Call SphToRect(SGeo, TAarde)
        Call EclVSOP2000_equFK52000(TAarde.x, TAarde.Y, TAarde.Z)
        Call RectToSph(TAarde, sZon)
        sAarde = SGeo
        
        Call PlanetPosHi(0, t, sAarde, False)
        Call PlutoPos(t, SHelio)
        Call EclToRect(SHelio, Obliquity(0), TPluto)
        Dist = Sqr((TAarde.x + TPluto.x) * (TAarde.x + TPluto.x) + (TAarde.Y + TPluto.Y) * (TAarde.Y + TPluto.Y) + (TAarde.Z + TPluto.Z) * (TAarde.Z + TPluto.Z))
        Call PlutoPos(t - Dist * LightTimeConst, SHelio)
        Call EclToRect(SHelio, Obliquity(0), TPluto)
        Dist = Sqr((TAarde.x + TPluto.x) * (TAarde.x + TPluto.x) + (TAarde.Y + TPluto.Y) * (TAarde.Y + TPluto.Y) + (TAarde.Z + TPluto.Z) * (TAarde.Z + TPluto.Z))
        RA = atan2(TPluto.Y + TAarde.Y, TPluto.x + TAarde.x)
        If RA < 0 Then
            RA = RA + Pi2
        End If
        Decl = asin((TPluto.Z + TAarde.Z) / Dist)
        Call RiseSet(T0, deltaT, RA, Decl, RA, Decl, RA, Decl, Height, ObsLon, ObsLat, RTS)
        
    End If
    
    If RTS.Rise < 0 Then RTS.Rise = 0
    If RTS.Setting < 0 Then RTS.Setting = Pi2

    Opk = RTS.Rise * RToD / 15
    Ond = RTS.Setting * RToD / 15
End Sub

Sub Uitvoeren(nPlaneet)
Dim ZonGebieden(366 + 7) As tZonGebied
Dim StartDate As tDatum
Dim JD_ZT As Double
Dim JD_WT As Double
Dim ObsLat As Double, ObsLon As Double
Dim sLatitude As String, sLongitude As String
Dim TimeZone As Double
Dim zonopk As Double, zonond As Double
Dim Opk As Double, Ond As Double
Dim geb1 As tgebied, geb2 As tgebied
Dim gebied As tgebied, resgeb As tgebied
Dim i As Long
Dim oldgebied As tgebied
Dim holdgebied As tgebied
Dim hstr As String
Dim JD As Double

Dim ddate As tDatum, hdate As tDatum
Dim weeknr As Long
Dim BeginWeekNr As Long, EindWeekNr As Long
Dim fNaam
Dim oZichtbaar As Boolean, nAantweken As Long
Dim Y As Long

Me.rtfResultaat.Text = ""

'fNaam = Array("", "Mercury", "Venus", "", "Mars", "Jupiter", "Saturn", "Uranus", "Neptune", "Pluto")
fNaam = Array("", "Mercurius", "Venus", "", "Mars", "Jupiter", "Saturnus", "Uranus", "Neptunus", "Pluto")

ddate.jj = frmPlanets.Year
ddate.mm = 1
ddate.DD = 1
sUitvoer = ""
BeginWeekNr = Int(ddate.jj * 100#) + 1
Call WeekDate(BeginWeekNr, ddate)
JD0 = KalenderNaarJD(ddate)
Call Zomertijd_Wintertijd(frmPlanets.Year, JD_ZT, JD_WT)

Call WeekDate(BeginWeekNr, ddate)

Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Latitude", _
            sLatitude)
ObsLat = Val(sLatitude) * Pi / 180
Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Longitude", _
            sLongitude)
ObsLon = Val(sLongitude) * Pi / 180

EindWeekNr = Int((frmPlanets.Year + 1) * 100#) + 1
Call WeekDate(EindWeekNr, hdate)
hdate = JDNaarKalender(KalenderNaarJD(hdate) - 7)
'hdate = JDNaarKalender(JD)
EindWeekNr = WeekOfYear(hdate)
nAantweken = Val(Right(str(EindWeekNr), 2))

TimeZone = 24 '{een start waarde}

For zonteller = 1 To 7 * (EindWeekNr - BeginWeekNr + 1)
  ZonGebieden(zonteller).gevuld = False
Next
For i = nPlaneet To nPlaneet
Me.AutoRedraw = True
Cls
Me.Refresh
Me.AutoRedraw = False

Call TekenKader(True, nAantweken)

If i <> 3 Then
'    rtfResultaat.Text = rtfResultaat.Text + "============================================================" + vbCrLf
'    rtfResultaat.Text = rtfResultaat.Text + fNaam(i) + vbCrLf
'    rtfResultaat.Text = rtfResultaat.Text + "-----------------" + vbCrLf
    lbPlaneten.Caption = "Zichtbaarheidsdiagram van " & fNaam(i) & " voor het jaar " & Format(frmPlanets.Year, "0")
    oZichtbaar = False
    hstr = ""
  oldgebied.aantgeb = 0
  TimeZone = 24
  Y = y0 + 15
  JD = KalenderNaarJD(ddate)
  Call WeekDate(BeginWeekNr, ddate)
'  Date = StartDate
  oweeknr = -1
  weeknr = WeekOfYear(ddate)
  zonteller = 0
  Do While weeknr <= EindWeekNr
     zonteller = zonteller + 1
     oTimeZone = TimeZone
     If oweeknr <> weeknr Then
         TimeZone = TijdCorrectie(JD, JD_ZT, JD_WT)
         oweeknr = weeknr
     End If
     With ZonGebieden(zonteller)
         If Not gevuld Then Call ov(0, ddate, ObsLon, ObsLat, TimeZone, 0, zonopk, zonond)
         If i < 4 Then
            Call ov(0, ddate, ObsLon, ObsLat, TimeZone, -3 * DToR, Opk, Ond)
         ElseIf i < 9 Then
           Call ov(0, ddate, ObsLon, ObsLat, TimeZone, -6 * DToR, Opk, Ond)
         ElseIf i = 9 Then
           Call ov(0, ddate, ObsLon, ObsLat, TimeZone, -18 * DToR, Opk, Ond)
         End If
'         If (i = 1) Or (i = 4) Or (i = 9) Then
             Call MaakGebied(Opk, Ond, geb1)
             Call Inverse(geb1)
             ZonGebieden(zonteller).gevuld = True
             ZonGebieden(zonteller).gebied = geb1
'         End If
     End With
     If i <= 2 Then
       Call ov(i, ddate, ObsLon, ObsLat, TimeZone, 3 * DToR, Opk, Ond)
     Else
       If i <= 7 Then
         Call ov(i, ddate, ObsLon, ObsLat, TimeZone, 5 * DToR, Opk, Ond)
       Else
         Call ov(i, ddate, ObsLon, ObsLat, TimeZone, 10 * DToR, Opk, Ond)
       End If
     End If
     
     Call MaakGebied(Opk, Ond, geb2)
     If oTimeZone <> TimeZone Then
       oldgebied.aantgeb = 0
     Else
       oldgebied = resgeb
     End If
     With ZonGebieden(zonteller)
       Call ResGebieden(.gebied, geb2, resgeb)
       'Call DrukGrafRes(resgeb, oldgebied, y, ZonOpk, ZonOnd)
       If oZichtbaar <> IsZichtbaar(resgeb) Then
        hstr = hstr + vbCrLf
       End If
       oZichtbaar = IsZichtbaar(resgeb)
       Call DrukafRes(ddate, resgeb, hstr)
       holdgebied = oldgebied
       Call DrukGrafRes(resgeb, oldgebied, Y, zonopk, zonond)
       oldgebied = holdgebied
     End With
     Y = Y + 15
     JD = KalenderNaarJD(ddate)
     ddate = JDNaarKalender(JD + 1)
     weeknr = WeekOfYear(ddate)
     DoEvents
Loop
Call TekenKader(False, nAantweken)
rtfResultaat.Text = rtfResultaat.Text + hstr
End If
Next
End Sub

Private Sub DrukafRes(ddate As tDatum, geb As tgebied, hstr As String)

Dim i As Long
Dim AantZichtUur As Double
Dim blnZichtbaar As Boolean

blnZichtbaar = False
For i = 1 To geb.aantgeb
    If geb.Y(i) = True Then blnZichtbaar = True
Next

If Not blnZichtbaar Then
    Exit Sub
End If
    hstr = hstr + Format(ddate.DD, "00") + "-" + Format(ddate.mm, "00") + "-" + Format(ddate.jj, "0000") + " : " & vbTab
   AantZichtUur = 0
   If geb.Y(1) = True Then
       hstr = hstr + "00h00m -- "
       hstr = hstr + StrHMS(geb.x(1) * 15 * DToR, 2) + "  "
       AantZichtUur = geb.x(1)
   End If
   For i = 2 To geb.aantgeb
     If geb.Y(i) = True Then
         hstr = hstr + StrHMS(geb.x(i - 1) * 15 * DToR, 2) + " -- "
         If i = geb.aantgeb Then
             hstr = hstr + "24h00m  "
             AantZichtUur = AantZichtUur + 24 - geb.x(i - 1)
         Else
             hstr = hstr + StrHMS(geb.x(i) * 15 * DToR, 2) + "  "
             AantZichtUur = AantZichtUur + geb.x(i) - geb.x(i - 1)
         End If
     End If
  Next
  hstr = hstr + "    ----> " + StrHMS(AantZichtUur * 15 * DToR, 2) + vbCrLf
End Sub

Private Function IsZichtbaar(geb As tgebied)
Dim i As Long

IsZichtbaar = False
For i = 1 To geb.aantgeb
    If geb.Y(i) = True Then IsZichtbaar = True
Next
End Function
Private Sub Form_Activate()
nPlaneet = 1
Me.AutoRedraw = True
Call Uitvoeren(nPlaneet)
End Sub


Private Sub DrukGrafRes(resgeb As tgebied, oldgebied As tgebied, Y As Long, zonopk As Double, zonond As Double)


Dim i As Long, j As Long, k As Long
Dim hGeb As tgebied
Dim hresgeb As tgebied
Dim holdgebied As tgebied

'    For i = 1 To resgeb.aantgeb
'        hresgeb.x(i) = resgeb.x(i)
'        hresgeb.y(i) = resgeb.y(i)
'    Next
'    hresgeb.aantgeb = resgeb.aantgeb
    
    hresgeb = resgeb
'    For i = 1 To oldgebied.aantgeb
'        holdgebied.x(i) = oldgebied.x(i)
'        holdgebied.y(i) = oldgebied.y(i)
'    Next
'    holdgebied.aantgeb = oldgebied.aantgeb
    holdgebied = oldgebied

  picDiagram.Refresh
   j = 99
   For i = 1 To hresgeb.aantgeb
        If hresgeb.x(i) > 12 Then
           If i < j Then
              j = i
           End If
        End If
    Next
   i = j
   Do While i <= hresgeb.aantgeb
        hGeb.x(i - j + 1) = hresgeb.x(i) + 12
        hGeb.x(i - j + 1) = hGeb.x(i - j + 1) - 24 * Int(hGeb.x(i - j + 1) / 24)
        hGeb.Y(i - j + 1) = hresgeb.Y(i)
        i = i + 1
   Loop
   If j = 99 Then j = 0
   k = i - j
   i = 1
   Do While i < j
        hGeb.x(i + k) = hresgeb.x(i) + 12
        hGeb.x(i + k) = hGeb.x(i + k) - 24 * Int(hGeb.x(i + k) / 24)
        hGeb.Y(i + k) = hresgeb.Y(i)
        i = i + 1
   Loop
   hGeb.aantgeb = hresgeb.aantgeb
   hresgeb = hGeb
'   For i = 1 To resgeb.aantgeb
'        resgeb.x(i) = resgeb.x(i) * 15 / 24
'   Next

   j = 99
   For i = 1 To holdgebied.aantgeb
        If holdgebied.x(i) > 12 Then
           If i < j Then
              j = i
           End If
        End If
   Next
   i = j
   Do While i <= holdgebied.aantgeb
        hGeb.x(i - j + 1) = holdgebied.x(i) + 12
        hGeb.x(i - j + 1) = hGeb.x(i - j + 1) - 24 * Int(hGeb.x(i - j + 1) / 24)
        hGeb.Y(i - j + 1) = holdgebied.Y(i)
        i = i + 1
   Loop
   If j = 99 Then j = 0
   k = i - j
   i = 1
   Do While i < j
        hGeb.x(i + k) = holdgebied.x(i) + 12
        hGeb.x(i + k) = hGeb.x(i + k) - 24 * Int(hGeb.x(i + k) / 24)
        hGeb.Y(i + k) = holdgebied.Y(i)
        i = i + 1
   Loop
   If j <> 0 Then
      hGeb.aantgeb = holdgebied.aantgeb
      holdgebied = hGeb
'      For i = 1 To oldgebied.aantgeb
'          oldgebied.x(i) = oldgebied.x(i) * 15 / 24
'      Next
   End If
   
   zonopk = zonopk + 12
   zonopk = zonopk - 24 * Int(zonopk / 24)
'   zonopk = zonopk * 15 / 24
   zonond = zonond + 12
   zonond = zonond - 24 * Int(zonond / 24)
'   zonond = zonond * 15 / 24

   If hresgeb.Y(1) = True Then
        DrawStyle = vbSolid
        ForeColor = RGB(200, 100, 100)
        Line (x0, Y)-(x0 + Int(nC * hresgeb.x(1)), Y)
'       If (holdgebied.aantgeb > 0) And (holdgebied.y(1) = True) Then
'        ForeColor = RGB(255, 0, 0)
'        DrawStyle = vbSolid
        
'         Line (x0 + (nC * hresgeb.x(1)), y)-(x0 + (nC * holdgebied.x(1)), y - 1)
'         {pen.style = psDot}
'       End If
   End If

   For i = 2 To hresgeb.aantgeb
     If hresgeb.Y(i) = True Then
      DrawStyle = vbSolid
      ForeColor = RGB(50, 50, 155)
      If i = hresgeb.aantgeb Then
           Line (x0 + (nC * hresgeb.x(i - 1)), Y)-(x0 + (nC * hresgeb.x(i)), Y)
       Else
           Line (x0 + (nC * hresgeb.x(i - 1)), Y)-(x0 + (nC * hresgeb.x(i)), Y)
       End If
      ' If (holdgebied.aantgeb > 0) And (holdgebied.y(i) = True) Then
      '   ForeColor = RGB(0, 0, 0)
      '   DrawStyle = vbSolid
      '   Line (x0 + (nC * hresgeb.x(i - 1)), y)-(x0 + (nC * holdgebied.x(i - 1)), y - 1)
      '   Line (x0 + (nC * hresgeb.x(i)), y)-(x0 + (nC * holdgebied.x(i)), y - 1)
'{         pen.style = psDot  }
      ' End If
     End If
   Next
    ForeColor = RGB(0, 0, 0)
    DrawStyle = vbSolid
   PSet (x0 + Int(nC * zonopk), Y), RGB(0, 0, 0)
   PSet (x0 + Int(nC * zonond), Y), RGB(0, 0, 0)

'   picDiagram.Refresh
'   picDiagram.Cls
End Sub

Private Sub DrukGrafRes2(resgeb As tgebied, oldgebied As tgebied, Y As Long, zonopk As Double, zonond As Double)

Dim i As Long

   If resgeb.Y(1) = True Then
'       Pixels(x0, y) = clblack
'       Pixels(x0 + Int(40 * resgeb.x(1)), y) = clblack
   End If

   For i = 2 To resgeb.aantgeb
     If resgeb.Y(i) = True Then
       If i = resgeb.aantgeb Then
'           Pixels(x0 + Int(40 * resgeb.x(i - 1)), y) = clblack
'           Pixels(x0 + 40 * 15, y) = clblack
       Else
 '         Pixels(x0 + Int(40 * resgeb.x(i - 1)), y) = clblack
 '         Pixels(x0 + Int(40 * resgeb.x(i)), y) = clblack
       End If
      End If
    Next
'   Pixels(x0 + Int(40 * zonopk), y) = clblack
'   Pixels(x0 + Int(40 * zonond), y) = clblack
End Sub

Sub TekenRechthoek(x1 As Long, y1 As Long, x2 As Long, y2 As Long)
    Line (x1, y1)-(x1, y2), , B
    Line (x1, y2)-(x2, y2), , B
    Line (x2, y2)-(x2, y1), , B
    Line (x2, y1)-(x1, y1), , B
End Sub

Sub TekenKader(lTotaal As Boolean, Aantweken As Long)

Dim i As Long, k As Long, a1 As Long, a2 As Long, x As Long, Y As Long
Const nStapY As Long = 105
If lTotaal Then
    Y = y0 + nStapY * Aantweken
    x = x0 + nC * 24
    For i = 1 To Aantweken
      Y = y0 + nStapY * i
      If i - 4 * Int(i / 4) = 0 Then
          a1 = Int(i / 4)
          Labelv1(a1 - 1).top = Y - 100 - nC
          Labelv1(a1 - 1).Left = x0 - 400
          Labelv1(a1 - 1).Caption = Format(i - 3, "00")
      End If
    Next
    'a1 = a1 + 1
    Y = y0 + nStapY * 56
    Labelv1(a1).top = Y - 100 - nC
    Labelv1(a1).Left = x0 - 400
    If Aantweken = 53 Then
        Labelv1(a1).Caption = Format(53, "00")
    Else
        Labelv1(a1).Caption = Format(1, "00")
    End If
    Labelv1(a1).Visible = False
    a1 = Int(nC / 3)
    a2 = 2 * a1
    '5473/53
    Y = y0 + nStapY * (Aantweken)
    For i = 0 To 24 - 1
      x = x0 + (i * nC)
      k = i + 12
      k = k - 24 * Int(i / 12)
      'Labelh1(i).BackStyle = 0
      Labelh1(i).top = y0 - 300
      Labelh1(i).Left = x0 - 100 + i * nC
      Labelh1(i).Caption = Format(k, "00")
    Next
    'Labelh1(24).BackStyle = 0
    Labelh1(24).top = y0 - 300
    Labelh1(24).Left = x0 - 100 + 24 * nC
    Labelh1(24).Caption = Format(12, "00")
    Me.Refresh
End If
    DrawStyle = vbSolid
    FillStyle = 1
    Y = y0 + nStapY * (Aantweken)
    x = x0 + nC * 24
    If lTotaal Then
        Call TekenRechthoek(x0, y0 - 1, x + 1, Y + 1)
    Else
        Call TekenRechthoek(x0, y0 - 1, x, Y)
    End If
    
    For i = 1 To Aantweken
      Y = y0 + nStapY * i
      If i - 4 * Int(i / 4) = 0 Then
          a1 = Int(i / 4)
          Line (x0, Y)-(x, Y)
      Else
'        Line (x0, Y)-(x0 + 5, Y)
'        Line (x - 5, Y)-(x, Y)
      End If
    Next
    a1 = Int(nC / 3)
    a2 = 2 * a1
    Y = y0 + nStapY * (Aantweken)
    For i = 0 To 24 - 1
      x = x0 + (i * nC)
      Line (x + a1, y0)-(x + a1, y0 + 5)
      Line (x + a2, y0)-(x + a2, y0 + 5)
      Line (x + a1, Y - 1)-(x + a1, Y - 5)
      Line (x + a2, Y - 1)-(x + a2, Y - 5)
      Line (x, y0)-(x, Y)
    Next
End Sub

Private Sub rtfResultaat_KeyUp(KeyCode As Integer, Shift As Integer)
Dim fNaam
Dim nfile
Dim nVorig As Long
Dim sTempFile As String
Dim objspecialfolder As New clsSpecialFolder

nVorig = nPlaneet
Select Case KeyCode
    Case 33 And Shift = 2
        nPlaneet = nPlaneet - 1
        If nPlaneet = 0 Then nPlaneet = 1
        If nPlaneet = 3 Then nPlaneet = 2
    Case 34 And Shift = 2
        nPlaneet = nPlaneet + 1
        If nPlaneet = 3 Then nPlaneet = 4
        If nPlaneet > 9 Then nPlaneet = 9
    Case 83 'Ctrl-S
        fNaam = Array("", "Mercurius", "Venus", "", "Mars", "Jupiter", "Saturnus", "Uranus", "Neptunus", "Pluto")
        Me.AutoRedraw = False
        BitBlt picDiagram.hdc, _
        0, 0, 6465, 10545, _
        Me.hdc, 0, 0, SRCCOPY
        
        picDiagram.Refresh
        sTempFile = objspecialfolder.TemporaryFolder + "\OV_" + fNaam(nPlaneet) + frmPlanets.Year
        
        Call SavePicture(picDiagram.Image, sTempFile + ".bmp")
        Me.AutoRedraw = False
        nfile = FreeFile
        Open sTempFile + ".txt" For Output As nfile
        Print #nfile, Me.rtfResultaat.Text
        Close (nfile)
End Select



If Not nVorig = nPlaneet Then Call Uitvoeren(nPlaneet)
End Sub

