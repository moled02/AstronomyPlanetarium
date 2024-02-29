VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmPlanets 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Astronomy"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   14295
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmPlanets.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   14295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerAnimate 
      Enabled         =   0   'False
      Left            =   10920
      Top             =   720
   End
   Begin VB.CommandButton cmdAnimate 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ANIMATE"
      Height          =   315
      Left            =   9120
      TabIndex        =   39
      Top             =   780
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTabPlanets 
      Height          =   7815
      Left            =   120
      TabIndex        =   26
      Top             =   1200
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   10
      TabsPerRow      =   10
      TabHeight       =   520
      TabCaption(0)   =   "Maan"
      TabPicture(0)   =   "frmPlanets.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "listInfo(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Zon"
      TabPicture(1)   =   "frmPlanets.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "listInfo(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Mercurius"
      TabPicture(2)   =   "frmPlanets.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "listInfo(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Venus"
      TabPicture(3)   =   "frmPlanets.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "listInfo(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Mars"
      TabPicture(4)   =   "frmPlanets.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "listInfo(4)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Jupiter"
      TabPicture(5)   =   "frmPlanets.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "picHidden"
      Tab(5).Control(1)=   "lstMoonsJup"
      Tab(5).Control(2)=   "listInfo(5)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Saturnus"
      TabPicture(6)   =   "frmPlanets.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "picHiddenSat"
      Tab(6).Control(1)=   "lstMoonsSat"
      Tab(6).Control(2)=   "listInfo(6)"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "Uranus"
      TabPicture(7)   =   "frmPlanets.frx":03CE
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "listInfo(7)"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Neptunus"
      TabPicture(8)   =   "frmPlanets.frx":03EA
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "listInfo(8)"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "Pluto"
      TabPicture(9)   =   "frmPlanets.frx":0406
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "listInfo(9)"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).ControlCount=   1
      Begin VB.PictureBox picHiddenSat 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         Height          =   1335
         Left            =   -68040
         ScaleHeight     =   1275
         ScaleWidth      =   6435
         TabIndex        =   41
         Top             =   5040
         Width           =   6495
      End
      Begin VB.PictureBox picHidden 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000004&
         Height          =   1095
         Left            =   -68040
         ScaleHeight     =   1035
         ScaleWidth      =   6555
         TabIndex        =   40
         Top             =   5160
         Width           =   6615
      End
      Begin VB.ListBox lstMoonsSat 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4050
         Left            =   -68040
         TabIndex        =   38
         Top             =   720
         Width           =   6495
      End
      Begin VB.ListBox lstMoonsJup 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4050
         Left            =   -68040
         TabIndex        =   37
         Top             =   720
         Width           =   6495
      End
      Begin VB.ListBox listInfo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4110
         Index           =   9
         ItemData        =   "frmPlanets.frx":0422
         Left            =   -74400
         List            =   "frmPlanets.frx":0429
         MultiSelect     =   2  'Extended
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   720
         Width           =   9015
      End
      Begin VB.ListBox listInfo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4110
         Index           =   8
         ItemData        =   "frmPlanets.frx":0437
         Left            =   -74400
         List            =   "frmPlanets.frx":043E
         MultiSelect     =   2  'Extended
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   720
         Width           =   9015
      End
      Begin VB.ListBox listInfo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4110
         Index           =   7
         ItemData        =   "frmPlanets.frx":044C
         Left            =   -74400
         List            =   "frmPlanets.frx":0453
         MultiSelect     =   2  'Extended
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   720
         Width           =   9015
      End
      Begin VB.ListBox listInfo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   6810
         Index           =   6
         ItemData        =   "frmPlanets.frx":0461
         Left            =   -74400
         List            =   "frmPlanets.frx":0468
         MultiSelect     =   2  'Extended
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   720
         Width           =   6255
      End
      Begin VB.ListBox listInfo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   6585
         Index           =   5
         ItemData        =   "frmPlanets.frx":0476
         Left            =   -74400
         List            =   "frmPlanets.frx":047D
         MultiSelect     =   2  'Extended
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   720
         Width           =   6255
      End
      Begin VB.ListBox listInfo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4785
         Index           =   4
         ItemData        =   "frmPlanets.frx":048B
         Left            =   -74400
         List            =   "frmPlanets.frx":0492
         MultiSelect     =   2  'Extended
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   720
         Width           =   9015
      End
      Begin VB.ListBox listInfo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4110
         Index           =   3
         ItemData        =   "frmPlanets.frx":04A0
         Left            =   -74400
         List            =   "frmPlanets.frx":04A7
         MultiSelect     =   2  'Extended
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   720
         Width           =   9015
      End
      Begin VB.ListBox listInfo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4110
         Index           =   2
         ItemData        =   "frmPlanets.frx":04B5
         Left            =   -74400
         List            =   "frmPlanets.frx":04BC
         MultiSelect     =   2  'Extended
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   720
         Width           =   9015
      End
      Begin VB.ListBox listInfo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4110
         Index           =   1
         ItemData        =   "frmPlanets.frx":04CA
         Left            =   -74400
         List            =   "frmPlanets.frx":04D1
         MultiSelect     =   2  'Extended
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   720
         Width           =   9015
      End
      Begin VB.ListBox listInfo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4110
         Index           =   0
         ItemData        =   "frmPlanets.frx":04DF
         Left            =   600
         List            =   "frmPlanets.frx":04E1
         MultiSelect     =   2  'Extended
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   720
         Width           =   9015
      End
   End
   Begin VB.CheckBox chkGrootstePrecisie 
      Caption         =   "Greatest Precision"
      Height          =   255
      Left            =   3720
      TabIndex        =   25
      Top             =   840
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.Timer tmrInterval 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8760
      Top             =   7560
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Working objects"
      ForeColor       =   &H00000000&
      Height          =   3615
      Left            =   9720
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   3375
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Frame1"
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   60
         TabIndex        =   21
         Top             =   720
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   60
         TabIndex        =   20
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Check1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   60
         TabIndex        =   19
         Top             =   1680
         Width           =   1875
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmPlanets.frx":04E3
         Left            =   120
         List            =   "frmPlanets.frx":04E5
         TabIndex        =   18
         Text            =   "Combo1"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Command1"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label1"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   180
         TabIndex        =   23
         Top             =   2760
         Width           =   1815
      End
   End
   Begin VB.CommandButton ComputeButton 
      BackColor       =   &H00C0C0C0&
      Caption         =   "COMPUTE"
      Default         =   -1  'True
      Height          =   315
      Left            =   7320
      TabIndex        =   9
      Top             =   780
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Set Date && Time To"
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   7380
      TabIndex        =   15
      Top             =   60
      Width           =   1695
      Begin VB.CommandButton SetNowButton 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Now"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Time"
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   3120
      TabIndex        =   11
      Top             =   60
      Width           =   4215
      Begin VB.TextBox Hrs 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Min 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   600
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Sec 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Text            =   "00"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Set00HrButton 
         BackColor       =   &H00C0C0C0&
         Caption         =   "00 Hr"
         Height          =   315
         Left            =   2700
         TabIndex        =   6
         Top             =   240
         Width           =   675
      End
      Begin VB.CommandButton SetNoonButton 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Noon"
         Height          =   315
         Left            =   3420
         TabIndex        =   7
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Set to"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   14
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.Frame DateTime1Frame 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Date"
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   2895
      Begin VB.TextBox Year 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   2
         Text            =   "2001"
         Top             =   240
         Width           =   555
      End
      Begin VB.ComboBox MonthSelect 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmPlanets.frx":04E7
         Left            =   720
         List            =   "frmPlanets.frx":04E9
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   795
      End
      Begin VB.ComboBox DaySelect 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmPlanets.frx":04EB
         Left            =   120
         List            =   "frmPlanets.frx":04ED
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar statusBar2 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   24
      Top             =   9135
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20628
            MinWidth        =   176
            Text            =   "Local Star Time"
            TextSave        =   "Local Star Time"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1614
            MinWidth        =   1605
            TextSave        =   "28/02/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   291
            MinWidth        =   282
         EndProperty
      EndProperty
   End
   Begin VB.Image imgJup 
      Height          =   5010
      Left            =   1200
      Picture         =   "frmPlanets.frx":04EF
      Top             =   2160
      Width           =   5220
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings"
      End
      Begin VB.Menu mnuEmpty2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearTmp 
         Caption         =   "&Clear tempdir."
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuAlmanac 
         Caption         =   "&Almanac"
      End
      Begin VB.Menu mnuSterrenkaart 
         Caption         =   "&StarChart"
      End
      Begin VB.Menu mnuJupiter 
         Caption         =   "&Jupiter"
         Begin VB.Menu mnuJupiterMoonsPhenomena 
            Caption         =   "&Phenomena"
         End
         Begin VB.Menu mnuJupiterMoons 
            Caption         =   "&Movements"
         End
         Begin VB.Menu mnuJupiterDiagram 
            Caption         =   "&Diagram"
         End
         Begin VB.Menu mnuEmpty1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuJupiterMeridiaan 
            Caption         =   "&Meridians"
         End
      End
      Begin VB.Menu mnuSaturnus 
         Caption         =   "&Saturn"
         Begin VB.Menu mnuSaturnusMoonsPhenomena 
            Caption         =   "&Phenomena"
         End
         Begin VB.Menu mnuSaturnusMoons 
            Caption         =   "&Movements"
         End
         Begin VB.Menu mnuSaturnusDiagram 
            Caption         =   "&Diagram"
         End
      End
      Begin VB.Menu mnuEmpty 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTabelen 
         Caption         =   "&Tabels"
      End
      Begin VB.Menu mnuSurvey 
         Caption         =   "&Survey "
      End
      Begin VB.Menu mnuEphemerides 
         Caption         =   "&Ephemerides"
      End
      Begin VB.Menu mnuEclipse 
         Caption         =   "E&clipse"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "&Info"
      End
   End
End
Attribute VB_Name = "frmPlanets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const MERGEPAINT = &HBB0226
Private Const SRCAND = &H8800C6
Private Const SRCCOPY = &HCC0020
Private schaalJup As Double
Private Sub schrijfMaan(ByRef pl As tPlaneet_Maan)
Dim nPlaneet As Long
listInfo(nPlaneet).AddItem "Maan appearent: " + StrHMS_DMS(pl.RA_app * 180 / Pi, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(pl.Decl_app * 180 / Pi, 7, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Position Angle: " + StrHMS_DMS(pl.moonPhysData.x * 180 / Pi, 3, 0, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Parall. Angle : " + StrHMS_DMS(pl.parAngle * 180 / Pi, 3, 0, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Fase          : " + Format(pl.moonPhysData.k, "0.00000")
listInfo(nPlaneet).AddItem "Bright Limb   : " + StrHMS_DMS(pl.moonPhysData.x * 180 / Pi, 3, 0, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Terminator    : " + StrHMS_DMS(pl.moonPhysData.T * 180 / Pi, 1, 1, True, False, "g", 5)
listInfo(nPlaneet).AddItem "Libration in l: " + StrHMS_DMS(-pl.moonPhysData.L * 180 / Pi, 1, 1, True, False, "g", 4)
listInfo(nPlaneet).AddItem "Libration in b: " + StrHMS_DMS(pl.moonPhysData.B * 180 / Pi, 1, 1, True, False, "g", 4)
listInfo(nPlaneet).AddItem "Distance      : " + Format(pl.sGeo.r, "000000.00 km")
listInfo(nPlaneet).AddItem "Sterrenbeeld  : " + pl.sterbeeld
If pl.riseSet.Rise < 0 Then
    listInfo(nPlaneet).AddItem "Opkomst       : ------"
Else
    listInfo(nPlaneet).AddItem "Opkomst       : " + StrHMS_DMS(pl.riseSet.Rise * 180 / Pi, 3, 0, False, False, "h", 2)
End If
If pl.riseSet.Transit < 0 Then
    listInfo(nPlaneet).AddItem "Doorgang      : ------"
Else
    listInfo(nPlaneet).AddItem "Doorgang      : " + StrHMS_DMS(pl.riseSet.Transit * 180 / Pi, 3, 0, False, False, "h", 2)
End If
If pl.riseSet.Setting < 0 Then
    listInfo(nPlaneet).AddItem "Ondergang     : ------"
Else
    listInfo(nPlaneet).AddItem "Ondergang     : " + StrHMS_DMS(pl.riseSet.Setting * 180 / Pi, 3, 0, False, False, "h", 2)
End If
End Sub
Private Sub schrijfZon(ByRef pl As tPlaneet_Zon)
Dim nPlaneet As Long
    nPlaneet = 1
    listInfo(nPlaneet).AddItem "Zon           : " + StrHMS_DMS(180 / Pi * pl.RA2000, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / Pi * pl.Decl2000, 7, 2, True, False, "g", 3)
    listInfo(nPlaneet).AddItem "Appearent     : " + StrHMS_DMS(180 / Pi * pl.RA_app, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / Pi * pl.Decl_app, 7, 2, True, False, "g", 3)
    listInfo(nPlaneet).AddItem "Azimuth/hoogte: " + StrHMS_DMS(180 / Pi * pl.Azimuth, 3, 0, False, False, "g", 3) + vbTab + StrHMS_DMS(180 / Pi * pl.Hoogte, 3, 0, True, False, "g", 3)
    If pl.RTS18.Rise < 0 Then
        listInfo(nPlaneet).AddItem "Opkomst   -18 : ------"
    Else
        listInfo(nPlaneet).AddItem "Opkomst   -18 : " + StrHMS_DMS(pl.RTS18.Rise * 180 / Pi, 3, 0, False, False, "h", 2)
    End If
    If pl.RTS6.Rise < 0 Then
        listInfo(nPlaneet).AddItem "Opkomst   - 6 : ------"
    Else
        listInfo(nPlaneet).AddItem "Opkomst   - 6 : " + StrHMS_DMS(pl.RTS6.Rise * 180 / Pi, 3, 0, False, False, "h", 2)
    End If
    If pl.RTS.Rise < 0 Then
        listInfo(nPlaneet).AddItem "Opkomst       : ------"
    Else
        listInfo(nPlaneet).AddItem "Opkomst       : " + StrHMS_DMS(pl.RTS.Rise * 180 / Pi, 3, 0, False, False, "h", 2)
    End If
    
    If pl.RTS.Transit < 0 Then
        listInfo(nPlaneet).AddItem "Doorgang      : ------"
    Else
        listInfo(nPlaneet).AddItem "Doorgang      : " + StrHMS_DMS(pl.RTS.Transit * 180 / Pi, 3, 0, False, False, "h", 2)
    End If
    
    If pl.RTS.Setting < 0 Then
        listInfo(nPlaneet).AddItem "Ondergang     : ------"
    Else
        listInfo(nPlaneet).AddItem "Ondergang     : " + StrHMS_DMS(pl.RTS.Setting * 180 / Pi, 3, 0, False, False, "h", 2)
    End If
    If pl.RTS6.Setting < 0 Then
        listInfo(nPlaneet).AddItem "Ondergang - 6 : ------"
    Else
        listInfo(nPlaneet).AddItem "Ondergang - 6 : " + StrHMS_DMS(pl.RTS6.Setting * 180 / Pi, 3, 0, False, False, "h", 2)
    End If
    If pl.RTS18.Setting < 0 Then
        listInfo(nPlaneet).AddItem "Ondergang -18 : ------"
    Else
        listInfo(nPlaneet).AddItem "Ondergang -18 : " + StrHMS_DMS(pl.RTS18.Setting * 180 / Pi, 3, 0, False, False, "h", 2)
    End If
    listInfo(nPlaneet).AddItem "Sterrenbeeld  : " + pl.sterbeeld
    listInfo(nPlaneet).AddItem "Parall. Angle : " + StrHMS_DMS(pl.parAngle * 180 / Pi, 3, 0, True, False, "g", 3)
    listInfo(nPlaneet).AddItem "p             : " + StrHMS_DMS(pl.SunPhysData.P * 180 / Pi, 1, 1, True, False, "g", 4)
    listInfo(nPlaneet).AddItem "b0            : " + StrHMS_DMS(pl.SunPhysData.b0 * 180 / Pi, 1, 1, True, False, "g", 4)
    listInfo(nPlaneet).AddItem "l0            : " + StrHMS_DMS(pl.SunPhysData.L0 * 180 / Pi, 1, 1, False, False, "g", 4)
End Sub
Private Sub schrijfVenus(ByRef pl As tPlaneet_Venus)
Dim nPlaneet As Long

nPlaneet = 3
listInfo(nPlaneet).AddItem "------------------------------------------------------------------"
listInfo(nPlaneet).AddItem "Venus         : " + StrHMS_DMS(180 / Pi * pl.RA2000, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / Pi * pl.Decl2000, 7, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Appearent     : " + StrHMS_DMS(180 / Pi * pl.RA_app, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / Pi * pl.Decl_app, 7, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Azimuth/hoogte: " + StrHMS_DMS(180 / Pi * pl.Azimuth, 3, 0, False, False, "g", 3) + vbTab + StrHMS_DMS(180 / Pi * pl.Hoogte, 3, 0, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Magnitude     : " + Format(pl.Magnitude, "0.0")
listInfo(nPlaneet).AddItem "Elongation    : " + StrHMS_DMS(pl.Elongation * 180 / Pi, 1, 1, True, False, "g", 4)
listInfo(nPlaneet).AddItem "Sterrenbeeld  : " + pl.sterbeeld
listInfo(nPlaneet).AddItem "Opkomst       : " + StrHMS_DMS(pl.RTS.Rise * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Doorgang      : " + StrHMS_DMS(pl.RTS.Transit * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Ondergang     : " + StrHMS_DMS(pl.RTS.Setting * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Phase         : " + Format(pl.phase, "0.000")
listInfo(nPlaneet).AddItem "PhaseAngle    : " + StrHMS_DMS(pl.PhaseAngle * 180 / Pi, 1, 1, True, False, "g", 4)
listInfo(nPlaneet).AddItem "Parall. Angle : " + StrHMS_DMS(pl.parAngle * 180 / Pi, 3, 0, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Diameter      : " + StrHMS_DMS(2 * pl.Semidiameter * SToR * RToD, 4, 1, True, False, "g", 4)

End Sub
Private Sub schrijfMercurius(ByRef pl As tPlaneet_Mercurius)
Dim nPlaneet As Long

nPlaneet = 2
listInfo(nPlaneet).AddItem "------------------------------------------------------------------"
listInfo(nPlaneet).AddItem "Mercurius     : " + StrHMS_DMS(180 / Pi * pl.RA2000, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / Pi * pl.Decl2000, 7, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Appearent     : " + StrHMS_DMS(180 / Pi * pl.RA_app, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / Pi * pl.Decl_app, 7, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Azimuth/hoogte: " + StrHMS_DMS(180 / Pi * pl.Azimuth, 3, 0, False, False, "g", 3) + vbTab + StrHMS_DMS(180 / Pi * pl.Hoogte, 3, 0, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Magnitude     : " + Format(pl.Magnitude, "0.0")
listInfo(nPlaneet).AddItem "Elongation    : " + StrHMS_DMS(pl.Elongation * 180 / Pi, 1, 1, True, False, "g", 4)
listInfo(nPlaneet).AddItem "Sterrenbeeld  : " + pl.sterbeeld
listInfo(nPlaneet).AddItem "Opkomst       : " + StrHMS_DMS(pl.RTS.Rise * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Doorgang      : " + StrHMS_DMS(pl.RTS.Transit * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Ondergang     : " + StrHMS_DMS(pl.RTS.Setting * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Phase         : " + Format(pl.phase, "0.000")
listInfo(nPlaneet).AddItem "PhaseAngle    : " + StrHMS_DMS(pl.PhaseAngle * 180 / Pi, 1, 1, True, False, "g", 4)
listInfo(nPlaneet).AddItem "Parall. Angle : " + StrHMS_DMS(pl.parAngle * 180 / Pi, 3, 0, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Diameter      : " + StrHMS_DMS(2 * pl.Semidiameter * SToR * RToD, 4, 1, True, False, "g", 4)

End Sub
Private Sub schrijfMars(ByRef pl As tPlaneet_Mars)
Dim nPlaneet As Long

nPlaneet = 4
listInfo(nPlaneet).AddItem "------------------------------------------------------------------"
listInfo(nPlaneet).AddItem "Mars          : " + StrHMS_DMS(180 / Pi * pl.RA2000, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / Pi * pl.Decl2000, 7, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Appearent     : " + StrHMS_DMS(180 / Pi * pl.RA_app, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / Pi * pl.Decl_app, 7, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Azimuth/hoogte: " + StrHMS_DMS(180 / Pi * pl.Azimuth, 3, 0, False, False, "g", 3) + vbTab + StrHMS_DMS(180 / Pi * pl.Hoogte, 3, 0, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Magnitude     : " + Format(pl.Magnitude, "0.0")
listInfo(nPlaneet).AddItem "Elongation    : " + StrHMS_DMS(pl.Elongation * 180 / Pi, 1, 1, True, False, "g", 4)
listInfo(nPlaneet).AddItem "Sterrenbeeld  : " + pl.sterbeeld
listInfo(nPlaneet).AddItem "Opkomst       : " + StrHMS_DMS(pl.RTS.Rise * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Doorgang      : " + StrHMS_DMS(pl.RTS.Transit * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Ondergang     : " + StrHMS_DMS(pl.RTS.Setting * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Phase         : " + Format(pl.phase, "0.000")
listInfo(nPlaneet).AddItem "PhaseAngle    : " + StrHMS_DMS(pl.PhaseAngle * 180 / Pi, 1, 1, True, False, "g", 4)
listInfo(nPlaneet).AddItem "Parall. Angle : " + StrHMS_DMS(pl.parAngle * 180 / Pi, 3, 0, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Diameter      : " + StrHMS_DMS(2 * pl.Semidiameter * SToR * RToD, 4, 1, True, False, "g", 4)
listInfo(nPlaneet).AddItem "DS            : " + StrHMS_DMS(pl.MarsPhysData.DS * 180 / Pi, 1, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "DE            : " + StrHMS_DMS(pl.MarsPhysData.DE * 180 / Pi, 1, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "P             : " + StrHMS_DMS(pl.MarsPhysData.P * 180 / Pi, 1, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "q             : " + StrHMS_DMS(pl.MarsPhysData.qq / 3600, 4, 2, True, False, "g", 4)
listInfo(nPlaneet).AddItem "Q             : " + StrHMS_DMS(pl.MarsPhysData.Q * 180 / Pi, 1, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Om            : " + StrHMS_DMS(pl.MarsPhysData.Om * 180 / Pi, 1, 2, False, False, "g", 3)

End Sub
Private Sub tekenJupiter(ByRef pl As tPlaneet_Jupiter)
Dim schaal As Double
Dim ii As Long
Dim vTemp As TVECTOR
Dim vsTemp As TVECTOR
Dim vMaan As TVECTOR
Dim vsMaan As TVECTOR
schaal = schaalJup
    picHidden.FillStyle = 1
    'picHidden.FillColor = RGB(255, 255, 255)
    'picHidden.Circle (picHidden.ScaleWidth / 2, picHidden.ScaleHeight / 2), Schaal, RGB(0, 0, 0), , , 1 / 1.071374
    picHidden.FillColor = RGB(0, 0, 0)
    picHidden.Line (0, 0)-(picHidden.ScaleWidth, picHidden.ScaleHeight), , BF
        picHidden.PaintPicture imgJup, picHidden.ScaleWidth / 2 - schaal, picHidden.ScaleHeight / 2 - schaal, schaal * 2, schaal * 2
 
    For ii = 1 To 4
        vTemp = pl.vMaan(ii)
        vsTemp = pl.vsMaan(ii)
        vsMaan = pl.vsMaan(ii)
        vMaan = pl.vMaan(ii)
        vTemp.Y = vTemp.Y * 1.071374
        vsTemp.Y = vsTemp.Y * 1.071374
    
    
    ' teken maantje
    If (vTemp.Z > 0) And (Abs((vTemp.x * vTemp.x + vTemp.Y * vTemp.Y)) < 1) Then
    '  {bedekt}
        'picHidden.PaintPicture imgBedekt, frmJupiterMoons.picHidden.ScaleWidth / 2 - vTemp(i).x * Schaal - 25, frmJupiterMoons.picHidden.ScaleHeight / 2 + vTemp(i).y * Schaal - 25, 50, 50, , , , , vbMergeCopy
        Call TekenCirkelKlein(picHidden, vTemp, schaal, schaal, RGB(150, 150, 150), 15)
    Else   '{niet bedekt}
        If (vTemp.Z > 0) And (Abs((vsMaan.x * vsMaan.x + vsMaan.Y * vsMaan.Y)) < 1) Then
        '    {verduisterd (schaduw jupiter op maantje)}
            'picHidden.PaintPicture imgVerduisterd, frmJupiterMoons.picHidden.ScaleWidth / 2 - vTemp(i).x * Schaal - 25, frmJupiterMoons.picHidden.ScaleHeight / 2 + vTemp(i).y * Schaal - 25, 50, 50, , , , , vbMergeCopy
            Call TekenCirkelKlein(picHidden, vTemp, schaal, schaal, RGB(0, 0, 255), 15)
        Else  '{niet bedekt en niet verduisterd}
            '{zichtbaar}
            Call TekenCirkelKlein(picHidden, vTemp, schaal, schaal, RGB(255, 255, 255), 15)
            'picHidden.PaintPicture imgMaan, frmJupiterMoons.picHidden.ScaleWidth / 2 - vTemp(i).x * Schaal - 25, frmJupiterMoons.picHidden.ScaleHeight / 2 + vTemp(i).y * Schaal - 25, 50, 50, , , , , vbMergeCopy
        End If
    End If
    If (vTemp.Z < 0) And (Abs((vsMaan.x * vsMaan.x + vsMaan.Y * vsMaan.Y)) < 1) Then
           '{schaduw op jupiter}
           'picHidden.PaintPicture imgSchaduw, frmJupiterMoons.picHidden.ScaleWidth / 2 - vSMaan(i).x * Schaal - 25, frmJupiterMoons.picHidden.ScaleHeight / 2 + vSMaan(i).y * Schaal - 25, 50, 50, , , , , vbMergeCopy
           Call TekenCirkelKlein(picHidden, vsMaan, schaal, schaal, RGB(0, 0, 0), 15)
    End If
    
Next

End Sub
Private Sub schrijfJupiter(ByRef pl As tPlaneet_Jupiter)

Dim nPlaneet As Long

nPlaneet = 5
listInfo(nPlaneet).AddItem "------------------------------------------------------------------"
listInfo(nPlaneet).AddItem "Jupiter       : " + StrHMS_DMS(180 / Pi * pl.RA2000, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / Pi * pl.Decl2000, 7, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Appearent     : " + StrHMS_DMS(180 / Pi * pl.RA_app, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / Pi * pl.Decl_app, 7, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Azimuth/hoogte: " + StrHMS_DMS(180 / Pi * pl.Azimuth, 3, 0, False, False, "g", 3) + vbTab + StrHMS_DMS(180 / Pi * pl.Hoogte, 3, 0, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Magnitude     : " + Format(pl.Magnitude, "0.0")
listInfo(nPlaneet).AddItem "Elongation    : " + StrHMS_DMS(pl.Elongation * 180 / Pi, 1, 1, True, False, "g", 4)
listInfo(nPlaneet).AddItem "Sterrenbeeld  : " + pl.sterbeeld
listInfo(nPlaneet).AddItem "Opkomst       : " + StrHMS_DMS(pl.RTS.Rise * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Doorgang      : " + StrHMS_DMS(pl.RTS.Transit * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Ondergang     : " + StrHMS_DMS(pl.RTS.Setting * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Phase         : " + Format(pl.phase, "0.000")
listInfo(nPlaneet).AddItem "PhaseAngle    : " + StrHMS_DMS(pl.PhaseAngle * 180 / Pi, 1, 1, True, False, "g", 4)
listInfo(nPlaneet).AddItem "Parall. Angle : " + StrHMS_DMS(pl.parAngle * 180 / Pi, 3, 0, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Diameter      : " + StrHMS_DMS(2 * pl.Semidiameter * SToR * RToD, 4, 1, True, False, "g", 4)
listInfo(nPlaneet).AddItem "Polardiameter : " + StrHMS_DMS(2 * pl.PolarSemiDiameter * SToR * RToD, 4, 1, True, False, "g", 4)

listInfo(nPlaneet).AddItem "DS            : " + StrHMS_DMS(pl.JupiterPhysData.DS * 180 / Pi, 1, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "DE            : " + StrHMS_DMS(pl.JupiterPhysData.DE * 180 / Pi, 1, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Om1           : " + StrHMS_DMS(pl.JupiterPhysData.Om1 * 180 / Pi, 1, 2, False, False, "g", 3)
listInfo(nPlaneet).AddItem "Om2           : " + StrHMS_DMS(pl.JupiterPhysData.Om2 * 180 / Pi, 1, 2, False, False, "g", 3)
listInfo(nPlaneet).AddItem "P             : " + StrHMS_DMS(pl.JupiterPhysData.P * 180 / Pi, 1, 2, True, False, "g", 3)
Dim MoonName
Dim sText As String
Dim ii As Long
MoonName = Array("", "Io       ", "Europa   ", "Ganymedes", "Callisto ")
sText = ""
    For ii = 1 To 4
        sText = "M " + MoonName(ii) + " " + FormatX(pl.vMaan(ii).x, "##0.00000") + " " + _
                        FormatX(pl.vMaan(ii).Y, "##0.00000") + vbTab + FormatX(pl.vMaan(ii).Z, "##0.00000") + " "
        
        sText = sText + " " + FormatX(Sqr(pl.vMaan(ii).x * pl.vMaan(ii).x + pl.vMaan(ii).Y * pl.vMaan(ii).Y), "##0.00000")
        lstMoonsJup.AddItem (sText)
        sText = "S " + MoonName(ii) + " " + FormatX(pl.vsMaan(ii).x, "##0.00000") + " " + _
                        FormatX(pl.vsMaan(ii).Y, "##0.00000") + vbTab + FormatX(pl.vsMaan(ii).Z, "##0.00000") + vbTab
        
        sText = sText + " " + FormatX(Sqr(pl.vsMaan(ii).x * pl.vsMaan(ii).x + pl.vsMaan(ii).Y * pl.vsMaan(ii).Y), "##0.00000")
        lstMoonsJup.AddItem (sText)
        If pl.situatieMaan(ii, 1) <> "" Then
            lstMoonsJup.AddItem (pl.situatieMaan(ii, 1))
        End If
        If pl.situatieMaan(ii, 2) <> "" Then
            lstMoonsJup.AddItem (pl.situatieMaan(ii, 2))
        End If
    Next
End Sub
Private Sub schrijfSaturnus(ByRef pl As tPlaneet_Saturnus)
Dim nPlaneet As Long

nPlaneet = 6
listInfo(nPlaneet).AddItem "------------------------------------------------------------------"
listInfo(nPlaneet).AddItem "Saturnus      : " + StrHMS_DMS(180 / Pi * pl.RA2000, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / Pi * pl.Decl2000, 7, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Appearent     : " + StrHMS_DMS(180 / Pi * pl.RA_app, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / Pi * pl.Decl_app, 7, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Azimuth/hoogte: " + StrHMS_DMS(180 / Pi * pl.Azimuth, 3, 0, False, False, "g", 3) + vbTab + StrHMS_DMS(180 / Pi * pl.Hoogte, 3, 0, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Magnitude     : " + Format(pl.Magnitude, "0.0")
listInfo(nPlaneet).AddItem "Elongation    : " + StrHMS_DMS(pl.Elongation * 180 / Pi, 1, 1, True, False, "g", 4)
listInfo(nPlaneet).AddItem "Sterrenbeeld  : " + pl.sterbeeld
listInfo(nPlaneet).AddItem "Opkomst       : " + StrHMS_DMS(pl.RTS.Rise * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Doorgang      : " + StrHMS_DMS(pl.RTS.Transit * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Ondergang     : " + StrHMS_DMS(pl.RTS.Setting * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Phase         : " + Format(pl.phase, "0.000")
listInfo(nPlaneet).AddItem "PhaseAngle    : " + StrHMS_DMS(pl.PhaseAngle * 180 / Pi, 1, 1, True, False, "g", 4)
listInfo(nPlaneet).AddItem "Parall. Angle : " + StrHMS_DMS(pl.parAngle * 180 / Pi, 3, 0, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Diameter      : " + StrHMS_DMS(2 * pl.Semidiameter * SToR * RToD, 4, 1, True, False, "g", 4)
listInfo(nPlaneet).AddItem "Polardiameter : " + StrHMS_DMS(2 * pl.PolarSemiDiameter * SToR * RToD, 4, 1, True, False, "g", 4)
      
listInfo(nPlaneet).AddItem "B             : " + StrHMS_DMS(pl.SaturnRingData.B * 180 / Pi, 1, 3, True, False, "g", 4)
listInfo(nPlaneet).AddItem "Bd            : " + StrHMS_DMS(pl.SaturnRingData.Bd * 180 / Pi, 1, 3, True, False, "g", 4)
listInfo(nPlaneet).AddItem "dU            : " + StrHMS_DMS(pl.SaturnRingData.DeltaU * 180 / Pi, 1, 3, True, False, "g", 4)
listInfo(nPlaneet).AddItem "P             : " + StrHMS_DMS(pl.SaturnRingData.P * 180 / Pi, 1, 3, True, False, "g", 4)
listInfo(nPlaneet).AddItem "B             : " + StrHMS_DMS(pl.AltSaturnRingData.B * 180 / Pi, 1, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "B1            : " + StrHMS_DMS(pl.AltSaturnRingData.b1 * 180 / Pi, 1, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "U             : " + StrHMS_DMS(pl.AltSaturnRingData.u * 180 / Pi, 1, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "U1            : " + StrHMS_DMS(pl.AltSaturnRingData.u1 * 180 / Pi, 1, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "P             : " + StrHMS_DMS(pl.AltSaturnRingData.P * 180 / Pi, 1, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "P1            : " + StrHMS_DMS(pl.AltSaturnRingData.P1 * 180 / Pi, 1, 2, True, False, "g", 3)

listInfo(nPlaneet).AddItem "A: " & vbTab & "Axis" & vbTab & "ioAxis" & vbTab & "oiAxis" & vbTab & "iiAxis" & vbTab & "idAxis"
listInfo(nPlaneet).AddItem vbTab & StrHMS_DMS(pl.SaturnRingData.aAxis / 3600, 4, 1, True, False, "g", 3) & vbTab & _
     StrHMS_DMS(pl.SaturnRingData.ioaAxis / 3600, 4, 1, True, False, "g", 3) & vbTab & _
     StrHMS_DMS(pl.SaturnRingData.oiaAxis / 3600, 4, 1, True, False, "g", 3) & vbTab & _
     StrHMS_DMS(pl.SaturnRingData.iiaAxis / 3600, 4, 1, True, False, "g", 3) & vbTab & _
     StrHMS_DMS(pl.SaturnRingData.idaAxis / 3600, 4, 1, True, False, "g", 3)
listInfo(nPlaneet).AddItem "B: " & vbTab & "Axis" & vbTab & "ioAxis" & vbTab & "oiAxis" & vbTab & "iiAxis" & vbTab & "idAxis"
listInfo(nPlaneet).AddItem vbTab & StrHMS_DMS(pl.SaturnRingData.bAxis / 3600, 4, 1, True, False, "g", 3) & vbTab & _
     StrHMS_DMS(pl.SaturnRingData.iobAxis / 3600, 4, 1, True, False, "g", 3) & vbTab & _
     StrHMS_DMS(pl.SaturnRingData.oibAxis / 3600, 4, 1, True, False, "g", 3) & vbTab & _
     StrHMS_DMS(pl.SaturnRingData.iibAxis / 3600, 4, 1, True, False, "g", 3) & vbTab & _
     StrHMS_DMS(pl.SaturnRingData.idbAxis / 3600, 4, 1, True, False, "g", 3)
Dim MoonName
Dim sText
Dim ii
MoonName = Array("", "Mimas    ", "Enceladus", "Tethys   ", "Dione    ", _
                     "Rhea     ", "Titan    ", "Hyperion ", "Japetus  ")
sText = ""
    For ii = 1 To 8
        sText = "M " + MoonName(ii) + " " + FormatX(pl.satmanen(ii).x, "##0.00000") + " " + _
                        FormatX(pl.satmanen(ii).Y, "##0.00000") + " " + _
                        FormatX(pl.satmanen(ii).Z, "##0.00000") + vbTab
        
        sText = sText + FormatX(Sqr(pl.satmanen(ii).x * pl.satmanen(ii).x + pl.satmanen(ii).Y * pl.satmanen(ii).Y), "##0.00000")
        lstMoonsSat.AddItem (sText)
    Next
End Sub
Private Sub schrijfUranus(ByRef pl As tPlaneet_Uranus)
Dim nPlaneet As Long
nPlaneet = 7
listInfo(nPlaneet).AddItem "------------------------------------------------------------------"
listInfo(nPlaneet).AddItem "Uranus        : " + StrHMS_DMS(180 / Pi * pl.RA2000, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / Pi * pl.Decl2000, 7, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Appearent     : " + StrHMS_DMS(180 / Pi * pl.RA_app, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / Pi * pl.Decl_app, 7, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Azimuth/hoogte: " + StrHMS_DMS(180 / Pi * pl.Azimuth, 3, 0, False, False, "g", 3) + vbTab + StrHMS_DMS(180 / Pi * pl.Hoogte, 3, 0, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Magnitude     : " + Format(pl.Magnitude, "0.0")
listInfo(nPlaneet).AddItem "Elongation    : " + StrHMS_DMS(pl.Elongation * 180 / Pi, 1, 1, True, False, "g", 4)
listInfo(nPlaneet).AddItem "Sterrenbeeld  : " + pl.sterbeeld
listInfo(nPlaneet).AddItem "Opkomst       : " + StrHMS_DMS(pl.RTS.Rise * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Doorgang      : " + StrHMS_DMS(pl.RTS.Transit * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Ondergang     : " + StrHMS_DMS(pl.RTS.Setting * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Phase         : " + Format(pl.phase, "0.000")
listInfo(nPlaneet).AddItem "PhaseAngle    : " + StrHMS_DMS(pl.PhaseAngle * 180 / Pi, 1, 1, True, False, "g", 4)
listInfo(nPlaneet).AddItem "Parall. Angle : " + StrHMS_DMS(pl.parAngle * 180 / Pi, 3, 0, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Diameter      : " + StrHMS_DMS(2 * pl.Semidiameter * SToR * RToD, 4, 1, True, False, "g", 4)
End Sub
Private Sub schrijfNeptunus(ByRef pl As tPlaneet_Neptunus)
Dim nPlaneet As Long
nPlaneet = 8
listInfo(nPlaneet).AddItem "------------------------------------------------------------------"
listInfo(nPlaneet).AddItem "Neptunus      : " + StrHMS_DMS(180 / Pi * pl.RA2000, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / Pi * pl.Decl2000, 7, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Appearent     : " + StrHMS_DMS(180 / Pi * pl.RA_app, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / Pi * pl.Decl_app, 7, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Azimuth/hoogte: " + StrHMS_DMS(180 / Pi * pl.Azimuth, 3, 0, False, False, "g", 3) + vbTab + StrHMS_DMS(180 / Pi * pl.Hoogte, 3, 0, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Magnitude     : " + Format(pl.Magnitude, "0.0")
listInfo(nPlaneet).AddItem "Elongation    : " + StrHMS_DMS(pl.Elongation * 180 / Pi, 1, 1, True, False, "g", 4)
listInfo(nPlaneet).AddItem "Sterrenbeeld  : " + pl.sterbeeld
listInfo(nPlaneet).AddItem "Opkomst       : " + StrHMS_DMS(pl.RTS.Rise * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Doorgang      : " + StrHMS_DMS(pl.RTS.Transit * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Ondergang     : " + StrHMS_DMS(pl.RTS.Setting * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Phase         : " + Format(pl.phase, "0.000")
listInfo(nPlaneet).AddItem "PhaseAngle    : " + StrHMS_DMS(pl.PhaseAngle * 180 / Pi, 1, 1, True, False, "g", 4)
listInfo(nPlaneet).AddItem "Parall. Angle : " + StrHMS_DMS(pl.parAngle * 180 / Pi, 3, 0, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Diameter      : " + StrHMS_DMS(2 * pl.Semidiameter * SToR * RToD, 4, 1, True, False, "g", 4)
End Sub

Private Sub schrijfPluto(ByRef pl As tPlaneet_Pluto)
Dim nPlaneet As Long
nPlaneet = 9
listInfo(nPlaneet).AddItem "------------------------------------------------------------------"
listInfo(nPlaneet).AddItem "Neptunus      : " + StrHMS_DMS(180 / Pi * pl.RA2000, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / Pi * pl.Decl2000, 7, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Appearent     : " + StrHMS_DMS(180 / Pi * pl.RA_app, 7, 3, False, False, "h", 2) + vbTab + StrHMS_DMS(180 / Pi * pl.Decl_app, 7, 2, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Azimuth/hoogte: " + StrHMS_DMS(180 / Pi * pl.Azimuth, 3, 0, False, False, "g", 3) + vbTab + StrHMS_DMS(180 / Pi * pl.Hoogte, 3, 0, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Magnitude     : " + Format(pl.Magnitude, "0.0")
listInfo(nPlaneet).AddItem "Elongation    : " + StrHMS_DMS(pl.Elongation * 180 / Pi, 1, 1, True, False, "g", 4)
listInfo(nPlaneet).AddItem "Sterrenbeeld  : " + pl.sterbeeld
listInfo(nPlaneet).AddItem "Opkomst       : " + StrHMS_DMS(pl.RTS.Rise * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Doorgang      : " + StrHMS_DMS(pl.RTS.Transit * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Ondergang     : " + StrHMS_DMS(pl.RTS.Setting * 180 / Pi, 3, 0, False, False, "h", 2)
listInfo(nPlaneet).AddItem "Phase         : " + Format(pl.phase, "0.000")
listInfo(nPlaneet).AddItem "PhaseAngle    : " + StrHMS_DMS(pl.PhaseAngle * 180 / Pi, 1, 1, True, False, "g", 4)
listInfo(nPlaneet).AddItem "Parall. Angle : " + StrHMS_DMS(pl.parAngle * 180 / Pi, 3, 0, True, False, "g", 3)
listInfo(nPlaneet).AddItem "Diameter      : " + StrHMS_DMS(2 * pl.Semidiameter * SToR * RToD, 4, 1, True, False, "g", 4)
End Sub

Private Sub calcu_Obl(T As Double, ByRef obl As tPlaneet_Obl)
Dim sLatitude As String, sLongitude As String
With obl
    .deltaT = ApproxDeltaT(T)
    Call NutationConst(T, .NutLon, .NutObl)
    .obl = Obliquity(T)
    
    Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Latitude", _
                sLatitude)
    .ObsLat = Val(sLatitude) * Pi / 180
    Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Longitude", _
                sLongitude)
    .ObsLon = Val(sLongitude) * Pi / 180
    .LAST = SiderealTime(T) + .NutLon * Cos(.obl) - .ObsLon
    Call ObserverCoord(.ObsLat, 0, .RhoCosPhi, .RhoSinPhi)
End With
End Sub
Private Sub calcu_RTS(Planet As Long, T0 As Double, height As Double, ByRef RTS As tRiseSetTran)
Dim alg As tPlaneet_Obl
Dim t0min1 As Double, t0plus1 As Double
Dim sAarde As TSVECTOR, sGeo As TSVECTOR, sHelio As TSVECTOR
Dim RA As Double, Decl As Double, RA1 As Double, Decl1 As Double, RA2 As Double, Decl2 As Double

If Planet = 0 Then
   ' bepalen opkomst e.d.
    Call calcu_Obl(T0, alg)
    Call PlanetPosHi(0, T0 - 1 / 36525, sAarde, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(sHelio, sAarde, sGeo)
    Call EclToEqu(sGeo.L, sGeo.B, alg.obl, RA1, Decl1)
    
    Call PlanetPosHi(0, T0, sAarde, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(sHelio, sAarde, sGeo)
    Call EclToEqu(sGeo.L, sGeo.B, alg.obl, RA, Decl)
    
    Call PlanetPosHi(0, T0 + 1 / 36525, sAarde, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(sHelio, sAarde, sGeo)
    Call EclToEqu(sGeo.L, sGeo.B, alg.obl, RA2, Decl2)
Else
    t0min1 = T0 - 1 / 36525
    ' bepalen opkomst e.d.
    Call calcu_Obl(t0min1, alg)
    Call PlanetPosHi(0, t0min1, sAarde, chkGrootstePrecisie.value = 0)
    Call PlanetPosHi(Planet, t0min1, sHelio, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(sHelio, sAarde, sGeo)
    Call PlanetPosHi(Planet, t0min1 - sGeo.r * LightTimeConst, sHelio, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(sHelio, sAarde, sGeo)
    Call EclToEqu(sGeo.L + alg.NutLon, sGeo.B, alg.obl + alg.NutObl, RA1, Decl1)
    Call Aberration(t0min1, alg.obl, FK5System, RA1, Decl1)
    
    ' bepalen opkomst e.d.
    t0plus1 = T0 + 1 / 36525
    Call calcu_Obl(t0plus1, alg)
    Call PlanetPosHi(0, t0plus1, sAarde, chkGrootstePrecisie.value = 0)
    Call PlanetPosHi(Planet, t0plus1, sHelio, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(sHelio, sAarde, sGeo)
    Call PlanetPosHi(Planet, t0plus1 - sGeo.r * LightTimeConst, sHelio, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(sHelio, sAarde, sGeo)
    Call EclToEqu(sGeo.L + alg.NutLon, sGeo.B, alg.obl + alg.NutObl, RA2, Decl2)
    Call Aberration(t0plus1, alg.obl, FK5System, RA2, Decl2)
    
    Call calcu_Obl(T0, alg)
    Call PlanetPosHi(0, T0, sAarde, chkGrootstePrecisie.value = 0)
    Call PlanetPosHi(Planet, T0, sHelio, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(sHelio, sAarde, sGeo)
    Call PlanetPosHi(Planet, T0 - sGeo.r * LightTimeConst, sHelio, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(sHelio, sAarde, sGeo)
    Call EclToEqu(sGeo.L + alg.NutLon, sGeo.B, alg.obl + alg.NutObl, RA, Decl)
    Call Aberration(T0, alg.obl, FK5System, RA, Decl)
End If
   
    Call riseSet(T0, alg.deltaT, RA1, Decl1, RA, Decl, RA2, Decl2, height, alg.ObsLon, alg.ObsLat, RTS)
    If RTS.flags > 0 Then
       RTS.Rise = -1
       RTS.Setting = -1
    End If
End Sub
Private Sub calcu_algemeen(ByRef alg As tPlaneet_algemeen)
Dim dat As tDatum
Dim tt As Double
Dim sLatitude As String
Dim sLongitude As String

With alg
    Call Zomertijd_Wintertijd(Year, .JD_ZT, .JD_WT)
    dat.jj = Year
    dat.MM = MonthSelect.ListIndex + 1
    dat.DD = DaySelect
    tt = (Hrs + Min / 60 + Sec / 3600) / 24
    dat.DD = dat.DD + tt
    .T = JDToT(KalenderNaarJD(dat))
    .deltaT = ApproxDeltaT(.T)
    .T0 = (floor(.T * 36525 + 0.50001) - 0.5) / 36525 + TijdCorrectie(KalenderNaarJD(dat) + 0.2, .JD_ZT, .JD_WT) / 36525#
    .DtofUT = .T0 + secToT * .deltaT
    .T = .T + TijdCorrectie(KalenderNaarJD(dat) + 0.2, .JD_ZT, .JD_WT) / 36525# + .deltaT * secToT
    Call NutationConst(.T, .NutLon, .NutObl)
    .obl = Obliquity(.T)
    
    Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Latitude", _
                sLatitude)
    .ObsLat = Val(sLatitude) * Pi / 180
    Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Longitude", _
                sLongitude)
    .ObsLon = Val(sLongitude) * Pi / 180
    .LAST = SiderealTime(.T) + .NutLon * Cos(.obl) - .ObsLon
    Call ObserverCoord(.ObsLat, .height, .RhoCosPhi, .RhoSinPhi)
End With
End Sub
Private Sub calcu_zon(ByRef alg As tPlaneet_algemeen, ByRef zon As tPlaneet_Zon)
'======================== ZON ========================
    Dim sAarde As TSVECTOR, sHelio As TSVECTOR, sGeo As TSVECTOR
    Dim RA1 As Double, Decl1 As Double
    Dim RA As Double, Decl As Double
    Dim RA2 As Double, Decl2 As Double

With zon
    .sHelio.L = 0: .sHelio.B = 0: .sHelio.r = 0
    Call PlanetPosHi(0, alg.T, .sAarde, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(.sHelio, .sAarde, .sGeo)
    Call PlanetPosHi(0, alg.T - .sGeo.r * LightTimeConst, .sAarde, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(.sHelio, .sAarde, .sGeo)
    Call EclToEqu(.sGeo.L, .sGeo.B, alg.obl, .RA2000, .Decl2000)
    Call SunPhysEphemeris(alg.T, .sGeo.L, alg.obl, alg.NutLon, .SunPhysData)
    .C = CarringtonRotation(TToJD(alg.T))
    .JDOfCarr = JDOfCarringtonRotation(.C)
    .parAngle = ParallacticAngle(.RA2000, .Decl2000, alg.ObsLat, alg.LAST)
    Call SterBld(.RA2000, .Decl2000, 0#, .sterbeeld)

    Call ConvertVSOP_FK5(alg.T, .sGeo.L, .sGeo.B)
    Call EclToEqu(.sGeo.L + alg.NutLon, .sGeo.B, alg.obl + alg.NutObl, .RA_app, .Decl_app)
    Call Aberration(alg.T, alg.obl, FK5System, .RA_app, .Decl_app)
    Call EquToHor(.RA_app, .Decl_app, alg.LAST, alg.ObsLat, .Azimuth, .Hoogte)
        
    Call calcu_RTS(0, alg.T0, h0Sun, .RTS)
    Call calcu_RTS(0, alg.T0, -6 * DToR, .RTS6)
    Call calcu_RTS(0, alg.T0, -18 * DToR, .RTS18)
   
    
End With
End Sub
Private Sub calcu_maan(ByRef alg As tPlaneet_algemeen, ByRef maan As tPlaneet_Maan)
Dim RA As Double, Decl As Double
Dim sAarde As TVECTOR, sHelio As TVECTOR, sSun As TVECTOR
With maan
    If chkGrootstePrecisie = 0 Then
        Call modMoonPos.MoonPos(alg.T, .sMoon)
        Call EclToEqu(.sMoon.L, .sMoon.B, alg.obl, RA, Decl)
    Else
        Call Lune(TToJD(alg.T), RA, Decl, .dist, .dkm, .diam, .phase, .illum)
        Call Lune(TToJD(alg.T - .dist * LightTimeConst), RA, Decl, .dist, .dkm, .diam, .phase, .illum)
    
        RA = RA * Pi / 12
        Decl = Decl * Pi / 180
        'coordinaten zijn voor J2000. Omzetten naar huidige dag, en daarna appearent berekenen
        Call PrecessFK5(0, alg.T, RA, Decl)
    End If

    Call Nutation(alg.NutLon, alg.NutObl, alg.obl, RA, Decl)
    .RA_app = RA: .Decl_app = Decl
    Call EquToEcl(.RA_app, .Decl_app, alg.obl, .sGeo.L, .sGeo.B)
    If chkGrootstePrecisie = 0 Then
        .sGeo.r = .sMoon.r
    Else
        .sGeo.r = .dkm
    End If
    .sHelio.L = 0: .sHelio.B = 0: .sHelio.r = 0
Call PlanetPosHi(0, alg.T, .sAarde, chkGrootstePrecisie.value = 0)
Call HelioToGeo(.sHelio, .sAarde, .sSun)
Call PlanetPosHi(0, alg.T - .sSun.r * LightTimeConst, .sAarde, chkGrootstePrecisie.value = 0)
Call HelioToGeo(.sHelio, .sAarde, .sSun)

Call MoonPhysEphemeris(alg.T, .sGeo, .sSun, alg.obl, alg.NutLon, alg.NutObl, .moonPhysData)
.parAngle = ParallacticAngle(.RA_app, .Decl_app, alg.ObsLat, alg.LAST)
Call SterBld(.RA_app, .Decl_app, 0#, .sterbeeld)
.parallax = asin(EarthRadius / .sGeo.r)
.moonHeight = MoonSetHeight(.parallax)

Dim sMoon As TSVECTOR
Dim diam As Double, dist As Double, dkm As Double, phase As Double, illum As Double
Dim RA1 As Double, Decl1 As Double, dist1 As Double, dkm1 As Double
Dim RA2 As Double, Decl2 As Double, dist2 As Double, dkm2 As Double
If chkGrootstePrecisie = 0 Then
    Call modMoonPos.MoonPos(alg.T0, sMoon)
    Call EclToEqu(sMoon.L, sMoon.B, alg.obl, RA, Decl)
Else
    Call Lune(TToJD(alg.T0), RA, Decl, dist, dkm, diam, phase, illum)
    Call Lune(TToJD(alg.T0 - dist * LightTimeConst), RA, Decl, dist, dkm, diam, phase, illum)
    RA = RA * Pi / 12
    Decl = Decl * Pi / 180
    'coordinaten zijn voor J2000. Omzetten naar huidige dag, en daarna appearent berekenen
    Call PrecessFK5(0, alg.T0, RA, Decl)
End If

Call Nutation(alg.NutLon, alg.NutObl, alg.obl, RA, Decl)

If chkGrootstePrecisie = 0 Then
    Call modMoonPos.MoonPos(alg.T0 - 1 / 36525, sMoon)
    Call EclToEqu(sMoon.L, sMoon.B, alg.obl, RA1, Decl1)
Else
    Call Lune(TToJD(alg.T0 - 1 / 36525), RA1, Decl1, dist1, dkm1, diam, phase, illum)
    Call Lune(TToJD(alg.T0 - dist1 * LightTimeConst - 1 / 36525), RA1, Decl1, dist1, dkm1, diam, phase, illum)
    RA1 = RA1 * Pi / 12
    Decl1 = Decl1 * Pi / 180
    'coordinaten zijn voor J2000. Omzetten naar huidige dag, en daarna appearent berekenen
    Call PrecessFK5(0, alg.T0, RA1, Decl1)
End If
Call Nutation(alg.NutLon, alg.NutObl, alg.obl, RA1, Decl1)

If chkGrootstePrecisie = 0 Then
    Call modMoonPos.MoonPos(alg.T0 + 1 / 36525, sMoon)
    Call EclToEqu(sMoon.L, sMoon.B, alg.obl, RA2, Decl2)
Else
    Call Lune(TToJD(alg.T0 + 1 / 36525), RA2, Decl2, dist2, dkm2, diam, phase, illum)
    Call Lune(TToJD(alg.T0 - dist1 * LightTimeConst + 1 / 36525), RA2, Decl2, dist2, dkm2, diam, phase, illum)
    RA2 = RA2 * Pi / 12
    Decl2 = Decl2 * Pi / 180
    'coordinaten zijn voor J2000. Omzetten naar huidige dag, en daarna appearent berekenen
    Call PrecessFK5(0, alg.T0, RA2, Decl2)
End If

Call Nutation(alg.NutLon, alg.NutObl, alg.obl, RA2, Decl2)

Call riseSet(alg.T0, alg.deltaT, RA1, Decl1, RA, Decl, RA2, Decl2, .moonHeight, alg.ObsLon, alg.ObsLat, .riseSet)
End With
End Sub
Private Sub calcu_mercurius(ByRef alg As tPlaneet_algemeen, ByRef mercurius As tPlaneet_Mercurius)
Dim sAarde As TSVECTOR, sHelio As TSVECTOR, sGeo As TSVECTOR
Dim RA1 As Double, Decl1 As Double
Dim RA As Double, Decl As Double
Dim RA2 As Double, Decl2 As Double

With mercurius
'(INTERFACE_DATE, INTERFACE_TIME, Q1)
'ListInfo(nPlaneet).AddItem Q
    Call PlanetPosHi(0, alg.T, .sAarde, chkGrootstePrecisie.value = 0)
    Call PlanetPosHi(1, alg.T, .sHelio, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(.sHelio, .sAarde, .sGeo)
    Call PlanetPosHi(1, alg.T - .sGeo.r * LightTimeConst, .sHelio, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(.sHelio, .sAarde, .sGeo)
    Call EclToEqu(.sGeo.L, .sGeo.B, alg.obl, .RA2000, .Decl2000)
    Call ConvertVSOP_FK5(alg.T, .sGeo.L, .sGeo.B)
    Call EclToEqu(.sGeo.L + alg.NutLon, .sGeo.B, alg.obl + alg.NutObl, .RA_app, .Decl_app)
    Call Aberration(alg.T, alg.obl, FK5System, .RA_app, .Decl_app)
    Call EquToHor(.RA_app, .Decl_app, alg.LAST, alg.ObsLat, .Azimuth, .Hoogte)
    
    .phase = CalcPhase(.sHelio.r, .sAarde.r, .sGeo.r)
    .PhaseAngle = acos(2 * .phase - 1)
    .Elongation = CalcElongation(.sHelio.r, .sAarde.r, .sGeo.r)
    If modpi(.sHelio.L - .sGeo.L) > 0 Then .Elongation = -.Elongation
    .Magnitude = PlanetMagnitude(1, .sHelio.r, .sGeo.r, .PhaseAngle, 0, 0)
    .Semidiameter = PlanetSemiDiameter(1, .sGeo.r, PolarSemiDiameter)
    .parAngle = ParallacticAngle(.RA2000, .Decl2000, alg.ObsLat, alg.LAST)
    Call SterBld(.RA2000, .Decl2000, 0#, .sterbeeld)
    
    Call calcu_RTS(1, alg.T0, h0Planet, .RTS)

End With
End Sub
Private Sub calcu_venus(ByRef alg As tPlaneet_algemeen, ByRef venus As tPlaneet_Venus)
Dim sAarde As TSVECTOR, sHelio As TSVECTOR, sGeo As TSVECTOR
Dim RA1 As Double, Decl1 As Double
Dim RA As Double, Decl As Double
Dim RA2 As Double, Decl2 As Double

With venus
'(INTERFACE_DATE, INTERFACE_TIME, Q1)
'ListInfo(nPlaneet).AddItem Q
    Call PlanetPosHi(0, alg.T, .sAarde, chkGrootstePrecisie.value = 0)
    Call PlanetPosHi(2, alg.T, .sHelio, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(.sHelio, .sAarde, .sGeo)
    Call PlanetPosHi(2, alg.T - .sGeo.r * LightTimeConst, .sHelio, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(.sHelio, .sAarde, .sGeo)
    Call EclToEqu(.sGeo.L, .sGeo.B, alg.obl, .RA2000, .Decl2000)
    Call ConvertVSOP_FK5(alg.T, .sGeo.L, .sGeo.B)
    Call EclToEqu(.sGeo.L + alg.NutLon, .sGeo.B, alg.obl + alg.NutObl, .RA_app, .Decl_app)
    Call Aberration(alg.T, alg.obl, FK5System, .RA_app, .Decl_app)
    Call EquToHor(.RA_app, .Decl_app, alg.LAST, alg.ObsLat, .Azimuth, .Hoogte)
    
    .phase = CalcPhase(.sHelio.r, .sAarde.r, .sGeo.r)
    .PhaseAngle = acos(2 * .phase - 1)
    .Elongation = CalcElongation(.sHelio.r, .sAarde.r, .sGeo.r)
    If modpi(.sHelio.L - .sGeo.L) > 0 Then .Elongation = -.Elongation
    .Magnitude = PlanetMagnitude(2, .sHelio.r, .sGeo.r, .PhaseAngle, 0, 0)
    .Semidiameter = PlanetSemiDiameter(2, .sGeo.r, PolarSemiDiameter)
    .parAngle = ParallacticAngle(.RA_app, .Decl_app, alg.ObsLat, alg.LAST)
    Call SterBld(.RA2000, .Decl2000, 0#, .sterbeeld)
    
    Call calcu_RTS(2, alg.T0, h0Planet, .RTS)

End With
End Sub

Private Sub calcu_mars(ByRef alg As tPlaneet_algemeen, ByRef mars As tPlaneet_Mars)

Dim sAarde As TSVECTOR, sHelio As TSVECTOR, sGeo As TSVECTOR
Dim RA1 As Double, Decl1 As Double
Dim RA As Double, Decl As Double
Dim RA2 As Double, Decl2 As Double

With mars
    Call PlanetPosHi(0, alg.T, .sAarde, chkGrootstePrecisie.value = 0)
    Call PlanetPosHi(4, alg.T, .sHelio, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(.sHelio, .sAarde, .sGeo)
    Call PlanetPosHi(4, alg.T - .sGeo.r * LightTimeConst, .sHelio, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(.sHelio, .sAarde, .sGeo)
    Call EclToEqu(.sGeo.L, .sGeo.B, alg.obl, .RA2000, .Decl2000)
    .phase = CalcPhase(.sHelio.r, .sAarde.r, .sGeo.r)
    .PhaseAngle = acos(2 * .phase - 1)
    .Elongation = CalcElongation(.sHelio.r, .sAarde.r, .sGeo.r)
    If modpi(.sHelio.L - .sGeo.L) > 0 Then .Elongation = -.Elongation
    .Magnitude = PlanetMagnitude(4, .sHelio.r, .sGeo.r, .PhaseAngle, SaturnRingData.DeltaU, SaturnRingData.B)
    .Semidiameter = PlanetSemiDiameter(4, .sGeo.r, PolarSemiDiameter)
    .parAngle = ParallacticAngle(.RA2000, .Decl2000, alg.ObsLat, alg.LAST)
    Call ConvertVSOP_FK5(alg.T, .sGeo.L, .sGeo.B)
    Call EclToEqu(.sGeo.L + alg.NutLon, .sGeo.B, alg.obl + alg.NutObl, .RA_app, .Decl_app)
    Call Aberration(alg.T, alg.obl, FK5System, .RA_app, .Decl_app)
    Call EquToHor(.RA_app, .Decl_app, alg.LAST, alg.ObsLat, .Azimuth, .Hoogte)
    Call SterBld(.RA2000, .Decl2000, 0#, .sterbeeld)
    Call MarsPhysEphemeris(alg.T, .sHelio, .sAarde, .sGeo, _
                                alg.obl, alg.NutLon, alg.NutObl, _
                                .MarsPhysData)
    
    Call calcu_RTS(4, alg.T0, h0Planet, .RTS)
End With
End Sub
Private Sub calcu_jupiter(ByRef alg As tPlaneet_algemeen, ByRef jupiter As tPlaneet_Jupiter)

Dim sAarde As TSVECTOR, sHelio As TSVECTOR, sGeo As TSVECTOR
Dim RA1 As Double, Decl1 As Double
Dim RA As Double, Decl As Double
Dim RA2 As Double, Decl2 As Double

With jupiter
    Call PlanetPosHi(0, alg.T, .sAarde, chkGrootstePrecisie.value = 0)
    Call PlanetPosHi(5, alg.T, .sHelio, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(.sHelio, .sAarde, .sGeo)
    Call PlanetPosHi(5, alg.T - .sGeo.r * LightTimeConst, .sHelio, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(.sHelio, .sAarde, .sGeo)
    Call EclToEqu(.sGeo.L, .sGeo.B, alg.obl, .RA2000, .Decl2000)
    .phase = CalcPhase(.sHelio.r, .sAarde.r, .sGeo.r)
    .PhaseAngle = acos(2 * .phase - 1)
    .Elongation = CalcElongation(.sHelio.r, .sAarde.r, .sGeo.r)
    If modpi(.sHelio.L - .sGeo.L) > 0 Then .Elongation = -.Elongation
    .Magnitude = PlanetMagnitude(5, .sHelio.r, .sGeo.r, .PhaseAngle, SaturnRingData.DeltaU, SaturnRingData.B)
    .Semidiameter = PlanetSemiDiameter(5, .sGeo.r, PolarSemiDiameter)
    .parAngle = ParallacticAngle(.RA2000, .Decl2000, alg.ObsLat, alg.LAST)
    Call ConvertVSOP_FK5(alg.T, .sGeo.L, .sGeo.B)
    Call EclToEqu(.sGeo.L + alg.NutLon, .sGeo.B, alg.obl + alg.NutObl, .RA_app, .Decl_app)
    Call Aberration(alg.T, alg.obl, FK5System, .RA_app, .Decl_app)
    Call EquToHor(.RA_app, .Decl_app, alg.LAST, alg.ObsLat, .Azimuth, .Hoogte)
    Call SterBld(.RA2000, .Decl2000, 0#, .sterbeeld)
    
    Call JupiterPhysEphemeris(alg.T + alg.deltaT / 36525 / 86400, .sHelio, .sAarde, .sGeo, _
                                   alg.obl, alg.NutLon, alg.NutObl, _
                                  .JupiterPhysData)
              
              
    Call calcu_RTS(5, alg.T0, h0Planet, .RTS)
    
'    Call riseSet(alg.T0, alg.deltaT, RA1, Decl1, RA, Decl, RA2, Decl2, h0Planet, alg.ObsLon, alg.ObsLat, .RTS)

    Dim MoonName
    Dim tJup As Double, rJup As Double
    Dim ii As Long
    Dim vTemp As TVECTOR
    Dim vsTemp As TVECTOR
    MoonName = Array("", "Io       ", "Europa   ", "Ganymedes", "Callisto ")
    tJup = alg.T - ApproxDeltaT(alg.T) * secToT
    For ii = 1 To 4
        Call BerekenPositieMaan(ii, TToJD(tJup), False, .vMaan(ii), rJup)
        Call BerekenPositieMaan(ii, TToJD(tJup), True, .vsMaan(ii), rJup)
        vTemp = .vMaan(ii)
        vsTemp = .vsMaan(ii)
        vTemp.Y = vTemp.Y * 1.071374
        vsTemp.Y = vsTemp.Y * 1.071374
    
        If (vTemp.Z > 0) Then
            If (Abs((vTemp.x * vTemp.x + vTemp.Y * vTemp.Y)) < 1) Then
                .situatieMaan(ii, 1) = MoonName(ii) + " : is bedekt"
            End If
            If (Abs((vsTemp.x * vsTemp.x + vsTemp.Y * vsTemp.Y)) < 1) Then
                .situatieMaan(ii, 2) = MoonName(ii) + " : is verduisterd"
            End If
        Else
            If (Abs((vTemp.x * vTemp.x + vTemp.Y * vTemp.Y)) < 1) Then
                .situatieMaan(ii, 1) = MoonName(ii) + " : trekt voorlangs"
            End If
            If (Abs((vsTemp.x * vsTemp.x + vsTemp.Y * vsTemp.Y)) < 1) Then
                .situatieMaan(ii, 2) = MoonName(ii) + " : heeft schaduwovergang"
            End If
        End If
        
    Next
End With
End Sub

Private Sub calcu_saturnus(ByRef alg As tPlaneet_algemeen, ByRef saturnus As tPlaneet_Saturnus)
Dim sAarde As TSVECTOR, sHelio As TSVECTOR, sGeo As TSVECTOR
Dim RA1 As Double, Decl1 As Double
Dim RA As Double, Decl As Double
Dim RA2 As Double, Decl2 As Double

With saturnus
    Call PlanetPosHi(0, alg.T, .sAarde, chkGrootstePrecisie.value = 0)
    Call PlanetPosHi(6, alg.T, .sHelio, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(.sHelio, .sAarde, .sGeo)
    Call PlanetPosHi(6, alg.T - .sGeo.r * LightTimeConst, .sHelio, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(.sHelio, .sAarde, .sGeo)
    Call EclToEqu(.sGeo.L, .sGeo.B, alg.obl, .RA2000, .Decl2000)
    Call SaturnRing(alg.T, .sHelio, .sGeo, alg.obl, alg.NutLon, alg.NutObl, .SaturnRingData)
    Call AltSaturnRing(alg.T, .sHelio, .sGeo, alg.obl, alg.NutLon, alg.NutObl, .AltSaturnRingData)
    .phase = CalcPhase(.sHelio.r, .sAarde.r, .sGeo.r)
    .PhaseAngle = acos(2 * .phase - 1)
    .Elongation = CalcElongation(.sHelio.r, .sAarde.r, .sGeo.r)
    If modpi(.sHelio.L - .sGeo.L) > 0 Then .Elongation = -.Elongation
    .Magnitude = PlanetMagnitude(6, .sHelio.r, .sGeo.r, .PhaseAngle, .SaturnRingData.DeltaU, .SaturnRingData.B)
    .Semidiameter = PlanetSemiDiameter(6, .sGeo.r, .PolarSemiDiameter)
    Call CorrectSaturnSemiDiameter(.SaturnRingData.B, .PolarSemiDiameter)
    .parAngle = ParallacticAngle(.RA2000, .Decl2000, alg.ObsLat, alg.LAST)
    Call SterBld(.RA2000, .Decl2000, 0#, .sterbeeld)
    Call ConvertVSOP_FK5(alg.T, .sGeo.L, .sGeo.B)
    Call EclToEqu(.sGeo.L + alg.NutLon, .sGeo.B, alg.obl + alg.NutObl, .RA_app, .Decl_app)
    Call Aberration(alg.T, alg.obl, FK5System, .RA_app, .Decl_app)
    Call EquToHor(.RA_app, .Decl_app, alg.LAST, alg.ObsLat, .Azimuth, .Hoogte)
          
    Call calcu_RTS(6, alg.T0, h0Planet, .RTS)
    
    Call CalcSaturnMoons(TToJD(alg.T - secToT * ApproxDeltaT(alg.T)), .satmanen)
End With
End Sub

Private Sub calcu_uranus(ByRef alg As tPlaneet_algemeen, ByRef uranus As tPlaneet_Uranus)

Dim sAarde As TSVECTOR, sHelio As TSVECTOR, sGeo As TSVECTOR
Dim RA1 As Double, Decl1 As Double
Dim RA As Double, Decl As Double
Dim RA2 As Double, Decl2 As Double

With uranus
    Call PlanetPosHi(0, alg.T, .sAarde, chkGrootstePrecisie.value = 0)
    Call PlanetPosHi(7, alg.T, .sHelio, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(.sHelio, .sAarde, .sGeo)
    Call PlanetPosHi(7, alg.T - .sGeo.r * LightTimeConst, .sHelio, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(.sHelio, .sAarde, .sGeo)
    Call EclToEqu(.sGeo.L, .sGeo.B, alg.obl, .RA2000, .Decl2000)
    .phase = CalcPhase(.sHelio.r, .sAarde.r, .sGeo.r)
    .PhaseAngle = acos(2 * .phase - 1)
    .Elongation = CalcElongation(.sHelio.r, .sAarde.r, .sGeo.r)
    If modpi(.sHelio.L - .sGeo.L) > 0 Then .Elongation = -.Elongation
    .Magnitude = PlanetMagnitude(7, .sHelio.r, .sGeo.r, .PhaseAngle, SaturnRingData.DeltaU, SaturnRingData.B)
    .Semidiameter = PlanetSemiDiameter(7, .sGeo.r, PolarSemiDiameter)
    .parAngle = ParallacticAngle(.RA2000, .Decl2000, alg.ObsLat, alg.LAST)
    Call SterBld(.RA2000, .Decl2000, 0#, .sterbeeld)
    Call ConvertVSOP_FK5(alg.T, .sGeo.L, .sGeo.B)
    Call EclToEqu(.sGeo.L + alg.NutLon, .sGeo.B, alg.obl + alg.NutObl, .RA_app, .Decl_app)
    Call Aberration(alg.T, alg.obl, FK5System, .RA_app, .Decl_app)
    Call EquToHor(.RA_app, .Decl_app, alg.LAST, alg.ObsLat, .Azimuth, .Hoogte)
    Call calcu_RTS(7, alg.T0, h0Planet, .RTS)
End With
End Sub


Private Sub calcu_neptunus(ByRef alg As tPlaneet_algemeen, ByRef neptunus As tPlaneet_Neptunus)

Dim sAarde As TSVECTOR, sHelio As TSVECTOR, sGeo As TSVECTOR
Dim RA1 As Double, Decl1 As Double
Dim RA As Double, Decl As Double
Dim RA2 As Double, Decl2 As Double

With neptunus
    Call PlanetPosHi(0, alg.T, .sAarde, chkGrootstePrecisie.value = 0)
    Call PlanetPosHi(8, alg.T, .sHelio, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(.sHelio, .sAarde, .sGeo)
    Call PlanetPosHi(8, alg.T - .sGeo.r * LightTimeConst, .sHelio, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(.sHelio, .sAarde, .sGeo)
    Call EclToEqu(.sGeo.L, .sGeo.B, alg.obl, .RA2000, .Decl2000)
    .phase = CalcPhase(.sHelio.r, .sAarde.r, .sGeo.r)
    .PhaseAngle = acos(2 * .phase - 1)
    .Elongation = CalcElongation(.sHelio.r, .sAarde.r, .sGeo.r)
    If modpi(.sHelio.L - .sGeo.L) > 0 Then .Elongation = -.Elongation
    .Magnitude = PlanetMagnitude(8, .sHelio.r, .sGeo.r, .PhaseAngle, SaturnRingData.DeltaU, SaturnRingData.B)
    .Semidiameter = PlanetSemiDiameter(8, .sGeo.r, PolarSemiDiameter)
    .parAngle = ParallacticAngle(.RA2000, .Decl2000, alg.ObsLat, alg.LAST)
    Call SterBld(.RA2000, .Decl2000, 0#, .sterbeeld)
    Call ConvertVSOP_FK5(alg.T, .sGeo.L, .sGeo.B)
    Call EclToEqu(.sGeo.L + alg.NutLon, .sGeo.B, alg.obl + alg.NutObl, .RA_app, .Decl_app)
    Call Aberration(alg.T, alg.obl, FK5System, .RA_app, .Decl_app)
    Call EquToHor(.RA_app, .Decl_app, alg.LAST, alg.ObsLat, .Azimuth, .Hoogte)
    Call calcu_RTS(8, alg.T0, h0Planet, .RTS)
End With
End Sub

Private Sub calcu_pluto(ByRef alg As tPlaneet_algemeen, ByRef pluto As tPlaneet_Pluto)
Dim sAarde As TSVECTOR, sHelio As TSVECTOR, sGeo As TSVECTOR
Dim RA1 As Double, Decl1 As Double
Dim RA As Double, Decl As Double
Dim RA2 As Double, Decl2 As Double
Dim dist As Double

With pluto
    Dim TAarde As TVECTOR
    Dim sZon As TSVECTOR
    Dim TPluto As TVECTOR
    .sHelio.L = 0: .sHelio.B = 0: .sHelio.r = 0
    Call PlanetPosHi(0, alg.T, .sAarde, chkGrootstePrecisie.value = 0)
    Call HelioToGeo(.sHelio, .sAarde, .sGeo)
    Call SphToRect(.sGeo, TAarde)
    Call EclToEqu(.sGeo.L, .sGeo.B, alg.obl, .RA2000, .Decl2000)
    ' Call Reduction2000(0, RA, Decl)
    'coordinaten omzetten naar J2000
    Call PrecessFK5(alg.T, 0, .RA2000, .Decl2000)
    Call EquToEcl(.RA2000, .Decl2000, Obliquity(0), .sGeo.L, .sGeo.B)
    Call SphToRect(.sGeo, TAarde)
    Call EclVSOP2000_equFK52000(TAarde.x, TAarde.Y, TAarde.Z)
    Call RectToSph(TAarde, sZon)
    .sAarde = .sGeo
    
    Call PlanetPosHi(0, alg.T, .sAarde, chkGrootstePrecisie.value = 0)
    Call PlutoPos(alg.T, .sHelio)
    Call EclToRect(.sHelio, Obliquity(0), TPluto)
    dist = Sqr((TAarde.x + TPluto.x) * (TAarde.x + TPluto.x) + (TAarde.Y + TPluto.Y) * (TAarde.Y + TPluto.Y) + (TAarde.Z + TPluto.Z) * (TAarde.Z + TPluto.Z))
    Call PlutoPos(alg.T - dist * LightTimeConst, .sHelio)
    Call EclToRect(.sHelio, Obliquity(0), TPluto)
    dist = Sqr((TAarde.x + TPluto.x) * (TAarde.x + TPluto.x) + (TAarde.Y + TPluto.Y) * (TAarde.Y + TPluto.Y) + (TAarde.Z + TPluto.Z) * (TAarde.Z + TPluto.Z))
    .RA2000 = atan2(TPluto.Y + TAarde.Y, TPluto.x + TAarde.x)
    If .RA2000 < 0 Then
        .RA2000 = .RA2000 + Pi2
    End If
    .Decl2000 = asin((TPluto.Z + TAarde.Z) / dist)
    .phase = CalcPhase(.sHelio.r, .sAarde.r, dist)
    .PhaseAngle = acos(2 * .phase - 1)
    .Magnitude = PlanetMagnitude(9, .sHelio.r, dist, .PhaseAngle, SaturnRingData.DeltaU, SaturnRingData.B)
    .Semidiameter = PlanetSemiDiameter(9, dist, PolarSemiDiameter)
    .parAngle = ParallacticAngle(.RA2000, .Decl2000, alg.ObsLat, alg.LAST)
    Call SterBld(.RA2000, .Decl2000, 0#, .sterbeeld)
    Call EquToEcl(.RA2000, .Decl2000, alg.obl, .sGeo.L, .sGeo.B)
    Call ConvertVSOP_FK5(alg.T, .sGeo.L, .sGeo.B)
    Call EclToEqu(.sGeo.L + alg.NutLon, .sGeo.B, alg.obl + alg.NutObl, .RA_app, .Decl_app)
    Call Aberration(alg.T, alg.obl, FK5System, .RA_app, .Decl_app)
    Call EquToHor(.RA_app, .Decl_app, alg.LAST, alg.ObsLat, .Azimuth, .Hoogte)
    
    Call riseSet(alg.T0, alg.deltaT, .RA_app, .Decl_app, .RA_app, .Decl_app, .RA_app, .Decl_app, h0Planet, alg.ObsLon, alg.ObsLat, .RTS)
    'DMO, 03-07-2008 indien flags aangaf dat
    If .RTS.flags > 0 Then
       .RTS.Rise = -1
       .RTS.Setting = -1
    End If
End With
End Sub
Private Sub cmdAnimate_Click()
    If TimerAnimate.Enabled Then
        Me.TimerAnimate.Interval = 0
        Me.TimerAnimate.Enabled = False
        Me.cmdAnimate.Caption = "ANIMATE"
    Else
        Dim keuze As String
        keuze = InputBox("Verversen om (in seconden)", "Verversen", 1)
        If keuze <> "" Then
            Me.TimerAnimate.Interval = Val(keuze) * 1000
            Me.TimerAnimate.Enabled = True
            Me.TimerAnimate.Tag = InputBox("Stapgrootte (in seconden)", "Verversen", 1)
            Me.cmdAnimate.Caption = "Stop ANIMATE"
        End If
    End If
End Sub

Private Sub ComputeButton_Click()
    Call calcu
End Sub
Private Sub calcu(Optional nPlaneet As Integer = -1)

Dim alg As tPlaneet_algemeen
Dim maan As tPlaneet_Maan

Dim zon As tPlaneet_Zon
Dim mercurius As tPlaneet_Mercurius
Dim venus As tPlaneet_Venus
Dim mars As tPlaneet_Mars
Dim jupiter As tPlaneet_Jupiter
Dim saturnus As tPlaneet_Saturnus
Dim uranus As tPlaneet_Uranus
Dim neptunus As tPlaneet_Neptunus
Dim pluto As tPlaneet_Pluto
Dim I As Long
Call calcu_algemeen(alg)
If nPlaneet = -1 Then
    For I = 0 To 9: listInfo(I).Clear: Next
Else
    listInfo(nPlaneet).Clear
End If
If nPlaneet = -1 Or nPlaneet = 0 Then
    Call calcu_maan(alg, maan)
    Call schrijfMaan(maan)
End If
If nPlaneet = -1 Or nPlaneet = 1 Then
    Call calcu_zon(alg, zon)
    Call schrijfZon(zon)
End If
If nPlaneet = -1 Or nPlaneet = 2 Then
    Call calcu_mercurius(alg, mercurius)
    Call schrijfMercurius(mercurius)
End If
If nPlaneet = -1 Or nPlaneet = 3 Then
    Call calcu_venus(alg, venus)
    Call schrijfVenus(venus)
End If
If nPlaneet = -1 Or nPlaneet = 4 Then
    Call calcu_mars(alg, mars)
    Call schrijfMars(mars)
End If
If nPlaneet = -1 Or nPlaneet = 5 Then
    Call calcu_jupiter(alg, jupiter)
    Me.lstMoonsJup.Clear
    Call schrijfJupiter(jupiter)
    Call tekenJupiter(jupiter)
    Call frmSaturnus.tekenSaturnusRingsMoons(picHiddenSat, TToJD(alg.T - secToT * ApproxDeltaT(alg.T)), 3.5, False)
End If
If nPlaneet = -1 Or nPlaneet = 6 Then
    Call calcu_saturnus(alg, saturnus)
    Me.lstMoonsSat.Clear
    Call schrijfSaturnus(saturnus)
End If
If nPlaneet = -1 Or nPlaneet = 7 Then
    Call calcu_uranus(alg, uranus)
    Call schrijfUranus(uranus)
End If
If nPlaneet = -1 Or nPlaneet = 8 Then
    Call calcu_neptunus(alg, neptunus)
    Call schrijfNeptunus(neptunus)
End If
If nPlaneet = -1 Or nPlaneet = 9 Then
    Call calcu_pluto(alg, pluto)
    Call schrijfPluto(pluto)
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Long
Dim sText As String
Dim nTab As Integer
If Me.ActiveControl.Name = "listInfo" Then
    If KeyCode = 67 And Shift = 2 Then
        For nTab = 0 To 9
       ' Me.List1.
        For I = 1 To Me.listInfo(nTab).ListCount
           sText = sText & vbCrLf & Me.listInfo(nTab).List(I - 1)
        Next
       Clipboard.Clear
       Clipboard.SetText (sText)
       Next
    End If
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
Static nTmpStap As Long
If KeyAscii = 43 Then '+
    If schaalJup < 500 Then schaalJup = schaalJup + 10
    ComputeButton_Click
ElseIf KeyAscii = 45 Then '-
    If schaalJup > 150 Then schaalJup = schaalJup - 10
    ComputeButton_Click
End If
End Sub
Private Sub Form_Load()
' What to do when this program starts up.
Set g_word = CreateObject("Word.Application")
#If FRANS Then
    Me.Caption = "Astronomie"
    Me.mnuFile.Caption = "&Fichier"
    Me.mnuSettings.Caption = "&Mise"
    Me.mnuOptions.Caption = "&Options"
    Me.mnuAlmanac.Caption = "&Almanach"
    Me.mnuSterrenkaart.Caption = "&Carte du Ciel"
    Me.mnuJupiter.Caption = "&Jupiter"
    Me.mnuSaturnus.Caption = "&Saturne"
    Me.mnuTabelen.Caption = "&Tableaux"
    Me.mnuHelp.Caption = "A&ide"
    Me.mnuJupiterDiagram.Caption = "Graphique Jupiter"
    Me.mnuJupiterMeridiaan.Caption = "Méridian de Jupiter"
    Me.mnuJupiterMoons.Caption = "Satellites de Jupiter"
    Me.mnuJupiterMoonsPhenomena.Caption = "Phénomène de satellites de Jupiter"
    Me.mnuSaturnusDiagram.Caption = "Graphique Saturne"
    Me.mnuSaturnusMoons.Caption = "Satellites de Saturne"
    Me.mnuSaturnusMoonsPhenomena.Caption = "Phénomène de satellites de Saturne"
    Me.mnuAbout.Caption = "Info"
    Me.mnuInfo.Caption = "Aide"
    Me.chkGrootstePrecisie.Caption = "Grande Précision"
    Me.DateTime1Frame.Caption = "Date"
    Me.Frame2.Caption = "Temps"
    Me.SetNoonButton.Caption = "Midi"
    Me.Frame3.Caption = "Changer temps"
    Me.SetNowButton.Caption = "à présent"
    Me.ComputeButton.Caption = "Calculer"
    Me.statusBar2.Panels(1).Text = "Local sidérale temps"
#End If
Dim Q, D, M, Date0

'PI = 4 * Atn(1)

For D = 1 To 31
    DaySelect.AddItem Right("  " & Trim(D) & " ", 4)
Next D
    DaySelect.ListIndex = 0

    Q = "JanFebMarAprMayJunJulAugSepOctNovDec"
For M = 1 To 12
    MonthSelect.AddItem " " & Mid(Q, 3 * (M - 1) + 1, 3) & " "
Next M
    MonthSelect.ListIndex = 0

   Q = Trim(DaySelect.Text) & " " & Trim(MonthSelect.Text) & " " & Year.Text
'If BCOption.Value = True Then Q = Q & " BC" Else Q = Q & " AD"

' Set initial default startup date
  SET_INTERFACE_DATE_AND_TIME_TO_NOW
 Dim I As Integer
For I = 0 To 9: Me.listInfo(I).Clear: Next
schaalJup = 100
  SSTabPlanets.Tab = 0
  startTimer

End Sub



Private Sub Form_Terminate()
' What to do upon shutting down this program.
  Unload Me
  On Error GoTo word_einde:
Dim s As String
s = g_word.Documents(1).Name
Set g_word = Nothing
Exit Sub

word_einde:
    g_word.Quit
    Resume Next
End Sub

Private Sub Hrs_GotFocus()
    Hrs.SelStart = 0
    Hrs.SelLength = Len(Hrs.Text)
End Sub

Private Sub Hrs_LostFocus()
Hrs.Text = Format(Val(Hrs.Text), "0#")
End Sub

Private Sub Min_GotFocus()
    Min.SelStart = 0
    Min.SelLength = Len(Min.Text)
End Sub

Private Sub Min_LostFocus()
Min.Text = Format(Val(Min.Text), "0#")
End Sub

Private Sub mnuAbout_Click()
Dim sAppInfo As String
sAppInfo = ShowFileInfo(App.Path & "\" & App.EXEName & ".exe")
 MsgBox _
        "Astronomisch programma ontwikkeld door Dominique Molenkamp." & vbNewLine & _
        "Version " & App.Major & "." & App.Minor & "." & App.Revision & "  Copyright © 2006-2013 (" _
        & sAppInfo & ")", _
        , "About " & App.Title
End Sub
Function ShowFileInfo(filespec As String) As String
On Error GoTo fout:
    Dim fs, F, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set F = fs.GetFile(filespec)
    s = F.DateCreated
    ShowFileInfo = Format(s, "dd-mm-yyyy")
    Exit Function
fout:
    ShowFileInfo = ""
End Function


Private Sub mnuAlmanac_Click()
frmAlmanac.Show vbModal, Me
End Sub

Private Sub mnuClearTmp_Click()
Dim objspecialfolder As New clsSpecialFolder
Dim sTmpDirectory As String
Dim sFile As String

On Error Resume Next
sTmpDirectory = objspecialfolder.TemporaryFolder
sFile = Dir(sTmpDirectory & "\*.*", vbNormal)
While sFile <> ""
    Kill sTmpDirectory & "\" & sFile
    sFile = Dir()
Wend
End Sub

Private Sub mnuEclipse_Click()
frmEclipse.Show vbModal, Me
End Sub

Private Sub mnuEphemerides_Click()
frmEphem.Show vbModal, Me
End Sub

Private Sub mnuInfo_Click()
frmInfo.Show vbModal
End Sub

Private Sub mnuJupiterDiagram_Click()
frmJupiterDiagram.Show vbModal
End Sub

Private Sub mnuJupiterMeridiaan_Click()
frmJupiterMeridiaan.Show vbModal, Me
End Sub

Private Sub mnuJupiterMoons_Click()
frmJupiterMoons.Show vbModal
End Sub

Private Sub mnuJupiterMoonsPhenomena_Click()
frmJupiterPhenomena.Show vbModal, Me
End Sub

Private Sub mnuSaturnusDiagram_Click()
frmSaturnusDiagram.Show vbModal, Me
End Sub

Private Sub mnuSaturnusMoonsPhenomena_Click()
frmSaturnusPhenomena.Show vbModal, Me
End Sub

Private Sub mnuSaturnusMoons_Click()
frmSaturnus.Show vbModal, Me
End Sub

Private Sub mnuSettings_Click()
    frmSettings.Show
End Sub

Private Sub mnuSterrenkaart_Click()
frmSterren.Show vbModal, Me
End Sub

Private Sub mnuSurvey_Click()
frmOverzicht.Show vbModal, Me
End Sub

Private Sub mnuTabelen_Click()
frmTabellen.Show vbModal, Me
End Sub

Private Sub Sec_GotFocus()
    Sec.SelStart = 0
    Sec.SelLength = Len(Sec.Text)
End Sub

Private Sub Sec_LostFocus()
Sec.Text = Format(Sec.Text, "#0.00")
End Sub

Private Sub TimerAnimate_Timer()
Dim dat As tDatum
Dim tt As Double
Dim stijd As String
If Me.TimerAnimate.Interval > 0 Then
    dat.jj = Year
    dat.MM = MonthSelect.ListIndex + 1
    dat.DD = DaySelect
    tt = (Hrs + Min / 60 + (Sec + Val(Me.TimerAnimate.Tag)) / 3600) / 24
    dat.DD = dat.DD + tt
    dat = JDNaarKalender(KalenderNaarJD(dat))
    Year = dat.jj
    MonthSelect.ListIndex = dat.MM - 1
    DaySelect.ListIndex = Int(dat.DD) - 1
    stijd = Format(Frac(dat.DD), "hh:mm:ss")
    Hrs = Mid(stijd, 1, 2)
    Min = Mid(stijd, 4, 2)
    Sec = Mid(stijd, 7, 2)
'    Call SetNowButton_Click
    Call calcu(Me.SSTabPlanets.Tab)
End If
End Sub

Private Sub Year_Change()
ADJUST_MONTH_LENGTH
End Sub

Private Sub ADOption_Click()
' Set date to AD mode.
'  ADOption.FontBold = True
'  BCOption.FontBold = False
  ADJUST_MONTH_LENGTH
End Sub

Private Sub BCOption_Click()
' Set date to BC mode.
  'BCOption.FontBold = True
  'ADOption.FontBold = False
  ADJUST_MONTH_LENGTH
End Sub

Private Sub MonthSelect_Click()
ADJUST_MONTH_LENGTH
End Sub

Private Sub Set00HrButton_Click()
' Set interface clock time setting to 00:00
  Hrs.Text = "00": Min.Text = "00": Sec.Text = "00"
End Sub

Private Sub SetNoonButton_Click()
' Set interface clock time setting to 12:00
  Hrs.Text = "12": Min.Text = "00": Sec.Text = "00"
End Sub

Private Sub ADJUST_MONTH_LENGTH()
       
End Sub

Private Sub SetNowButton_Click()
' Set interface date and time to Now according to system clock.

  SET_INTERFACE_DATE_AND_TIME_TO_NOW
  
End Sub

Private Sub SET_INTERFACE_DATE_AND_TIME_TO_NOW()
' Set interface date and time to Now according to system clock.

' DEPENDENCIES: ADJUST_MONTH_LENGTH

Dim Q, QD, QT, MMM, DD, YYYY, HH, MM, ss

' Read current time and date settings from system clock.
   Q = Now
  QT = Format(Q, "hh:mm:ss")
  QD = Format(Q, "dd mm yyyy")

' Set interface time setting to match system clock time.
  HH = Left(QT, 2)
  MM = Mid(QT, InStr(1, QT, ":") + 1, 2)
  ss = Mid(QT, InStr(InStr(1, QT, ":") + 1, QT, ":") + 1, 2)
  Hrs.Text = HH: Min.Text = MM: Sec.Text = ss

' Set interface date setting to match system clock date.
  MMM = Mid(QD, InStr(1, QD, " ") + 1, 3)
  DD = Mid(QD, 1, InStr(1, QD, " "))
  YYYY = Val(Mid(QD, InStr(InStr(1, QD, " ") + 1, QD, " ") + 1, 4))
  MonthSelect.ListIndex = MMM - 1
  ADJUST_MONTH_LENGTH
  DaySelect.ListIndex = DD - 1
  Year.Text = Trim(YYYY)
'  ADOption.Value = True
  
End Sub

Private Function INTERFACE_DATE() As String
' Return the current interface date setting as a date string
' in the standard format such as "20 MAY 1977 BC|AD"

  Dim Q, M, D, Y
  
   D = Right(" " & Trim(DaySelect.Text) & " ", 3)
   M = Trim(MonthSelect.Text) & " "
   Y = Year.Text
'If BCOption.Value = True Then y = y & " BC" Else y = y & " AD"
   Y = Right("      " & Y, 7)
  
   INTERFACE_DATE = D & M & Y
  
End Function

Private Function INTERFACE_TIME() As String
' Return the current interface time setting as a time string
' in the standard format such as "01:23:45"

  INTERFACE_TIME = Hrs.Text & ":" & Min.Text & ":" & Sec.Text
  
End Function

Private Sub startTimer()
    With tmrInterval
        .Enabled = False
        .Interval = 100 'Event every minute
        .Enabled = True
    End With
    updateTimerStatus
End Sub

Private Sub stopTimer()
    tmrInterval.Enabled = False
    statusBar2.Panels(1).Enabled = False
    Call calcu
End Sub

Private Sub tmrInterval_Timer()
    updateTimerStatus
End Sub
    
Private Sub updateTimerStatus()
Dim sdat As String
Dim dat As tDatum
    Dim iNext As Integer
    sdat = Format(Now(), "dd-mm-yyyy hh:mm:ss")
    dat.jj = Val(Mid(sdat, 7, 4))
    dat.MM = Val(Mid(sdat, 4, 2))
    dat.DD = Val(Mid(sdat, 1, 2)) + Val(Mid(sdat, 12, 2)) / 24# + Val(Mid(sdat, 15, 2)) / 1440# + Val(Mid(sdat, 18, 2)) / 86400#
    statusBar2.Panels(2).Text = Format(PlaatselijkeSterrentijd(dat), "hh:mm:ss")
    statusBar2.Panels(4).Text = Format(sdat, "hh:mm:ss")
End Sub

Private Sub Year_GotFocus()
    Year.SelStart = 0
    Year.SelLength = Len(Year.Text)
End Sub

Private Sub BerekenPositieMaan(maan As Long, JD As Double, bShadow As Boolean, _
                             ByRef vMaan As TVECTOR, ByRef rJup As Double)

Dim T As Double, deltaT As Double, DtofUT As Double, obl As Double
Dim NutLon As Double, NutObl As Double, TimeZone As Double
Dim sHelio As TSVECTOR, ShelioJ As TSVECTOR, SNiks As TSVECTOR, SEarth As TSVECTOR, sGeo As TSVECTOR
Dim I As Long
Dim vDummy As TVECTOR, vdummy2 As TVECTOR
Dim nArg As Double

T = JDToT(JD)
deltaT = ApproxDeltaT(T)
DtofUT = T + deltaT * secToT

obl = Obliquity(T)

'{ Main Calculations }
Call PlanetPosHi(5, T, sHelio, True)
Call PlanetPosHi(0, T, SEarth, True)
Call HelioToGeo(sHelio, SEarth, sGeo)

'{ Do just one light time iteration }

For I = 1 To 1
    Call PlanetPosHi(5, DtofUT - sGeo.r * LightTimeConst, sHelio, True)
    Call HelioToGeo(sHelio, SEarth, sGeo)
Next
'{---------------------------------------------------------------------------}
  T = DtofUT - sGeo.r * LightTimeConst
  If Not bShadow Then
      Call JSatEclipticPosition(DUMMY_SATELLITE, T, vDummy)
      Call JSatViewFrom(DUMMY_SATELLITE, vDummy, sGeo, vDummy, False, vDummy, True)
      Call JSatEclipticPosition(maan, T, vMaan)
      Call JSatViewFrom(maan, vMaan, sGeo, vDummy, False, vMaan, True)
      rJup = 1
  Else
      Call JSatEclipticPosition(DUMMY_SATELLITE, T, vdummy2)
      Call JSatViewFrom(DUMMY_SATELLITE, vdummy2, sHelio, vdummy2, False, vdummy2, False)
      Call JSatEclipticPosition(maan, T, vMaan)
      Call JSatViewFrom(maan, vMaan, sHelio, vdummy2, False, vMaan, False)
      rJup = 1
      If vMaan.Z < 0 Then
'        rJup = 1 + (-8.74789916 * vMaan.Z) / (2095.20826331 * sGeo.r) - rMaanJup(maan)
      End If
'rJup = 1 + (-8.74789916 * vMaan.Z) / (2095.20826331 * sGeo.r) - rMaanJup(maan)
  End If
End Sub


Private Function FormatX(expression, fformat, Optional blnMetPunt As Boolean = False)
Dim sFormat As String
Dim nPos As Long
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

