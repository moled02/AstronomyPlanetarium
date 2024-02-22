VERSION 5.00
Begin VB.Form frmSettings 
   Caption         =   "Settings"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtAltitude 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Text            =   "0"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtScheidingsteken 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "|"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1493
      TabIndex        =   5
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox txtLongitude 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Text            =   "-05.08"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtLatitude 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Text            =   "52.05"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Altitude"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   3360
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label3 
      Caption         =   "Separation sign"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Latitude north"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Western longitude"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
ret = SetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", REG_SZ, "Latitude", _
      Me.txtLatitude)
ret = SetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", REG_SZ, "Longitude", _
      Me.txtLongitude)
ret = SetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", REG_SZ, "Scheidingsteken", _
      Me.txtScheidingsteken)
ret = SetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", REG_SZ, "Altitude", _
      Me.txtAltitude)
Unload Me
End Sub

Private Sub Form_Load()
Dim sLongitude As String
Dim sLatitude As String
Dim sScheidingsteken As String
Dim sAltitude As String
Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Latitude", _
            sLatitude)
Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Longitude", _
            sLongitude)
Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Scheidingsteken", _
            sScheidingsteken)
Call GetKeyValue(HKEY_CURRENT_USER, "Software\Belastingdienst\Astronomie", "Altitude", _
            sAltitude)
Me.txtLatitude = sLatitude
Me.txtLongitude = sLongitude
Me.txtScheidingsteken = sScheidingsteken
Me.txtAltitude = sAltitude
#If FRANS Then
    Me.Caption = "Mises de programme"
    Me.Label1.Caption = "Longitude Ouest"
    Me.Label2.Caption = "Latitude Nord"
    Me.Label3.Caption = "Signe"
    Me.Label4.Caption = "Altitude"
#End If
End Sub



