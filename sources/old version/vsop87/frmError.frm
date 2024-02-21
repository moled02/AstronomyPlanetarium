VERSION 5.00
Begin VB.Form frmError 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Oscar Monitor: Error Detected"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShow 
      Caption         =   "&Show"
      Default         =   -1  'True
      Height          =   375
      Left            =   2775
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1335
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.PictureBox picError 
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   120
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   0
      Top             =   240
      Width           =   540
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      Caption         =   "Msg"
      Height          =   975
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub cmdShow_Click()
    With frmSterrentijd
        If (.WindowState = vbMinimized) Then .WindowState = vbNormal
        .Visible = True
        SetForegroundWindow .hWnd
    End With
    Unload Me
End Sub
