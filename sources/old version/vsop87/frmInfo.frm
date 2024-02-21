VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Info"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   6255
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmInfo.frx":030A
      Top             =   720
      Width           =   6735
   End
   Begin VB.Label Label1 
      Caption         =   "Versie : "
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim objspecialfolder As New clsSpecialFolder
#If FRANS Then
    Me.Caption = "Information"
    Me.Label1.Caption = "Version : "
#End If
Me.Label1.Caption = Me.Label1.Caption & App.Major & "." & App.Minor & "." & App.Revision & " (" & _
ShowFileInfo(App.Path & "\" & App.EXEName & ".exe") & ")"
Me.Text1.Text = Me.Text1.Text & objspecialfolder.TemporaryFolder

End Sub

Function ShowFileInfo(filespec As String) As String
On Error GoTo fout:
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    s = f.DateCreated
    ShowFileInfo = Format(s, "dd-mm-yyyy")
    Exit Function
fout:
    ShowFileInfo = ""
End Function

