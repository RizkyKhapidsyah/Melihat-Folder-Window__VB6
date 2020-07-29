VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Melihat Folder Window"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Lihat"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub OpenDirectory(Directory As String)
ShellExecute 0, "Open", Directory, vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

Private Sub Command1_Click()
   'Ganti "C:\" di bawah dengan folder yang ingin Anda
   'lihat
   OpenDirectory ("C:\")
End Sub


