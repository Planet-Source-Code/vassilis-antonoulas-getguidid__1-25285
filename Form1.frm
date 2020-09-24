VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   435
      TabIndex        =   1
      Top             =   360
      Width           =   4080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create GUID"
      Height          =   390
      Left            =   3300
      TabIndex        =   0
      Top             =   1095
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Text1.Text = GetGuidID

End Sub
