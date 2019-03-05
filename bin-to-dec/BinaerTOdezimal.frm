VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   915
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   1950
   LinkTopic       =   "Form1"
   ScaleHeight     =   915
   ScaleWidth      =   1950
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDezimal 
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton cmbUmsetzen 
      Caption         =   "Umsetzen!"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtBinaer 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblBinaer 
      Caption         =   "Binörcode:"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbUmsetzen_Click()
For x = 1 To Len(txtBinaer.Text)
Zahl = Zahl + (Mid$(txtBinaer.Text, Len(txtBinaer.Text) - x + 1, 1) * (2 ^ x))
Next
txtDezimal.Text = Zahl
End Sub
