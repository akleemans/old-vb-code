VERSION 5.00
Begin VB.Form frmOgamer 
   Caption         =   "Ogame Zeiten berechnen"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   1590
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   1590
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Benötigt"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label13 
      Caption         =   "Zeitdauer"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Vorhanden"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   840
   End
   Begin VB.Label Label4 
      Caption         =   "Zuwachs"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   735
   End
End
Attribute VB_Name = "frmOgamer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
If Text1.Text <> 0 And Text2.Text <> 0 And Text3.Text <> 0 Then
 verzögerung = (Val(Text3.Text) - Val(Text2.Text)) / Val(Text1.Text) * 3600
 zeit = DateAdd("s", verzögerung, Time)
 Text11.Text = zeit
End If
End Sub
