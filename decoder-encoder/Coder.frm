VERSION 5.00
Begin VB.Form frmCoder 
   Caption         =   "Decoder (C) 05 by FoN_qwx"
   ClientHeight    =   1380
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4590
   ScaleHeight     =   1380
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMovement 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton cmbOK 
      Caption         =   "DECODE/ENCODE"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtCode 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox txtSource 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblCode 
      Caption         =   "Decoded:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.Label txtText 
      Caption         =   "Text:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblMovement 
      Caption         =   "Movement:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmCoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbOK_Click()
Source = txtSource.Text
movement = Val(txtMovement.Text)
For x = 1 To Len(Source)
Code = Code & Chr(Asc(Mid(Source, x, 1)) + movement)
Next
txtCode.Text = Code
End Sub
