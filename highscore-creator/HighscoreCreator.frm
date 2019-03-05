VERSION 5.00
Begin VB.Form frmHighscoreCreator 
   Caption         =   "Highscore Creator"
   ClientHeight    =   1230
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   1230
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmbOK 
      Caption         =   "Umsetzen"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txtText 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmHighscoreCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x, y, zahl, zahl2 As Double
Dim Gesamtstring, GesamtstringNeu, Pass, zeugs As String

Private Sub cmbOK_Click()
Pass = "AliGindahouse"

Gesamtstring = txtText.Text
GesamtstringNeu = ""
x = 0
y = 0

 Do
Schlaufenanfang:
  x = x + 1
  y = y + 1
  GesamtstringNeu = GesamtstringNeu & Chr$(Asc(Mid$(Gesamtstring, x, 1)) - (Asc(Mid$(Pass, y, 1))))
   If y = Len(Pass) Then y = 0
   If x = Len(Gesamtstring) Then GoTo Umgesetzt
 Loop

Umgesetzt:
txtText.Text = GesamtstringNeu
End Sub

