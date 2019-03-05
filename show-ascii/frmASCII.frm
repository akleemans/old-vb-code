VERSION 5.00
Begin VB.Form frmASCII 
   Caption         =   "ASCII Zeichen ausgeben"
   ClientHeight    =   750
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   ScaleHeight     =   750
   ScaleWidth      =   3840
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtASCII 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
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
      Width           =   3255
   End
End
Attribute VB_Name = "frmASCII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
vary = 0
Do
txtASCII.Text = txtASCII.Text & vary & " = " & Chr$(vary) & " | "
vary = vary + 1
If vary > 255 Then Exit Sub
Loop
End Sub

