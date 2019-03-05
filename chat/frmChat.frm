VERSION 5.00
Begin VB.Form frmChat 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Chatprogramm"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstNachrichten 
      Height          =   3765
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7335
   End
   Begin VB.CommandButton cmbSend 
      Caption         =   "Senden"
      Height          =   255
      Left            =   6720
      TabIndex        =   1
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txtNeueNachricht 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   6375
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Nachrichten(128) As String

Private Sub cmbSend_Click()
 If txtNeueNachricht = "" Then
  Nachricht = TextFeld_leer()
  lstNachrichten.AddItem = Nachricht
  Exit Sub
 End If
Nachricht = txtNeueNachricht.Text
Nachrichten = Nachricht_Zerschneiden(Nachricht)
lstNachrichten.AddItem = Antwort_vom_PC(Nachrichten)
End Sub
Public Sub Nachricht_Zerschneiden(Nachricht As String)
Dim Position As Double
Dim InhaltPosition As Double
Dim NachrichtenPos As Integer
Dim Wortlänge As Integer

 While Position <= Len(Nachricht)
  Position = Position + 1
  Wortlänge = Wortlänge + 1
  InhaltPosition = Mid(Nachricht, Position, 1)
   If InhaltPosition = "" Then
   Nachrichten(NachrichtenPos) = Mid(Nachricht, Position - Wortlänge, Wortlänge)
   NachrichtenPos = NachrichtenPos + 1
   Wortlänge = 0
   End If
 Wend
End Sub
Public Sub Antwort_vom_PC(Nachrichten As String)

End Sub
Public Sub Neuer_User()

End Sub
Public Sub TextFeld_leer()
Randomize Timer

End Sub
