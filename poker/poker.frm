VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Poker Stack Control"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Money to spend"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4215
      Begin VB.Label lblMaxBet 
         Caption         =   "Maximal stack before leaving:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label lblBankroll 
         Caption         =   "5% of my bankroll:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Available Money"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtMoney 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Available Money:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtMoney_Change()
If txtMoney.Text <> "" Then
lblBankroll.Caption = "5% of my bankroll: " & Round(CInt(txtMoney.Text) / 20, 2)
lblMaxBet.Caption = "Maximal stack before leaving: " & Round(CInt(txtMoney.Text) / 9, 2)
End If
End Sub
Private Sub txtMoney_KeyPress(KeyAscii As Integer)
   If InStr("1234567890." & Chr$(8), Chr$(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End Sub
