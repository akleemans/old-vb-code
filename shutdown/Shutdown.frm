VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2205
   LinkTopic       =   "Form1"
   ScaleHeight     =   780
   ScaleWidth      =   2205
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   120
   End
   Begin VB.CommandButton cmbOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Double
Private Sub CmbOK_Click()
a = 1
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
a = a + 1
If a = 28000 Then
Shell "shutdown -s"
End If
cmbOK.Caption = a
End Sub
