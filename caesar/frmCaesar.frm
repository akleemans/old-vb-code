VERSION 5.00
Begin VB.Form frmCaesar 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Entschlüsseln im Cäsar-System"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtVerschiebung 
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Top             =   480
      Width           =   495
   End
   Begin VB.CheckBox chkAuto 
      Caption         =   "Automatisches Speichern"
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txtCodewort 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.OptionButton optSchwierig 
      Caption         =   "Cäsar mit Codewort"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.OptionButton optEinfach 
      Caption         =   "Cäsar einfach"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Verschiebung:"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmCaesar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAuto_Click()
If txtVerschiebung.Enabled = True Then txtVerschiebung.Enabled = False
If txtVerschiebung.Enabled = False Then txtVerschiebung.Enabled = True

End Sub

Private Sub cmdOK_Click()
Dim text, text2(256), text3 As String
Dim x, y, Verschiebung As Double
Verschiebung = Val(txtVerschiebung.text)
Pfad = App.Path & "\text.txt"
Open Pfad For Input As #1
Input #1, text
Close #1
If optEinfach.Value = True And chkAuto.Value = 1 Then

Open App.Path & "\Erg" & y & ".txt" For Output As #2

 For y = 1 To 26
  For x = 1 To Len(text)
     If Asc(Mid$(text, x, 1)) + y > 91 And Asc(Mid$(text, x, 1)) + y < 95 Or Asc(Mid$(text, x, 1)) + y > 121 Then
     text2(x) = Chr$(Asc(Mid$(text, x, 1)) + y - 26)
     GoTo weiter
   End If
   text2(x) = Chr$(Asc(Mid$(text, x, 1)) + y)
weiter:
  Next
   text3 = ""
   For z = 1 To Len(text)
    text3 = text3 & text2(z)
   Next
  Print #2, text3
 Next


Close #2
End If
End Sub

