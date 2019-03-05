VERSION 5.00
Begin VB.Form frmCheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmbDelete 
      Caption         =   "Markierte löschen"
      Height          =   555
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   1200
      TabIndex        =   5
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton cmbHinzu 
      Caption         =   "Regel hinzufügen..."
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblElemente 
      Caption         =   "Anzahl Elemente: 0"
      Height          =   255
      Left            =   7560
      TabIndex        =   8
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label lblVorschau 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   11655
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   12000
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label3 
      Caption         =   "zu"
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   1800
      Width           =   375
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12240
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label lblText 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   11535
   End
End
Attribute VB_Name = "frmCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim text, textneu, von(1 To 30), zu(1 To 30) As String
Dim x, i, i1, zaehlen As Integer
Private Sub cmbDelete_Click()
von(List1.ListIndex + 1) = ""
For i = 1 To zaehlen - 1
    If i > 1 Then
        If von(i - 1) = "" Then
        von(i - 1) = von(i)
        zu(i - 1) = zu(i)
        von(i) = ""
        zu(i) = ""
        End If
    End If
Next
zaehlen = zaehlen - 1

List1.Clear
For i = 1 To zaehlen
Call adder(von(i), zu(i))
Next
End Sub
Private Sub Form_Load()
text = "Qzltmjaqt kqcg je jszseh ltnxiitm, asc ism pqtz isbptm iecc. Sktz dtgug ctptm aqz iso, xk je jqbp iqg Hszktm secntmmcg. Atobpt Hszkt psg jtz Pqmgtzlzemj jqtctz Ctqgt?"
text = LCase(text)
lblText.Caption = text
lblVorschau.Caption = text
zaehlen = 0
End Sub
Private Sub cmbHinzu_Click()
zaehlen = zaehlen + 1
von(zaehlen) = Text1.text
zu(zaehlen) = Text2.text
Call adder(Text1.text, Text2.text)
Text1.text = ""
Text2.text = ""
Call calc
lblElemente.Caption = "Anzahl Elemente: " & zaehlen

End Sub
Private Sub calc() 'Wendet Regeln am Originaltext an

textneu = text

For i = 1 To Len(textneu)  'Alle Zeichen durchgehen
    For i1 = 1 To zaehlen  'Alle Regeln an aktuellem Zeichen durchgehen
    
        If Mid(textneu, i, 1) = von(i1) Then
        Mid(textneu, i, 1) = zu(i1)
        GoTo hierweiter
        End If
        
    Next
hierweiter:
Next

lblVorschau.Caption = textneu
textneu = ""
End Sub
Private Function adder(ByVal x As String, ByVal y As String)
'Zu logfile-label hinzufügen
List1.AddItem x & " >> " & y
End Function
