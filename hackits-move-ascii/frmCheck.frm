VERSION 5.00
Begin VB.Form frmCheck 
   Caption         =   "Check"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   12150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmbZurueck 
      Caption         =   "zurück"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmbWeiter 
      Caption         =   "Nächste Kombination..."
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   11535
   End
End
Attribute VB_Name = "frmCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim text, textneu As String
Dim x, i As Integer
Private Sub Form_Load()
text = "Qzltmjaqt kqcg je jszseh ltnxiitm, asc ism pqtz isbptm iecc. Sktz dtgug ctptm aqz iso, xk je jqbp iqg Hszktm secntmmcg. Atobpt Hszkt psg jtz Pqmgtzlzemj jqtctz Ctqgt?"
text = LCase(text)
Label1.Caption = text
End Sub
Private Sub cmbWeiter_Click()
Call calculate(1)
End Sub
Private Sub cmbZurueck_Click()
Call calculate(-1)
End Sub
Private Function calculate(ByVal x As Integer)
For i = 1 To Len(text)
    If Mid(text, i, 1) = "a" And x = -1 Then
    textneu = textneu & "z"
    ElseIf Mid(text, i, 1) = "z" And x = 1 Then
    textneu = textneu & "a"
    ElseIf Mid(text, i, 1) = " " Or Mid(text, i, 1) = "?" Or Mid(text, i, 1) = "!" Or Mid(text, i, 1) = "." Then
    textneu = textneu & Mid(text, i, 1)
    Else
    textneu = textneu & Chr(Asc(Mid(text, i, 1)) + x)
    End If
Next
text = textneu
Label1.Caption = text
textneu = ""
End Function
