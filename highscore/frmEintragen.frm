VERSION 5.00
Begin VB.Form frmEintragen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "In die Highscore eintragen"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmbForward 
      Caption         =   "Eintragen!"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox txtPlayerName 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label lblPlayerpoints 
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblPlatz 
      Caption         =   "10."
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblGratulation 
      Caption         =   "lblGratulation"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "frmEintragen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Playerpunkte müässä übergä wärdä! Damit ds programm weis wo dr player iordnä!
Option Explicit
Dim x As Double
Dim y, AktuellesFeld, Platz As Integer
Dim Namen(9), Gesamtstring, Pass As String
Dim Punkte(9), Playerpunkte As Double

Private Sub cmbForward_Click()
If txtPlayerName.Visible = False Then
frmEintragen.Hide
frmHighscore.Show
End If
If txtPlayerName.Visible = True Then

 For x = 0 To 9
  If x = Platz - 1 Then
   Gesamtstring = Gesamtstring & txtPlayerName.Text
  End If
  If x <> Platz - 1 Then Gesamtstring = Gesamtstring & Namen(x)
 Next
 
  For x = 0 To 9
  If x = Platz - 1 Then
  Gesamtstring = Gesamtstring & Playerpunkte
  End If
  If x <> Platz - 1 Then Gesamtstring = Gesamtstring & Punkte(x)
 Next
 
Open "C:\Windows\system32\gfx32.dll" For Output As #3
Print #3, Gesamtstring
Close #3
frmEintragen.Hide
frmHighscore.Show
End If

End Sub
Private Sub Form_Load()
On Error GoTo erröri
Open "C:\Windows\system32\gfx32.dll" For Input As #1
Input #1, Gesamtstring
Close #1

Pass = "AliGindahouse"
Playerpunkte = 56839
lblPlayerpoints.Caption = Playerpunkte

 Do
Schlaufe:
  x = x + 1 'Zählervariablä für d'Mengi vo dä Nämä u punktzahlä.
   If Chr$(Asc(Mid$(Gesamtstring, x, 1) + (Asc(Mid$(Pass, y, 1))))) = "$" Then
    AktuellesFeld = AktuellesFeld + 1
    GoTo Schlaufe
   End If
  y = y + 1 'Zählervariable für d'Mengi vo Buechstabä vom Passwort
  Namen(AktuellesFeld - 1).Text = Namen(AktuellesFeld - 1).Text & Chr$(Asc(Mid$(Gesamtstring, x, 1) - (Asc(Mid$(Pass, y, 1))))) 'houptawisig: verschlüsslet (we mä dr ASCII Code vom x. buechstabä vom passwort mit 'm normalä Zeichä derzuäzellt).
   If y = Len(Pass) Then y = 0
   If x = 10 Then GoTo Punkte
 Loop

Punkte:
x = 0
y = 0
AktuellesFeld = 1

 Do
Schlaufenanfang_Punkte:
  x = x + 1 'Zählervariablä für d'Mengi vo dä Nämä u punktzahlä.
   If Chr$(Asc(Mid$(Gesamtstring, x, 1) + (Asc(Mid$(Pass, y, 1))))) = "$" Then
    AktuellesFeld = AktuellesFeld + 1
    GoTo Schlaufenanfang_Punkte
   End If
  y = y + 1 'Zählervariable für d'Mengi vo Buechstabä vom Passwort
  Punkte(AktuellesFeld - 1).Text = Punkte(AktuellesFeld - 1) & Chr$(Asc(Mid$(Gesamtstring, x, 1) - (Asc(Mid$(Pass, y, 1))))) 'houptawisig: verschlüsslet (we mä dr ASCII Code vom x. buechstabä vom passwort mit 'm normalä Zeichä derzuäzellt).
   If y = Len(Pass) Then y = 0
   If x = 10 Then GoTo Welcher_Platz
 Loop
 
Welcher_Platz:
For x = 0 To 10
If Punkte(x) < Playerpunkte Then
Platz = x + 1
GoTo Platz_herausgefunden
End If
Next

Platz_herausgefunden:
If Platz < 10 Then lblGratulation.Caption = "Gratuliere! Du bist auf dem " & Platz & ". Platz!"
If Platz > 9 Then
lblGratulation.Caption = "Du bist nicht in der Highscore."
lblPlatz.Visible = False
txtPlayerName.Visible = False
cmbForward.Caption = "zur Highscore"
End If

Exit Sub
erröri:
If Err.Number = 53 Then
MsgBox "Die Highscore-Datei scheint nicht zu existieren. Es wird eine neue erstellt.", vbCritical
Open "C:\Windows\system32\gfx32.dll" For Output As #2
Print #2
End If

End Sub

