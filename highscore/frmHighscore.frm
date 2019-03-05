VERSION 5.00
Begin VB.Form frmHighscore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Highscore"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Lissen"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmbBack 
      Caption         =   "Zurück"
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label lblPlayerPoints 
      Height          =   255
      Index           =   9
      Left            =   4320
      TabIndex        =   33
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label lblPlayerPoints 
      Height          =   255
      Index           =   8
      Left            =   4320
      TabIndex        =   32
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblPlayerPoints 
      Height          =   255
      Index           =   7
      Left            =   4320
      TabIndex        =   31
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lblPlayerPoints 
      Height          =   255
      Index           =   6
      Left            =   4320
      TabIndex        =   30
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lblPlayerPoints 
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   29
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblPlayerPoints 
      Height          =   255
      Index           =   4
      Left            =   4320
      TabIndex        =   28
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblPlayerPoints 
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   27
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblPlayerPoints 
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   26
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblPlayerPoints 
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   25
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblPoints 
      Caption         =   "Punkte:"
      Height          =   255
      Left            =   4320
      TabIndex        =   24
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblPlayerPoints 
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   23
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblName 
      Caption         =   "Name:"
      Height          =   255
      Left            =   960
      TabIndex        =   22
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblPlayerName 
      Height          =   255
      Index           =   9
      Left            =   960
      TabIndex        =   21
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label lblPlayerName 
      Height          =   255
      Index           =   8
      Left            =   960
      TabIndex        =   20
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label lblPlayerName 
      Height          =   255
      Index           =   7
      Left            =   960
      TabIndex        =   19
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Label lblPlayerName 
      Height          =   255
      Index           =   6
      Left            =   960
      TabIndex        =   18
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label lblPlayerName 
      Height          =   255
      Index           =   5
      Left            =   960
      TabIndex        =   17
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label lblPlayerName 
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   16
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label lblPlayerName 
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   15
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label lblPlayerName 
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   14
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label lblPlayerName 
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   13
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label lblPlayerName 
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   12
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label lblIndex 
      Caption         =   "Platz:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblPlayerIndex 
      Caption         =   "10"
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   9
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label lblPlayerIndex 
      Caption         =   "9"
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   8
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label lblPlayerIndex 
      Caption         =   "8"
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   7
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label lblPlayerIndex 
      Caption         =   "7"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   6
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblPlayerIndex 
      Caption         =   "6"
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   5
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblPlayerIndex 
      Caption         =   "5"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lblPlayerIndex 
      Caption         =   "4"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblPlayerIndex 
      Caption         =   "3"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label lblPlayerIndex 
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   255
   End
   Begin VB.Label lblPlayerIndex 
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   255
   End
End
Attribute VB_Name = "frmHighscore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variablädeklaration notwendig!
Option Explicit
Dim Gesamtstring, Pass As String
Dim x As Double
Dim AktuellesFeld, y As Integer
Private Sub cmbBack_Click()
'Houptframe ladä. frmHauptframe mit m original namä ersetzä.
frmHighscore.Hide
'frmHauptframe.Show
End Sub
Private Sub Form_Load()
On Error GoTo erröri
'Passwort bestimmä. Chasch natürlech o es anders nä. Achtung: ned über 127 Buechstabä!;-) Süsch muesch obä Pass als double oder float deklarierä. Im moment sötti aber Integer (256 Zahläwärtä) längä.
Pass = "AliGindahouse"

'Datei iläsä wo d'datä drin si. D Struktur vom verschlüssletä züg isch so: Namä1|Namä2|Namä3... bis 10 u näär Namä10|Punktzahl1|Punktzahl2|Punkzahl3 etc.
Open "C:\Windows\system32\gfx32.dll" For Input As #1
Print #1, Gesamtstring
Close #1

'Variabläbeschtimmig für d'schloufä unä
x = 0
AktuellesFeld = 1

'---------------------------------Schloufä 1-------------------------------------

'Tataaaa!! hiä wirds passwort entschlüsslet(isch rächt kompliziert...;-) u grad id fälder kopiert
 Do
Schlaufenanfang:
  x = x + 1 'Zählervariablä für d'Mengi vo dä Nämä u punktzahlä.
   If Chr$(Asc(Mid$(Gesamtstring, x, 1) + (Asc(Mid$(Pass, y, 1))))) = "$" Then
    AktuellesFeld = AktuellesFeld + 1
    GoTo Schlaufenanfang
   End If
  y = y + 1 'Zählervariable für d'Mengi vo Buechstabä vom Passwort
  lblPlayerName(AktuellesFeld - 1) = lblPlayerName(AktuellesFeld - 1) & Chr$(Asc(Mid$(Gesamtstring, x, 1) - (Asc(Mid$(Pass, y, 1))))) 'Das hiä isch d'houptawisig: Si Füegt am fäld lblPlayerName(x) das i, wo entsteit we mä dr ASCII Code vom x. buechstabä vom passwort vom verschlüssletä Zeichä abzieht.
   If y = Len(Pass) Then y = 0
   If x = 10 Then GoTo Punkte_einfügen
 Loop
 
Punkte_einfügen:
x = 0
y = 0
AktuellesFeld = 1

'---------------------------------Schloufä 2-------------------------------------

'D'schloufä für d'Punkt. Anderi fälder, süsch glich.
 Do
  y = y + 1 'Zählervariable für d'Mengi vo Buechstabä vom Passwort
Schlaufenanfang_der_Punkte:
  x = x + 1 'Zählervariablä für d'Mengi vo dä Nämä u punktzahlä.
   If Chr$(Asc(Mid$(Gesamtstring, x, 1) + (Asc(Mid$(Pass, y, 1))))) = "$" Then
    AktuellesFeld = AktuellesFeld + 1
    GoTo Schlaufenanfang_der_Punkte
   End If
  lblPlayerpoints(AktuellesFeld - 1) = lblPlayerpoints(AktuellesFeld - 1) & Chr$(Asc(Mid$(Gesamtstring, x, 1) - (Asc(Mid$(Pass, y, 1)))))
   If y = Len(Pass) Then y = 0
   If x = 10 Then Exit Sub
 Loop

erröri:
If Err.Number = 53 Then
MsgBox "Die Highscore-Datei scheint nicht zu existieren. Es wird eine neue erstellt.", vbCritical
Open "C:\Windows\system32\gfx32.dll" For Output As #2
Print #2
End If
End Sub

