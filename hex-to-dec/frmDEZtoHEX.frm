VERSION 5.00
Begin VB.Form frmDEZtoHEX 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HEX-Zahlen zu Dezimalzahlen konvertieren"
   ClientHeight    =   1710
   ClientLeft      =   150
   ClientTop       =   765
   ClientWidth     =   3735
   BeginProperty Font 
      Name            =   "DicotMedium"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDEZtoHEX.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   1710
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmbDELDEZ 
      Caption         =   "D"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton cmbDELHEX 
      Caption         =   "D"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton cmbMORPH 
      Caption         =   "Umwandeln"
      Height          =   435
      Left            =   960
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtDEZ 
      Height          =   375
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtHEX 
      Height          =   375
      Left            =   1920
      MaxLength       =   8
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblDEZ 
      Caption         =   "Dezimalzahl:"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblHEX 
      Caption         =   "HEX-Zahl:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Menu mnu_datei 
      Caption         =   "Datei"
      Begin VB.Menu mnu_help 
         Caption         =   "Hilfe..."
         Index           =   1
         Shortcut        =   ^H
      End
      Begin VB.Menu mnu_about 
         Caption         =   "About..."
         Index           =   2
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmDEZtoHEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DEZZahl, Hexzähler, Zähler, x, HEXER(100) As Double
Dim HEXZahl As String

Private Sub cmbDELDEZ_Click()
txtDEZ.Text = ""
End Sub

Private Sub cmbDELHEX_Click()
txtHEX.Text = ""
End Sub

Private Sub cmbMORPH_Click()
If txtHEX.Text = "" Then GoTo From_Dez_to_Hex
If txtDEZ.Text = "" Then GoTo From_Hex_to_Dez
If txtHEX.Text = "" And txtDEZ.Text = "" Or txtHEX.Text <> "" And txtDEZ.Text <> "" Then
MsgBox "Da stimmt was nicht..."
Exit Sub
End If

From_Hex_to_Dez:
Dim Ziffer As Double
Dim buchstabe As String
Ziffer = 0
DEZZahl = 0
x = 0

schlaufe:
Do
x = x + 1
If x > Len(txtHEX.Text) Then GoTo ende
buchstabe = Mid$(txtHEX.Text, Len(txtHEX.Text) - x + 1, 1)
If buchstabe = "A" Then DEZZahl = DEZZahl + (10 * (16 ^ (x - 1)))
If buchstabe = "B" Then DEZZahl = DEZZahl + (11 * (16 ^ (x - 1)))
If buchstabe = "C" Then DEZZahl = DEZZahl + (12 * (16 ^ (x - 1)))
If buchstabe = "D" Then DEZZahl = DEZZahl + (13 * (16 ^ (x - 1)))
If buchstabe = "E" Then DEZZahl = DEZZahl + (14 * (16 ^ (x - 1)))
If buchstabe = "F" Then DEZZahl = DEZZahl + (15 * (16 ^ (x - 1)))
Ziffer = Val(Mid$(txtHEX.Text, x, 1))
DEZZahl = DEZZahl + (Ziffer * (16 ^ (x - 1)))
If x = 100 Then GoTo ende
Loop

ende:
txtDEZ.Text = DEZZahl
Exit Sub

From_Dez_to_Hex:
Zähler = 0
Erase HEXER
HEXER(1) = Val(txtDEZ.Text)

'berechnen mit Geschachtelter Schlaufe
x = 1

' evt Optimieren mit - 16*16 abzählen (direkt 256 abzählen)
 Do
  Do
   If HEXER(x) - 16 <= 0 Then GoTo byebye
   HEXER(x) = HEXER(x) - 16
   Zähler = Zähler + 1
  Loop

byebye:
  HEXER(x + 1) = Zähler
  x = x + 1
  If x = 100 Then GoTo Auswertung_Zu_Zahlen
  Zähler = 0
 Loop

Auswertung_Zu_Zahlen:
x = 1
Do
If x = 100 Then Exit Sub
If HEXER(x) = 0 Then GoTo Nuller
If HEXER(x) > 9 Then
If HEXER(x) = 10 Then txtHEX.Text = "A" & txtHEX.Text
If HEXER(x) = 11 Then txtHEX.Text = "B" & txtHEX.Text
If HEXER(x) = 12 Then txtHEX.Text = "C" & txtHEX.Text
If HEXER(x) = 13 Then txtHEX.Text = "D" & txtHEX.Text
If HEXER(x) = 14 Then txtHEX.Text = "E" & txtHEX.Text
If HEXER(x) = 15 Then txtHEX.Text = "F" & txtHEX.Text
GoTo Nuller
End If
txtHEX.Text = HEXER(x) & txtHEX.Text
Nuller:
x = x + 1
Loop

End Sub

Private Sub mnu_about_Click(Index As Integer)
MsgBox "(c) 2004 by qwx(Ad Kleemans). For Questions email me: a.kleemans@gmx.ch or visit my website: www.qwx.net.ms. Have Fun!"
End Sub

Private Sub mnu_help_Click(Index As Integer)
MsgBox "Create a HEX-Code from a Decimal Code or a Decimal Code (Normal Number) from a HEX Code."
End Sub
