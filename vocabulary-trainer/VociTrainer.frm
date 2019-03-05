VERSION 5.00
Begin VB.Form frmVociTrainer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Voci Trainer 0.1 beta"
   ClientHeight    =   4590
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleMode       =   0  'User
   ScaleWidth      =   8650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Caption         =   "Bitte Modus eingeben:"
      Height          =   255
      Left            =   360
      TabIndex        =   29
      Top             =   1680
      Width           =   2655
      Begin VB.OptionButton optRandom 
         Caption         =   "Zufall"
         Height          =   255
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optDescending 
         Caption         =   "Absteigend"
         Height          =   255
         Left            =   1080
         TabIndex        =   30
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Bitte Modus eingeben:"
      Height          =   255
      Left            =   1200
      TabIndex        =   17
      Top             =   480
      Width           =   1695
      Begin VB.OptionButton optAbfrage 
         Caption         =   "Test"
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optUeben 
         Caption         =   "Üben"
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Abfrage"
      Height          =   4335
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton cmbExpand 
         Caption         =   ">"
         Height          =   375
         Left            =   3720
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   22
         ToolTipText     =   "Hier klicken, um Listen zu laden oder Noten anzuzeigen"
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   300
         Left            =   3120
         TabIndex        =   16
         ToolTipText     =   "Stoppt die Abfrage"
         Top             =   3050
         Width           =   855
      End
      Begin VB.TextBox txtAnswer 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox txtQuery 
         Enabled         =   0   'False
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2880
         Width           =   2775
      End
      Begin VB.OptionButton optReversed 
         Caption         =   "Umgekehrt"
         Height          =   195
         Left            =   1320
         TabIndex        =   11
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton optNormal 
         Caption         =   "Normal"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CheckBox chkDiff 
         Caption         =   "Auf Schwierigkeit achten"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   2175
      End
      Begin VB.ComboBox cboNumber 
         Height          =   315
         ItemData        =   "VociTrainer.frx":0000
         Left            =   1440
         List            =   "VociTrainer.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton cmbExercise 
         Caption         =   "OK"
         Height          =   300
         Left            =   3000
         TabIndex        =   3
         ToolTipText     =   "Beginnt die Abfrage"
         Top             =   2160
         Width           =   885
      End
      Begin VB.Label lblRichtung 
         Caption         =   "Richtung:"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblModus 
         Caption         =   "Modus:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         ToolTipText     =   "Legt fest, ob Vokabular geübt oder abgefragt werden soll"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblText2 
         Caption         =   "Anzahl Wörter:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label lblAnzeige 
         Caption         =   "Keine Aktion."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   3600
         Width           =   3615
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         X1              =   15
         X2              =   4024
         Y1              =   2655
         Y2              =   2655
      End
      Begin VB.Label lblMode 
         Caption         =   "Reihenfolge:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         ToolTipText     =   "Legt fest, ob die Fremdsprache oder die Herkunftssprache abgefragt werden soll."
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         X1              =   0
         X2              =   4080
         Y1              =   2640
         Y2              =   2640
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Notenverwaltung"
      Height          =   1815
      Left            =   4440
      TabIndex        =   2
      Top             =   2640
      Width           =   4095
      Begin VB.Label lblBewertung 
         Caption         =   "Bewertung: "
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label lblFalse 
         Caption         =   "Falsche: "
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblRichtige 
         Caption         =   "Richtige: "
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Listenverwaltung"
      Height          =   2295
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton cmbLoad 
         Height          =   300
         Left            =   3600
         Picture         =   "VociTrainer.frx":003A
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1440
         Width           =   300
      End
      Begin VB.TextBox txtList 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   23
         Text            =   "Keine Liste ausgewählt"
         Top             =   1440
         Width           =   3255
      End
      Begin VB.CommandButton cmbNewList 
         Caption         =   "Neue Liste anlegen"
         Enabled         =   0   'False
         Height          =   300
         Left            =   240
         TabIndex        =   21
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton cmbLoadList 
         Caption         =   "Liste laden"
         Height          =   300
         Left            =   2280
         TabIndex        =   7
         Top             =   1800
         Width           =   1125
      End
      Begin VB.Label lblListLoaded 
         Caption         =   "Keine Liste geladen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label lblLists 
         Caption         =   "Aktuell geladene Liste:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Menu mnuDatei 
      Caption         =   "Datei"
      Begin VB.Menu mnuSettings 
         Caption         =   "Einstellungen"
      End
      Begin VB.Menu mnuBeenden 
         Caption         =   "Beenden"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "?"
      Begin VB.Menu mnuVersion 
         Caption         =   "Version"
      End
      Begin VB.Menu mnuHilfe 
         Caption         =   "Hilfe"
      End
   End
End
Attribute VB_Name = "frmVociTrainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim words(1 To 1000), answer(1 To 1000), temp(1 To 1000) As String
Dim Listloaded, modeAsk As Boolean
Dim anzWords, gefragte, richtige, wort, i, wievielWoerterAbfragen, bewertung(1 To 1000), numb As Integer
Dim chance As Integer
Dim mittel As Double
Public Function Rand(ByVal low As Long, ByVal high As Long) As Long
  Rand = Int((high - low + 1) * Rnd) + low
End Function
Private Sub cmbExercise_Click()
'Validierung
If Listloaded = False Then
MsgBox "Bitte zuerst eine Voci-Datei laden!", vbExclamation
Exit Sub
End If

'Richtung
If optReversed.Value = True Then
  For i = 1 To anzWords
  temp(i) = words(i)
  words(i) = answer(i)
  answer(i) = temp(i)
  Next
End If

For i = 1 To anzWords
mittel = mittel + bewertung(i)
Next
mittel = Runden(mittel / anzWords, 1) - 1

cmbExercise.Enabled = False
cmdStop.Enabled = True
gefragte = 0
richtige = 0
lblAnzeige.Caption = ""
lblAnzeige.FontItalic = False
modeAsk = True
  If cboNumber.Text = "ganze Liste" Then
  wievielWoerterAbfragen = anzWords
  Else
  wievielWoerterAbfragen = CInt(cboNumber.Text)
  End If
cboNumber.Enabled = False
Call txtAnswer_KeyDown(13, 0)

End Sub
Private Sub cmbExpand_Click()
If frmVociTrainer.Width = 8740 Then
frmVociTrainer.Width = 4420
cmbExpand.Caption = ">"
Else
frmVociTrainer.Width = 8740
cmbExpand.Caption = "<"
End If

End Sub
Private Sub cmbLoadList_Click()
Dim teile As Variant
Dim n As Integer
Dim zeile As String
Dim Daten As String 'Hier wird der Inhalt der Textdatei gespeichert.

If txtList.Text = "" Or txtList.Text = "Keine Liste ausgewählt" Then
MsgBox "Bitte Liste angeben.", vbExclamation
Exit Sub
End If

If Dir$(txtList.Text) = "" Then
MsgBox "Bitte geben Sie einen gültige Dateinamen ein.", vbExclamation
Exit Sub
End If

If LCase(Right(txtList.Text, 4)) <> ".txt" Then
MsgBox "Sie versuchen, eine Datei zu laden, die möglicherweise einen Absturz des Programms nach sich ziehen kann.", vbExclamation
End If

numb = FreeFile
Open txtList.Text For Input As #numb
n = 0

Do Until EOF(numb)
  n = n + 1
  Line Input #numb, zeile
  teile = Split(zeile, ",")
  words(n) = teile(0)
  answer(n) = teile(1)
   If UBound(teile) - LBound(teile) + 1 = 3 Then
   bewertung(n) = CInt(teile(2))
   Else
   bewertung(n) = 0
   End If
Loop
Close #numb

anzWords = n

Listloaded = True
lblListLoaded.Caption = txtList.Text
MsgBox "Erfolgreich geladen!", vbInformation
End Sub
Private Sub cmdStop_Click()
'Resultate ausgeben
If frmVociTrainer.Width = 4420 Then
Call cmbExpand_Click
End If

lblRichtige.Caption = "Richtige: " & richtige
lblFalse.Caption = "Falsche: " & gefragte - richtige
lblBewertung.Caption = "Bewertung: Note " & Runden((richtige / gefragte) * 5 + 1, 1)

'clear
gefragte = 0
richtige = 0
txtAnswer.Text = ""
txtQuery.Text = ""
lblAnzeige.Caption = "Keine Aktion"
lblAnzeige.FontItalic = True
optUeben.Enabled = True
optAbfrage.Enabled = True
optNormal.Enabled = True
optReversed.Enabled = True
optRandom.Enabled = True
optDescending.Enabled = True
cmbExercise.Enabled = True
frmVociTrainer.SetFocus
cboNumber.Enabled = True
modeAsk = False

If optReversed.Value = True Then
  For i = 1 To anzWords
  temp(i) = answer(i)
  answer(i) = words(i)
  words(i) = temp(i)
  Next
End If

'Speichern
numb = FreeFile
Kill txtList.Text
Open txtList.Text For Output As #numb
For i = 1 To anzWords
Print #numb, words(i) & "," & answer(i) & "," & bewertung(i)
Next
Close #numb
cmdStop.Enabled = False
frmVociTrainer.SetFocus

End Sub
Private Sub Form_Load()
frmVociTrainer.Width = 4420
Randomize Timer
End Sub
Private Sub mnuBeenden_Click()
End
End Sub
Private Sub mnuHilfe_Click()
MsgBox "Sie benutzen das Voci-Programm zur Verwaltung und Abfrage von Vokabular in der Version 0.1. Vielen Dank für Ihre Nutzung!", vbInformation
End Sub
Private Sub mnuSettings_Click()
frmSettings.Show
End Sub
Private Sub mnuVersion_Click()
MsgBox "Aktuelle Version: V. 0.1. Datum: 13.03.2009", vbInformation
End Sub
Private Sub txtAnswer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And modeAsk = True Then
'1. Mal
If txtQuery.Text = "" Then
optUeben.Enabled = False
optAbfrage.Enabled = False
optNormal.Enabled = False
optReversed.Enabled = False
optRandom.Enabled = False
optDescending.Enabled = False
If optDescending.Value = True Then wort = 0
Else
'auswertung
  If txtAnswer.Text = answer(wort) Then
  richtige = richtige + 1
  lblAnzeige.Caption = "Richtig!"
  bewertung(wort) = bewertung(wort) + 1
  Else
  lblAnzeige.Caption = "Falsch. Richtig wäre: " & answer(wort)
  bewertung(wort) = bewertung(wort) - 1
  End If
gefragte = gefragte + 1
End If

If gefragte = wievielWoerterAbfragen Then
Call cmdStop_Click
Exit Sub
End If

'Üben

If optRandom.Value = True Then wort = Rand(1, anzWords)

If optRandom.Value = True And chkDiff.Value = True Then
  For i = 1 To anzWords
  wort = Rand(1, anzWords)
  chance = Rand(1, 10)
     If (chance > 8 And bewertung(wort) > mittel) Or bewertung(wort) < mittel Then
     Exit For
     End If
  Next
End If

If optDescending.Value = True Then wort = wort + 1


txtQuery.Text = words(wort)
txtAnswer.SetFocus
If optAbfrage.Value = True Then
txtAnswer.Text = ""
Else
txtAnswer.Text = answer(wort)
End If

End If
End Sub
Private Sub txtList_Click()
If txtList.Text = "Keine Liste ausgewählt" Then
txtList.Text = App.Path & "\"
txtList.SelStart = Len(txtList.Text)
txtList.FontItalic = False
End If
End Sub
Function Runden(Zahl As Double, Dezimalanzahl As Integer) As Double
    Runden = Int(Zahl * 10 ^ Dezimalanzahl + 0.5) / 10 ^ Dezimalanzahl
End Function
Private Sub txtList_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Call cmbLoadList_Click
End If
End Sub
Private Sub txtList_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0
End Sub
