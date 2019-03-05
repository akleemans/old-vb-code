VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImport 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ACL-Dateien zu Excel konvertieren"
   ClientHeight    =   2640
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4800
   Icon            =   "frmImport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmbChose2 
      Caption         =   "Wählen..."
      Height          =   285
      Left            =   3600
      TabIndex        =   7
      Top             =   1320
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   4080
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmSave 
      Caption         =   "Dateiname für neue Excel-Datei:"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   4575
      Begin VB.TextBox txtSave 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame frameImport 
      Caption         =   "Bitte die zu konvertierende ACL-Datei auswählen"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4575
      Begin VB.CommandButton cmbChose 
         Caption         =   "Wählen..."
         Height          =   285
         Left            =   3480
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtPfad 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmbImport 
      Caption         =   "ACL Importieren"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   1850
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4800
      Y1              =   2320
      Y2              =   2320
   End
   Begin VB.Label lblStatus 
      Caption         =   "Bereit"
      Height          =   255
      Left            =   0
      LinkTimeout     =   5
      TabIndex        =   2
      Top             =   2400
      Width           =   4815
   End
   Begin VB.Menu mnuDatei 
      Caption         =   "Datei"
      Index           =   1
      Begin VB.Menu mnuSettings 
         Caption         =   "Einstellungen"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "?"
      Index           =   2
      Begin VB.Menu mnuHilfe 
         Caption         =   "Hilfe"
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "Info..."
      End
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Coded August 2008 by KlAd
Option Explicit
Dim Datei, wort, zeichen, z As String
Dim inhalt As String
Dim zeile, bf, spalte, x As Integer
Dim i, i2 As Double
Dim v As Variant

Dim xlApp As Excel.Application
Dim xlWB As Excel.Workbook
Dim xlWS As Excel.Worksheet

Private Sub cmbChose_Click()
CommonDialog.ShowOpen
txtPfad.Text = CommonDialog.FileName
End Sub

Private Sub cmbChose2_Click()
CommonDialog.ShowSave
txtSave.Text = CommonDialog.FileName
End Sub
Private Sub cmbImport_Click()

'Errorhandling
On Error GoTo fehler

'Variablendeklaration
Datei = txtPfad.Text
spalte = 1
zeile = 1
bf = FreeFile
i = 0
i2 = 0
wort = ""

'Datei muss ACL-Format haben
If Datei = "" Or UCase$(Right(Datei, 4)) <> ".ACL" Then
MsgBox "Pfadangabe ungültig. Bitte geben Sie den vollständigen Pfad zur .ACL-Datei ein.", vbCritical
Exit Sub
ElseIf txtSave.Text = "" Then
MsgBox "Bitte geben Sie einen Namen für die zu erstellende Datei ein.", vbCritical
Exit Sub
End If

'Excel neu gestalten
Set xlApp = New Excel.Application
Set xlWB = xlApp.Workbooks.Add
Set xlWS = xlWB.Worksheets.Add

'Datei binär öffnen und EOF()-Zeichen löschen
lblStatus.Caption = "Lese Datei in Speicher ein..."
inhalt = removeChr26(1)
i = 0

'Filterdurchgange
Call Filter

'Datei lesen
Call FileWrite
  
'Excel speichern
xlWS.SaveAs txtSave.Text & ".xls"
xlApp.Quit
    
'Memory leeren
Set xlWS = Nothing
Set xlWB = Nothing
Set xlApp = Nothing

'Abschliessen
lblStatus.Caption = "Datei geschlossen. Vorgang abgeschlossen."
MsgBox "Vorgang abgeschlossen, Datei gespeichert.", vbInformation
GoTo ende

'Bei Fehler:
fehler:
MsgBox "Kritischer Fehler: " & Err.Description & " - Error number " & Err.Number, vbCritical
lblStatus.Caption = "Kritischer Fehler aufgetreten. Vorgang abgebrochen."

ende:
End Sub
Function chg()
If spalte = 1 Then
spalte = 2
ElseIf spalte = 2 Then
spalte = 1
zeile = zeile + 1
End If
End Function
Function removeChr26(x As Integer)
Dim inhalt As String

'Einlesen
Open Datei For Binary Access Read As #bf
inhalt = Space$(LOF(bf))
Get #bf, , inhalt
Close #bf

'Auswerten
For i = 1 To Len(inhalt)
    If Asc(Mid(inhalt, i, 1)) = 26 Then Mid(inhalt, i, 1) = Chr(124)
Next

removeChr26 = inhalt
End Function
Function Filter()
Dim showme As Integer
Dim check23, check28, check30 As Integer
showme = 0

lblStatus.Caption = "Filtere Datei auf Sonderzeichen..."

  While Len(inhalt) > i
    i = i + 1
    
    If i2 < (i - (Len(inhalt) / 100)) Then
    i2 = i
    lblStatus.Caption = "Filtere...   " & Round((i2 / Len(inhalt)) * 100, 1) & " %"
    Me.Refresh
    End If
    
    z = Mid(inhalt, i, 1)
    
    If showme = 1 Then
    MsgBox "Nummer: " & Asc(z) & " (Zeichen: " & z & ")"
    End If
    
    'Falschzeichen eliminieren
    '3 Typen von Zeichen:
    '1. Legale, im Wort enthaltene Zeichen
    '2. Zeichen, die ignoriert werden ==> Chr(95) '_'
    '3. Zeilenumbruchzeichen ==> Chr(124) '|'

    If Asc(z) = 0 Then
        Mid(inhalt, i, 1) = Chr(95)
        GoTo onceagain
    ElseIf Asc(z) = 9 Then
        'MsgBox "Tab-Zeichen (asc9) kommt an Stelle " & z & " vor."
        Mid(inhalt, i, 1) = " "
    ElseIf (Asc(z) >= 2 And Asc(z) <= 7) Or (Asc(z) >= 10 And Asc(z) <= 13) Then
        Mid(inhalt, i, 1) = Chr(124) 'Chr(95)
        GoTo onceagain
    ElseIf Asc(z) = 23 Then
        Mid(inhalt, i, 1) = "_"
        GoTo onceagain
    ElseIf Asc(z) = 24 Or Asc(z) = 25 Then
        Mid(inhalt, i, 1) = "'"
        GoTo onceagain
    ElseIf Asc(z) = 95 Then
        Mid(inhalt, i, 1) = Chr(45)
        GoTo onceagain
    Else
        GoTo onceagain
    End If
onceagain:
  Wend
  
End Function
Function FileWrite()
Dim permissionToWrite
i = 0
i2 = 0
permissionToWrite = 0

lblStatus.Caption = "Schreibe Zeilen aus Speicher..."

  While Len(inhalt) > i
    i = i + 1
    zeichen = Mid(inhalt, i, 1)

        If i2 < (i - (Len(inhalt) / 100)) Then
        i2 = i
        lblStatus.Caption = "Schreibe in Datei...   " _
        & Round((i2 / Len(inhalt)) * 100, 0) _
        & " % (Zeit verbleibend: " & Round((Len(inhalt) / 50000) _
        - ((Len(inhalt) / 50000) / Len(inhalt) * i), 0) & " s)"
        Me.Refresh
        End If
        
'Zu ignorierende Zeichen
  If Asc(zeichen) = 95 And i >= 3 Then
    If Mid(inhalt, i - 2, 3) = "___" And i < Len(inhalt) Then
    Mid(inhalt, i + 1, 1) = "|"
    End If
    GoTo igno
    
'Umbruchsgenerierende Zeichen
    ElseIf Asc(zeichen) = 124 Then
    If permissionToWrite = 1 Then
    xlWS.Cells(zeile, spalte).Value = wort
    Call chg
    End If
    wort = ""
    
'Normales Zeichen
    Else
        'Schreibmodus aktivieren
        If Asc(zeichen) = 8 And permissionToWrite = 0 Then
        permissionToWrite = 1
        wort = ""
        zeichen = ""
        End If
    wort = wort & zeichen
   End If

igno:
  Wend
End Function
Private Sub mnuHilfe_Click()
MsgBox "Das Programm konvertiert Auto Correct Lists, kurz ACL, zu Microsoft Excel-Dateien. Oben muss der volle Pfad der ACL-Datei eingegeben werden, unten reicht ein Dateiname mit Pfad. Unter diesem Namen wird die Excel-Tabelle dann gespeichert.", vbInformation
End Sub

Private Sub mnuInfo_Click()
MsgBox "Programmiert von Adrianus Kleemans. Bei Fragen und oder Fehlermeldungen bitte die Vorgehensweise inkl. Fehlercode an a.kleemans@gmail.com senden."
End Sub
Private Sub mnuSettings_Click()
frmSettings.Show
End Sub
