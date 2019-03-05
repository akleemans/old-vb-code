VERSION 5.00
Begin VB.Form frmClean 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Clean Profile Space"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmClean.frx":0000
   ScaleHeight     =   64.836
   ScaleMode       =   0  'User
   ScaleWidth      =   605.028
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   6855
   End
   Begin VB.CommandButton cmbClean 
      Caption         =   "Clean"
      Height          =   375
      Left            =   7200
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "frmClean"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Netzwerk As Object
Dim anzDateien As Integer
Dim pfad2 As String
Dim FSO As New FileSystemObject
Dim Folder As Folder
Dim sFolderPath As String
Dim sDestPath As String

Private Sub cmbClean_Click()
MsgBox "Bitte Firefox schliessen.", vbInformation

'Search
anzDateien = 0
'Call getFiles   'Files herausfinden

'Dateien verschieben
pfad2 = "\\serverpath\userhome$\" & Netzwerk.UserName & "\"

'---------------
' Ordner kopieren
' Welcher Ordner soll kopiert werden?
sFolderPath = txtPath.Text
' Wohin soll der Ordner kopiert werden?
sDestPath = "\\serverpath\userhome$\" & Netzwerk.UserName & "\"

' Kopiervorgang starten
Set Folder = FSO.GetFolder(sFolderPath)

'Folder.Copy sDestPath
'---------------------

' Welcher Ordner soll gelöscht werden?
sFolderPath = txtPath.Text
' Löschvorgang starten
Set Folder = FSO.GetFolder(sFolderPath)
' alles löschen
RmDir Folder
'Folder.Delete True
' Ordner/Dateien nur löschen, wenn diese nicht schreibgeschützt sind
'Folder.Delete False
'---------------------

'Unbennenen
Name "\\serverpath\userhome$\" & Netzwerk.UserName & "\Application Data" As "\\serverpath\userhome$\" & Netzwerk.UserName & "\_Backup"

'Found
MsgBox "Fertig!", vbInformation
End Sub
Private Sub Form_Load()
Set Netzwerk = CreateObject("wscript.network")

txtPath.Text = "C:\Documents and Settings\" & Netzwerk.UserName & "\Application Data\"
End Sub
