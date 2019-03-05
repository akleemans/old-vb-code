VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form frmLanKalender 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lan-Kalender"
   ClientHeight    =   3150
   ClientLeft      =   150
   ClientTop       =   765
   ClientWidth     =   9690
   Icon            =   "frmLanKalender.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmbaktualisieren 
      Caption         =   "Aktualisieren"
      Height          =   255
      Left            =   3960
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   2580
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   9480
      ExtentX         =   16722
      ExtentY         =   4551
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Menu mnuDatei 
      Caption         =   "Datei..."
      Begin VB.Menu mnumelden 
         Caption         =   "Lan melden"
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "Info"
      End
      Begin VB.Menu mnuBeenden 
         Caption         =   "Beenden"
      End
   End
End
Attribute VB_Name = "frmLankalender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbaktualisieren_Click()
brwWebBrowser.Refresh
End Sub
Private Sub Form_Load()
brwWebBrowser.Navigate "http://home.tiscalinet.ch/aetlk/kalander.html"
End Sub
Private Sub mnuBeenden_Click()
End
End Sub
Private Sub mnuInfo_Click()
MsgBox "Version 0.2 - Coded 22. Juni 2005 by FoN_qwx"
End Sub
Private Sub mnumelden_Click()
MsgBox "Bitte an a.kleemans@gmx.ch schreiben."
End Sub
