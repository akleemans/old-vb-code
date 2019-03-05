VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmCheck 
   Caption         =   "Check"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmbZurueck 
      Caption         =   "zurück"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   180
      Width           =   735
   End
   Begin VB.CommandButton cmbWeiter 
      Caption         =   "Nächste Kombination..."
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4335
      ExtentX         =   7646
      ExtentY         =   5530
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label lblAnzeige 
      Caption         =   "i = 0"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim zaehler, numbers(1 To 100) As Integer
Private Sub Form_Load()
Dim i, mul, c As Integer
Dim iString As String
c = 1

For i = 1126 To 6211
iString = CStr(i)
mul = CInt(Mid(iString, 1, 1)) * CInt(Mid(iString, 2, 1)) * CInt(Mid(iString, 3, 1)) * CInt(Mid(iString, 4, 1))
   
 If mul = 12 Then
 numbers(c) = i
 c = c + 1
 End If
Next
End Sub
Private Sub cmbWeiter_Click()
 If lblAnzeige.Caption = "i = 0" Then
 zaehler = 0
 ElseIf numbers(zaehler) = 6211 Then
 Exit Sub
 End If
zaehler = zaehler + 1
 lblAnzeige.Caption = "i = " & numbers(zaehler)
 WebBrowser1.Navigate ("http://www.isatcis.com/" & numbers(zaehler) & ".htm")
End Sub
Private Sub cmbZurueck_Click()
 If lblAnzeige.Caption = "i = 1126" Or lblAnzeige.Caption = "i = 0" Then
 Exit Sub
 End If
zaehler = zaehler - 1
 lblAnzeige.Caption = "i = " & numbers(zaehler)
 WebBrowser1.Navigate ("http://www.isatcis.com/" & numbers(zaehler) & ".htm")
End Sub
