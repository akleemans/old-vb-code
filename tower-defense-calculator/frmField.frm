VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmField 
   Caption         =   "Tower Defense Strategy Strength Calculater"
   ClientHeight    =   3135
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmdDialog 
      Left            =   5040
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtPath 
      ForeColor       =   &H80000011&
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Enter Path..."
      Top             =   240
      Width           =   5295
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblAnzeige 
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   5175
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu mNew 
         Caption         =   "New..."
      End
   End
   Begin VB.Menu mEtc 
      Caption         =   "?"
      Begin VB.Menu mInfo 
         Caption         =   "Info"
      End
   End
End
Attribute VB_Name = "frmField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tower1 As New tower
Dim strMap As String
'Dim tower As String 'Klasse definieren!
'Eigenschaften:
'.Name
'F  Fire
'P  Poison
'R  Rocket
'T  Thunderbolt
'.Damage
'.LoadingTime
'.Impact
'.InfluenceOnVelocity

Private Sub cmdCalculate_Click()

'Map auslesen
Open txtPath.Text For Binary As #1
strMap = Space$(LOF(1))
Get #1, , strMap
Close #1
MsgBox strMap

'In Koordinaten umsetzen


End Sub

Private Sub txtPath_Click()
txtPath.Text = ""
txtPath.ForeColor = &H80000008
cmdDialog.ShowOpen
txtPath.Text = cmdDialog.FileName
End Sub

'1. Einlesen
'2. In koordinaten umwandeln
'3. Beschreitungs- und Verlassenszeitpunkt jedes Quadrats ausrechnen
'4. Schaden zusammenrechnen

