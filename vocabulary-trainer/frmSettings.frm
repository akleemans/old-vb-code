VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Einstellungen"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   2640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FRAME1 
      Caption         =   "Einstellungen"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.CheckBox chkList 
         Caption         =   "Liste speichern"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmbOK 
         Caption         =   "OK"
         Height          =   300
         Left            =   840
         TabIndex        =   1
         Top             =   1200
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbOK_Click()
'speichern in ini-file

'"lastList=" & frmVociTrainer.lblListloaded.Caption
'"loadOnStartup=" & chkList.Value

frmSettings.Hide
End Sub
