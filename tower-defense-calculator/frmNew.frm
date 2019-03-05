VERSION 5.00
Begin VB.Form frmNew 
   Caption         =   "New Testfield"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3405
   LinkTopic       =   "Form2"
   ScaleHeight     =   7140
   ScaleWidth      =   3405
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmNew.frx":0000
      Left            =   2520
      List            =   "frmNew.frx":0007
      TabIndex        =   2
      Text            =   "11"
      Top             =   200
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmNew.frx":000E
      Left            =   840
      List            =   "frmNew.frx":0015
      TabIndex        =   1
      Text            =   "11"
      Top             =   200
      Width           =   735
   End
   Begin VB.CommandButton cmbCreate 
      Caption         =   "Create"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label lblWidth 
      Caption         =   "Width:"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblHeight 
      Caption         =   "Height:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
