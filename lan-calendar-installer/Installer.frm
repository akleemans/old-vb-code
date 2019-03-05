VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   2880
   Icon            =   "Installer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   2880
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MsgBox "Herzlich willkommen! Der Lan-Kalender wird nun installiert.", vbYesNo, "Installation Lan-Kalender"
Shell "copy App.path\run\lankalender.exe C:\Programme\Lan-Kalender\lankalender.exe"
Shell "del App.path\run\lankalender.exe"
Shell "copy App.path\run\lankalender.lnk C:\Programme\Lan-Kalender\lankalender.lnk"
Shell "del App.path\run\lankalender.lnk"
SubKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
Reg_SetString HKEY_LOCAL_MACHINE, SubKey, "Lan-Kalender", "C:\Programme\Lan-Kalender\lankalender.exe"
MsgBox "Installation erfolgreich abgeschlossen! Dieses Programm ist Freeware und wurde von FoN_qwx erstellt.", vbInformation
End Sub
