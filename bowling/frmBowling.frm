VERSION 5.00
Begin VB.Form frmBowling 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bowling"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   3300
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll4 
      Height          =   255
      Left            =   120
      Max             =   500
      Min             =   -500
      TabIndex        =   10
      Top             =   2880
      Width           =   2415
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   120
      Max             =   500
      TabIndex        =   7
      Top             =   2280
      Value           =   300
      Width           =   2415
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   120
      Max             =   100
      TabIndex        =   5
      Top             =   1680
      Value           =   30
      Width           =   2415
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      Max             =   250
      Min             =   -250
      TabIndex        =   1
      Top             =   1080
      Value           =   25
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bowling-Wurf ausführen"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "0"
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Drehung"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Geschwindigkeit:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "300"
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "30"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Verschiebung vertikal:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lblGeschw 
      Caption         =   "25"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Verschiebung horizontal:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "frmBowling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' zunächst die benötigten API-Deklarationen
Private Declare Sub mouse_event Lib "user32" ( _
  ByVal dwFlags As Long, _
  ByVal dx As Long, _
  ByVal dy As Long, _
  ByVal cButtons As Long, _
  ByVal dwExtraInfo As Long)
 
Private Const MOUSEEVENTF_ABSOLUTE = &H8000
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_MOVE = &H1
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Dim rechts, hinunter, kugel_down, kugel_links, geschw, drehung As Integer
Dim screenDef As String
Public Sub Mausklick(Optional Button As _
  MouseButtonConstants = vbLeftButton, _
  Optional XPos As Long = -1, _
  Optional YPos As Long = -1)
 
  ' Mauszeiger positionieren
  If XPos <> -1 Or YPos <> -1 Then
    mouse_event MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_MOVE, _
    XPos / Screen.Width * 65535, _
    YPos / Screen.Height * 65535, 0, 0
  End If
 
  ' Mausklick simulieren
  Select Case Button
    ' linke Maustaste
    Case vbLeftButton
      mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
      mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
 
    ' mittlere Maustaste
    Case vbMiddleButton
      mouse_event MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0
      mouse_event MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0
 
    ' rechte Maustaste
    Case vbRightButton
      mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
      mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
  End Select
End Sub
Public Function MoveMouse(ByVal rechts As Integer, ByVal hinunter As Integer) As Boolean
Dim X, Y, rechts1 As Integer
rechts1 = rechts

    For Y = 0 To 400 Step 6 'Gibt die Höhe an
    Call SetCursorPos(rechts, hinunter - Y)
    rechts1 = rechts1 '+ (Y / (100000 / drehung)) - 1
    rechts = Round(rechts1, 0)
    Sleep Round(1000 / geschw, 0)
    Next
End Function
Private Sub Command1_Click()
Dim i As Integer

'=========
screenDef = "arbeit"
kugel_links = -HScroll1.Value
kugel_down = HScroll2.Value
geschw = HScroll3.Value
drehung = HScroll4.Value + 1
'=========

If screenDef = "zuhause" Then
rechts = 600
hinunter = 600
ElseIf screenDef = "arbeit" Then
rechts = 1680 + 800
hinunter = 620
End If

Call SetCursorPos(rechts, hinunter)
Sleep 200
Mausklick vbLeftButton
Sleep 500
mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
Call SetCursorPos(rechts - kugel_links, hinunter + kugel_down)
Sleep 200
Call MoveMouse(rechts - kugel_links, hinunter + kugel_down)
Sleep 100
mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Private Sub HScroll1_Change()
lblGeschw.Caption = HScroll1.Value
End Sub
Private Sub HScroll2_Change()
Label3.Caption = HScroll2.Value
End Sub
Private Sub HScroll3_Change()
Label4.Caption = HScroll3.Value
End Sub
Private Sub HScroll4_Change()
Label7.Caption = HScroll4.Value
End Sub

