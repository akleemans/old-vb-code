VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2010
   LinkTopic       =   "Form1"
   ScaleHeight     =   765
   ScaleWidth      =   2010
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   37000
      Left            =   1680
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cbuttons As Long, ByVal dwExtraInfo As Long)

Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10
Public Enum MouseButtons
   LeftMouseButton
   RightMouseButton
   MiddleMouseButton
End Enum
Dim i As Integer

Public Sub MouseUp(MouseButton As MouseButtons)
   Select Case (MouseButton)
      Case LeftMouseButton
         Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
      Case MiddleMouseButton
         Call mouse_event(MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0)
      Case RightMouseButton
         Call mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
   End Select
End Sub

Public Sub MouseDown(MouseButton As MouseButtons)
   Select Case (MouseButton)
      Case LeftMouseButton
         Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
      Case MiddleMouseButton
         Call mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0)
      Case RightMouseButton
         Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
   End Select
End Sub

Public Sub MouseClick(MouseButton As MouseButtons)
   MouseDown (MouseButton)
   MouseUp (MouseButton)
End Sub
Private Sub Command1_Click()
Timer1.Enabled = True
Timer2.Enabled = True
End Sub
Private Sub Timer1_Timer()
MouseClick LeftMouseButton
MouseClick LeftMouseButton
MouseClick LeftMouseButton
MouseClick LeftMouseButton
MouseClick LeftMouseButton
MouseClick LeftMouseButton
MouseClick LeftMouseButton
MouseClick LeftMouseButton
MouseClick LeftMouseButton
MouseClick LeftMouseButton
End Sub
Private Sub Timer2_Timer()
Timer1.Enabled = False
End Sub
