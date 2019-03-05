VERSION 5.00
Begin VB.Form frmCheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   3630
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3000
      Top             =   2400
   End
   Begin VB.CommandButton cmbPlay 
      Caption         =   "play"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmbStop 
      Caption         =   "stop"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmbMutate 
      Caption         =   "Datensatz verändern"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2520
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmbDelete 
      Caption         =   "Markierte löschen"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton cmbRecord 
      Caption         =   "rec"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblStatus 
      Caption         =   "Mode: nothing"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblElemente 
      Caption         =   "Anzahl Elemente: 0"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1935
   End
End
Attribute VB_Name = "frmCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetKeyState Lib "user32.dll" ( _
    ByVal nVirtKey As Long _
) As Integer

Private Declare Function GetAsyncKeyState Lib "user32.dll" ( _
    ByVal vKey As Long _
) As Integer

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Sub keybd_event Lib "user32" (ByVal _
        bVk As Byte, ByVal bScan As Byte, ByVal dwFlags _
        As Long, ByVal dwExtraInfo As Long)
        
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

Const VK_LWIN = &H5B
Const VK_APPS = &H5D
Const KEYEVENTF_KEYUP = &H2


Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101

'Private Const VK_MENU As Long = &H12&
'Private Const VK_SHIFT As Long = &H10&
'Private Const VK_CONTROL As Long = &H11&
'Private Const VK_CAPITAL As Long = &H14&

Dim text, textneu, KeyUp, KeyDown, KeyLeft, KeyRight As String
Dim x, i, i1, AnzEntry, UpOrDown(1 To 1000), KeyEntry(1 To 1000) As Integer
Dim StartZeit, TheTime(1 To 1000) As Double
Private Sub cmbPlay_Click()
Timer1.Enabled = False
lblStatus.Caption = "Mode: play"
'vbKeyLeft = 37
'vbKeyUp = 38
'vbKeyRight= 39
'vbKeyDown = 40

'KEYEVENTF_KEYDOWN = 0   Tastendruck senden
'KEYEVENTF_KEYUP   = 2

i = 1
StartZeit = Timer
Wait 3000
MouseClick LeftMouseButton
Wait 10
Do While i <= AnzEntry
    If (Timer - StartZeit) >= TheTime(i) Then
    keybd_event KeyEntry(i), 0, UpOrDown(i), 0
    i = i + 1
        If i = AnzEntry Then GoTo finish
    End If
DoEvents
Loop

finish:
MsgBox "Done", vbInformation
End Sub
Private Sub cmbStop_Click()
Timer1.Enabled = False
lblStatus.Caption = "Mode: normal"
End Sub
Private Sub Form_Load()
AnzEntry = 0
End Sub
Private Sub cmbRecord_Click()
'Zeit starten
lblStatus.Caption = "Mode: rec"
KeyUp = "up"
KeyDown = "up"
KeyLeft = "up"
KeyRight = "up"

StartZeit = Timer
Wait 3000
MouseClick LeftMouseButton
Wait 10

'Tastendruck registrieren
Timer1.Enabled = True
End Sub
Public Function KeyPressed( _
    ByVal Key As KeyCodeConstants, _
    Optional ByVal Wait As Boolean = False _
  ) As Boolean
  
  'Status feststellen:
  KeyPressed = CBool(GetAsyncKeyState(Key) And &H8000)
  
  'Ggf. auf Loslassen warten:
  Wait = Wait And KeyPressed
  If Wait Then
    Do While CBool(GetAsyncKeyState(Key) And &H8000)
    Loop
  End If
End Function
Private Sub Timer1_Timer()
'up/down und left/right kommen komplementär vor

'Up
If CBool(GetAsyncKeyState(vbKeyUp) And &H8000) And KeyUp = "up" Then 'taste wird gedrückt
AnzEntry = AnzEntry + 1
KeyEntry(AnzEntry) = 38
UpOrDown(AnzEntry) = 0
TheTime(AnzEntry) = Round(Timer - StartZeit, 2)
List1.AddItem "KeyUp press " & Round(Timer - StartZeit, 2)
KeyUp = "down"

ElseIf CBool(GetAsyncKeyState(vbKeyUp) And &H8000) = False And KeyUp = "down" Then 'taste wird losgelassen
AnzEntry = AnzEntry + 1
KeyEntry(AnzEntry) = 38
UpOrDown(AnzEntry) = 2
TheTime(AnzEntry) = Round(Timer - StartZeit, 2)
List1.AddItem "KeyUp release " & Round(Timer - StartZeit, 2)
KeyUp = "up"
End If

'Down
If CBool(GetAsyncKeyState(vbKeyDown) And &H8000) And KeyDown = "up" Then
AnzEntry = AnzEntry + 1
KeyEntry(AnzEntry) = 40
UpOrDown(AnzEntry) = 0
TheTime(AnzEntry) = Round(Timer - StartZeit, 2)
List1.AddItem "KeyDown press " & Round(Timer - StartZeit, 2)
KeyDown = "down"

ElseIf CBool(GetAsyncKeyState(vbKeyDown) And &H8000) = False And KeyDown = "down" Then
AnzEntry = AnzEntry + 1
KeyEntry(AnzEntry) = 40
UpOrDown(AnzEntry) = 2
TheTime(AnzEntry) = Round(Timer - StartZeit, 2)
List1.AddItem "KeyDown release " & Round(Timer - StartZeit, 2)
KeyDown = "up"
End If

'Left
If CBool(GetAsyncKeyState(vbKeyLeft) And &H8000) And KeyLeft = "up" Then
AnzEntry = AnzEntry + 1
KeyEntry(AnzEntry) = 37
UpOrDown(AnzEntry) = 0
TheTime(AnzEntry) = Round(Timer - StartZeit, 2)
List1.AddItem "KeyLeft press " & Round(Timer - StartZeit, 2)
KeyLeft = "down"

ElseIf CBool(GetAsyncKeyState(vbKeyLeft) And &H8000) = False And KeyLeft = "down" Then
AnzEntry = AnzEntry + 1
KeyEntry(AnzEntry) = 37
UpOrDown(AnzEntry) = 2
TheTime(AnzEntry) = Round(Timer - StartZeit, 2)
List1.AddItem "KeyLeft release " & Round(Timer - StartZeit, 2)
KeyLeft = "up"
End If

'Right
If CBool(GetAsyncKeyState(vbKeyRight) And &H8000) And KeyRight = "up" Then
AnzEntry = AnzEntry + 1
KeyEntry(AnzEntry) = 39
UpOrDown(AnzEntry) = 0
TheTime(AnzEntry) = Round(Timer - StartZeit, 2)
List1.AddItem "KeyRight press " & Round(Timer - StartZeit, 2)
KeyRight = "down"

ElseIf CBool(GetAsyncKeyState(vbKeyRight) And &H8000) = False And KeyRight = "down" Then
AnzEntry = AnzEntry + 1
KeyEntry(AnzEntry) = 39
UpOrDown(AnzEntry) = 2
TheTime(AnzEntry) = Round(Timer - StartZeit, 2)
List1.AddItem "KeyRight release " & Round(Timer - StartZeit, 2)
KeyRight = "up"
End If

lblElemente.Caption = "Anzahl Elemente: " & AnzEntry
End Sub
Private Sub Wait(ByVal ms As Long)
  'Sleep without freezing
  If ms <= 40 Then
    Sleep ms
  Else
    Dim i As Long, j As Long
    j = ms \ 40
    For i = 1 To j
      Sleep 40
      DoEvents
    Next i
    Sleep ms - j * 40
  End If
  DoEvents
End Sub
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
