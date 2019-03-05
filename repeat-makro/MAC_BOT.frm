VERSION 5.00
Begin VB.Form frmMAC_BOT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MAC_BOT"
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboAuswahl 
      Height          =   315
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   3375
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   600
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   600
   End
   Begin VB.CommandButton cmbGo 
      Caption         =   "Execute!"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblPosition 
      Caption         =   "Position: X / Y"
      Height          =   255
      Left            =   4320
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblTime 
      Caption         =   "Time: 00:00:00"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Pfad:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmMAC_BOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'API-Funktion deklarieren
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI 'Variablentyp deklarieren
   X As Long
   Y As Long
End Type

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

Dim CursorPos As POINTAPI 'Variable deklarieren
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Dim startzeit As Double
Dim Zeile, Var1, Var2, Var3, befehle(1 To 1000), tue() As String
Dim countZeile, i, j As Integer
Private Sub cmbChangePath_Click()
CommonDialog1.ShowOpen
txtPath.Text = CommonDialog1.FileName
End Sub
Private Sub cmbGo_Click()
Call load
startzeit = 0

'Zeit los!
Timer1.Enabled = True
If startzeit = 0 Then startzeit = Timer

For i = 1 To countZeile - 1
tue() = Split(befehle(i), " ")

Do While CDbl(tue(0)) > (Timer - startzeit)
'Label2.Caption = Label2.Caption & CDbl(tue(0)) & " / " & (Timer - startzeit) & " "
Wait (CDbl(tue(0)) - (Timer - startzeit) - 0.1)
DoEvents
Loop

If CDbl(tue(0)) <= (Timer - startzeit) Then
    
    Select Case tue(1)
        Case "CLICK"
        Call SetCursorPos(CInt(tue(2)), CInt(tue(3)))
        Wait 10
        MouseClick LeftMouseButton
        
        Case "MOVETO"
        Call SetCursorPos(CInt(tue(2)), CInt(tue(3)))
        
        Case "LEFT_CLICK"
        MouseClick LeftMouseButton
        
        Case "LEFT_DOUBLE"
        MouseClick LeftMouseButton
        Wait 10
        MouseClick LeftMouseButton
       
        Case "LEFT_TRIPLE"
        MouseClick LeftMouseButton
        Wait 10
        MouseClick LeftMouseButton
        Wait 10
        MouseClick LeftMouseButton
        
        Case "LEFT_DOWN"
        MouseDown LeftMouseButton
        
        Case "LEFT_UP"
        MouseUp LeftMouseButton
        
        Case "RIGHT_CLICK"
        MouseClick RightMouseButton
        
        Case "PRINT"
        SendKeys tue(2)
           
        Case "PRINT_HOUR"
        SendKeys Int(Timer / 3600) + 1
        
        Case "PRINT_CLIP1"
        SendKeys Var1
        
        Case "PRINT_CLIP2"
        SendKeys Var2
        
        Case "PRINT_CLIP3" 'Formatierter Clip
        If Len(Var3) > 58 Then
        Var3 = Left(Var3, 58)
        End If
        'String säubern
        For j = 1 To Len(Var3)
        If Mid(Var3, j, 1) = "(" Then Mid(Var3, j, 1) = " "
        If Mid(Var3, j, 1) = ")" Then Mid(Var3, j, 1) = " "
        Next
        SendKeys Var3
        
        Case "GET_CLIP1"
        Var1 = Clipboard.GetText()
        
        Case "GET_CLIP2"
        Var2 = Clipboard.GetText()
        
        Case "GET_CLIP3" 'Formatierter Clip
        Var3 = Clipboard.GetText()
        
    End Select
End If
Next
MsgBox "Finished!", vbInformation
Timer1.Enabled = False
End Sub
Private Sub cmbMouse_Click()
If Timer2.Enabled = False Then
Timer2.Enabled = True
Else
Timer2.Enabled = False
End If
End Sub
Private Sub load()
countZeile = 0
'Datei öffnen
Dim myPath As String
myPath = "D:\Coding\MAC_BOT\makros\" & cboAuswahl.Text & ".do"

Open myPath For Input As #1
Do Until EOF(1)
Line Input #1, Zeile
'MsgBox Zeile
    If Left(Zeile, 1) <> "%" And Zeile <> "" Then
        If Left(Zeile, 1) = "<" Then
        
        End If
    countZeile = countZeile + 1
    befehle(countZeile) = Zeile
    End If
Loop
Close #1

End Sub
Private Sub Form_Load()
cboAuswahl.AddItem "Makro Example"
End Sub

Private Sub Timer1_Timer()
Timer2.Enabled = False
If startzeit = 0 Then startzeit = Timer
lblTime.Caption = "Time: " & Round(CStr(Timer - startzeit), 1)
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
Private Sub Timer2_Timer()
Call GetCursorPos(CursorPos) 'API-Funktion aufrufen
lblPosition.Caption = "Position: " & CursorPos.X & " / " & CursorPos.Y
'----------------------- Abbrechen der Applikation funktioniert noch nicht
'If CBool(GetAsyncKeyState(Asc("q")) And &H8000) Then End
'-----------------------
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
