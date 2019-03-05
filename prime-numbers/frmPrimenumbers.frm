VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrimenumbers 
   Caption         =   "Prime numbers"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optNew 
      Caption         =   "New"
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   2880
      Width           =   735
   End
   Begin VB.OptionButton optLoad 
      Caption         =   "Load List"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.CheckBox chkClean 
      Caption         =   "Clean list"
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   4680
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1680
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLoadList 
      Caption         =   "Load list"
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtEnd 
      Height          =   285
      Left            =   3360
      TabIndex        =   7
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txtStart 
      Height          =   285
      Left            =   3360
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   360
      Top             =   1080
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Go"
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   5040
      Width           =   1935
   End
   Begin VB.ListBox lstPrimes 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6780
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblTime 
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label lblTimeNeeded 
      Caption         =   "Time Needed:"
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label lblEnding 
      Caption         =   "Ending number:"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblStarting 
      Caption         =   "Starting number:"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblLoaded 
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblList 
      Caption         =   "Loaded list:"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmPrimenumbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x, Oner, One As Integer
Dim number, checksum, primes(1000000), NofPrimes, endnumber, seconds As Double
Private Sub cmdLoadList_Click()
CommonDialog1.Filter "PRN-Files (PRimeNumber-File) |*.prn"
CommonDialog1.ShowOpen
CommonDialog1.Path
CommonDialog.FileName
End Sub
Private Sub cmdStart_Click()
If chkClean.Value = 1 Then lstPrimes.Clear

endnumber = Val(txtEnd.Text)
number = 1
NofPrimes = 0 '****
seconds = 0
checksum = 0

Timer.Enabled = True
Do
 number = number + 1
 '1. Trick: If number has ending 0,2,4,5,6,8 it isn't a prime number
 If number <= 5 Then GoTo prime_check
 Oner = Mid(number, Len(number), 1)
 If Oner = 0 Or Oner = 2 Or Oner = 4 Or Oner = 5 Or Oner = 6 Or Oner = 8 Then GoTo nextnumber
 
 '2. Trick: If checksum of number can be divided by 3, it isn't a prime number
 For x = 1 To Len(number) Step 1
 One = Mid(number, Len(x), 1)
 checksum = checksum + One
 Next
 If checksum Mod 3 = 0 Then GoTo nextnumber

'Try if the number has got a rest by dividing it with prime number to it's root
prime_check:
For x = 1 To NofPrimes Step 1
If number Mod primes(x) = 0 Then GoTo nextnumber
If primes(x) * primes(x) >= number Then GoTo testing_finished
Next

testing_finished:
NofPrimes = NofPrimes + 1
primes(NofPrimes) = number

nextnumber:
If number = endnumber Then
Timer.Enabled = False
lblTime.Caption = "Time needed: " & seconds & " s"
Open "E:\Proggen\primes.prn" For Output As #1
For x = 1 To NofPrimes
Print #1, primes(x)
Next
Close #1

Exit Sub
End If
Loop

End Sub
Private Sub optLoad_Click()
txtStart.Enabled = True
End Sub
Private Sub optNew_Click()
txtStart.Enabled = False
End Sub
Private Sub Timer_Timer()
seconds = seconds + 0.01
If seconds Mod 1 = 0 Then lblTime.Caption = "Time needed: " & seconds & " s"
End Sub
