VERSION 5.00
Begin VB.Form frmSolve 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solve The Problem"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTick 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   3030
      TabIndex        =   1
      Top             =   2985
      Width           =   885
   End
   Begin VB.TextBox txtGuess 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   720
      TabIndex        =   0
      Top             =   2820
      Width           =   2235
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Stop Test"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   3090
      TabIndex        =   2
      Top             =   180
      Width           =   885
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show Answer"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   720
      TabIndex        =   3
      Top             =   2820
      Width           =   2235
   End
   Begin VB.Label lblAnswer 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "81"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   1770
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Label lblDenominator 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "x 9"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   1515
      TabIndex        =   6
      Top             =   1230
      Width           =   1365
   End
   Begin VB.Label lblEqualsBar 
      Caption         =   "_______"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   705
      TabIndex        =   5
      Top             =   1410
      Width           =   2235
   End
   Begin VB.Label lblNumerator 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   2325
      TabIndex        =   4
      Top             =   150
      Width           =   555
   End
End
Attribute VB_Name = "frmSolve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mOperand As String
Dim mNumerators() As String
Dim mDenominators() As String
Dim mStudyMode As Boolean
Dim NextDenominator As Single
Dim NextNumerator As Single
Dim NextDenominatorIndex As Integer
Dim NextNumeratorIndex As Integer

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    TodaysCount = TodaysCount + 1
    If Not mStudyMode Then
        tmrTick.Enabled = False
        If IsNumeric(txtGuess) Then
            If txtGuess = lblAnswer Then
                MsgBox "Correct!"
                TodaysCorrect = TodaysCorrect + 1
            Else
                MsgBox "Incorrect. The correct answer is " & lblAnswer & "."
            End If
        Else
            MsgBox "Incorrect. The correct answer is " & lblAnswer & "."
        End If
        txtGuess.SetFocus
    End If
    LoadNextQuestion
End Sub

Private Sub cmdShow_Click()
    cmdShow.Visible = False
    tmrTick.Enabled = False
End Sub

Public Sub StartTest(ByVal StudyMode As Boolean, ByVal Numerators As String, ByVal Denominators As String, ByVal RandomNum As Boolean, ByVal RandomDenom As Boolean, ByVal Operand As String)
    Randomize
    cmdShow.Visible = StudyMode
    lblAnswer.Visible = StudyMode
    txtGuess.Visible = Not StudyMode
    
    mStudyMode = StudyMode
    Numerators = ParseValues(Numerators)
    Denominators = ParseValues(Denominators)
    mNumerators = Split(Numerators, ",")
    mDenominators = Split(Denominators, ",")
    mOperand = Operand
    If RandomNum Then
        NextNumeratorIndex = -1
    Else
        NextNumeratorIndex = 0
    End If
    If RandomDenom Then
        NextDenominatorIndex = -1
    Else
        NextDenominatorIndex = 0
    End If
    LoadNextQuestion
    
    Me.Show vbModal
End Sub

Private Sub LoadNextQuestion()
    Dim Index As Integer
    
    tmrTick.Enabled = True
    If mStudyMode Then
        cmdShow.Visible = True
        If cmdShow.Visible Then
            cmdShow.SetFocus
        End If
    End If
    txtGuess = ""
    NextDenominator = GetNextNumber(NextDenominatorIndex, mDenominators())
    NextNumerator = GetNextNumber(NextNumeratorIndex, mNumerators())
    lblNumerator = NextNumerator
    If mOperand = "/" Then
        If NextDenominator = 0 Then
            NextDenominator = 1
        End If
    End If
    lblDenominator = mOperand & " " & NextDenominator
    Select Case mOperand
        Case "+": lblAnswer = NextNumerator + NextDenominator
        Case "-": lblAnswer = NextNumerator - NextDenominator
        Case "/": lblAnswer = Round(NextNumerator / NextDenominator, 2)
        Case "x": lblAnswer = NextNumerator * NextDenominator
    End Select
End Sub

Private Sub Form_Activate()
    If mStudyMode Then
        cmdShow.SetFocus
    Else
        txtGuess.SetFocus
    End If
End Sub

Private Sub tmrTick_Timer()
    TodaysTotalTime = TodaysTotalTime + 1
End Sub

