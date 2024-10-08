VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MathTest"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7080
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI Semibold"
      Size            =   9.75
      Charset         =   0
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDenominator 
      Caption         =   "Denominator"
      Height          =   825
      Left            =   2880
      TabIndex        =   11
      Top             =   1500
      Width           =   3975
      Begin VB.OptionButton optRandomDenom 
         Caption         =   "Random"
         Height          =   255
         Left            =   2490
         TabIndex        =   14
         Top             =   195
         Value           =   -1  'True
         Width           =   1425
      End
      Begin VB.TextBox txtValuesDenom 
         Height          =   375
         Left            =   1005
         TabIndex        =   13
         Text            =   "1-10"
         Top             =   315
         Width           =   1260
      End
      Begin VB.OptionButton optOrderDenom 
         Caption         =   "In Order"
         Height          =   255
         Left            =   2490
         TabIndex        =   12
         Top             =   480
         Width           =   1425
      End
      Begin VB.Label lblValueDenom 
         AutoSize        =   -1  'True
         Caption         =   "Values"
         Height          =   255
         Left            =   330
         TabIndex        =   15
         Top             =   360
         Width           =   585
      End
   End
   Begin VB.Frame fraNumerator 
      Caption         =   "Numerator"
      Height          =   825
      Left            =   2880
      TabIndex        =   6
      Top             =   360
      Width           =   3975
      Begin VB.OptionButton optOrderNum 
         Caption         =   "In Order"
         Height          =   255
         Left            =   2415
         TabIndex        =   9
         Top             =   480
         Width           =   1425
      End
      Begin VB.TextBox txtValuesNum 
         Height          =   375
         Left            =   990
         TabIndex        =   8
         Text            =   "2,5"
         Top             =   315
         Width           =   1260
      End
      Begin VB.OptionButton optRandomNum 
         Caption         =   "Random"
         Height          =   255
         Left            =   2415
         TabIndex        =   7
         Top             =   195
         Value           =   -1  'True
         Width           =   1425
      End
      Begin VB.Label lblValueNum 
         AutoSize        =   -1  'True
         Caption         =   "Values"
         Height          =   255
         Left            =   315
         TabIndex        =   10
         Top             =   360
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      Height          =   660
      Left            =   4065
      TabIndex        =   2
      Top             =   4575
      Width           =   1125
   End
   Begin VB.CommandButton cmdBegin 
      Caption         =   "Begin"
      Height          =   660
      Left            =   2880
      TabIndex        =   1
      Top             =   4575
      Width           =   1125
   End
   Begin VB.ListBox lstTests 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   375
      TabIndex        =   0
      Top             =   360
      Width           =   2280
   End
   Begin VB.Label lblTodaysCorrect 
      AutoSize        =   -1  'True
      Caption         =   "Correct Answers: "
      Height          =   255
      Left            =   3105
      TabIndex        =   5
      Top             =   3270
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label lblAverageTime 
      AutoSize        =   -1  'True
      Caption         =   "Average Time Per Question: "
      Height          =   255
      Left            =   3105
      TabIndex        =   4
      Top             =   3630
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Label lblTodaysCount 
      AutoSize        =   -1  'True
      Caption         =   "Today's Count: "
      Height          =   255
      Left            =   3105
      TabIndex        =   3
      Top             =   2910
      Visible         =   0   'False
      Width           =   1410
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const COUNT_CAPTION = "Today's Count: "
Const CORRECT_CAPTION = "Correct Answers: "
Const AVERAGE_CAPTION = "Average Time Per Question: "

Private Sub cmdBegin_Click()
    Dim Operand As String
    Dim StudyMode As Boolean
    Dim TestName As String
    
    TestName = lstTests.List(lstTests.ListIndex)
    StudyMode = (Left(TestName, 5) = "Study")
    If InStr(TestName, "Addition") > 0 Then
        Operand = "+"
    ElseIf InStr(TestName, "Subtraction") > 0 Then
        Operand = "-"
    ElseIf InStr(TestName, "Division") > 0 Then
        Operand = "/"
    ElseIf InStr(TestName, "Multiplication") > 0 Then
        Operand = "x"
    End If
    frmSolve.StartTest StudyMode, txtValuesNum, txtValuesDenom, optRandomNum, optRandomDenom, Operand
    SaveSetting App.Title, "Settings", TestName & "-TestDate", Date
    SaveSetting App.Title, "Settings", TestName & "-TodaysCorrect", TodaysCorrect
    SaveSetting App.Title, "Settings", TestName & "-TodaysCount", TodaysCount
    SaveSetting App.Title, "Settings", TestName & "-TodaysTotalTime", TodaysTotalTime
    SetScores
End Sub

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    PositionForm
    ControlsLoad
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer

    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "TestDate", Date
    End If
End Sub

Private Sub ControlsLoad()
    lstTests.AddItem "Study Addition"
    lstTests.AddItem "Study Subtraction"
    lstTests.AddItem "Study Multiplication"
    lstTests.AddItem "Study Division"
'    lstTests.AddItem "Study Fractions"
'    lstTests.AddItem "Study Percent"
    
    lstTests.AddItem "Test Addition"
    lstTests.AddItem "Test Subtraction"
    lstTests.AddItem "Test Multiplication"
    lstTests.AddItem "Test Division"
'    lstTests.AddItem "Test Fractions"
'    lstTests.AddItem "Test Percent"

    lstTests.ListIndex = 0
End Sub

Private Sub PositionForm()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
End Sub

Private Sub lstTests_Click()
    Dim TestName As String
    
    TestName = lstTests.List(lstTests.ListIndex)
    TestDate = GetSetting(App.Title, "Settings", TestName & "-TestDate", 0)
    If TestDate = Date Then
        TodaysCorrect = GetSetting(App.Title, "Settings", TestName & "-TodaysCorrect", 0)
        TodaysCount = GetSetting(App.Title, "Settings", TestName & "-TodaysCount", 0)
        TodaysTotalTime = GetSetting(App.Title, "Settings", TestName & "-TodaysTotalTime", 0)
    Else
        SaveSetting App.Title, "Settings", TestName & "-TestDate", Date
        SaveSetting App.Title, "Settings", TestName & "-TodaysCorrect", 0
        SaveSetting App.Title, "Settings", TestName & "-TodaysCount", 0
        SaveSetting App.Title, "Settings", TestName & "-TodaysTotalTime", 0
        TodaysCorrect = 0
        TodaysCount = 0
        TodaysTotalTime = 0
    End If
    SetScores
End Sub

Private Sub SetScores()
    Dim StudyMode As Boolean
    Dim TestName As String
    
    TestName = lstTests.List(lstTests.ListIndex)
    StudyMode = (Left(TestName, 5) = "Study")
    If StudyMode Then
        lblTodaysCorrect = CORRECT_CAPTION & "N/A"
    Else
        lblTodaysCorrect = CORRECT_CAPTION & TodaysCorrect
    End If
    lblTodaysCount = COUNT_CAPTION & TodaysCount
    If TodaysCount = 0 Then
        lblAverageTime = AVERAGE_CAPTION & "0"
    Else
        lblAverageTime = AVERAGE_CAPTION & Round(TodaysTotalTime / TodaysCount, 2)
    End If
End Sub

Private Sub txtValuesNum_Change()
    txtValuesNum = Trim(txtValuesNum)
End Sub

Private Sub txtValuesDenom_Change()
    txtValuesDenom = Trim(txtValuesDenom)
End Sub

