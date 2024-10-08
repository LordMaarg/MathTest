Attribute VB_Name = "basVariables"
Option Explicit

Public TestDate As Date
Public TodaysCorrect As Long
Public TodaysCount As Long
Public TodaysTotalTime As Long

Public Function ParseValues(ByVal Values As String) As String
    Dim AllValues As String
    Dim TestValues() As String
    Dim V As Integer
    Dim TV As Integer
    Dim TestValue As String
    Dim LV As Long
    Dim HV As Long
    
    TestValues = Split(Values, ",")
    For TV = 0 To UBound(TestValues)
        If InStr(TestValues(TV), "-") > 0 Then
            On Error Resume Next
            LV = Val(Split(TestValues(TV), "-")(0))
            HV = Val(Split(TestValues(TV), "-")(1))
            If LV > 999 Then LV = 999
            If HV > 999 Then HV = 999
            For V = LV To HV
                AllValues = AllValues & V & ","
            Next 'V
        Else
            AllValues = AllValues & TestValues(TV) & ","
        End If
    Next 'TV
    
    ParseValues = AllValues
End Function

Public Function GetNextNumber(NextIndex As Integer, Values() As String) As Integer
    Dim Index As Integer
    
    If NextIndex = -1 Then
        Index = Int(UBound(Values) * Rnd)
    Else
        Index = NextIndex
        NextIndex = NextIndex + 1
        If NextIndex > UBound(Values) Then
            NextIndex = 1
            Index = 0
        End If
    End If
    
    GetNextNumber = Values(Index)
End Function
