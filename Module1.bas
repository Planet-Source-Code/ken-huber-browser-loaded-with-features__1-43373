Attribute VB_Name = "Module1"
Public Expanded As Boolean
Public UserNames As String
Public HomePage As String
Public HistorySave As Integer

Function Check_Apos(x As String) As Boolean
    Dim temp As Boolean
    temp = False
    temp = CheckStr(x, "'")
    If temp = True Then
        MsgBox ("You cannot use an apostrophy.")
        Check_Apos = True
    Else
        Check_Apos = False
    End If
End Function



Function CheckStr(StringToCheck As String, Delim As String) As Boolean
    Dim ParsedData() As String
    Dim temp As Integer
    ParsedData = Split(StringToCheck, Delim)
    temp = UBound(ParsedData)
    If temp > 0 Then
        CheckStr = True
    Else
        CheckStr = False
    End If
End Function
