Imports System.Globalization
Imports System.Xml

Partial Public Class Utilities
  Partial Public Class GetEndTime

    Public Shared Function GetEndTime(ByVal controlCode As ControlCode) As Integer
      On Error GoTo Err_getendtime
      Dim MinutesTillEndOfCycle As Integer : MinutesTillEndOfCycle = 0 'Return this value

      With controlCode
        'Get current program and step number
        Dim ProgNum As Integer : ProgNum = .Parent.ProgramNumber
        Dim StepNum As Integer : StepNum = .Parent.StepNumber

        'Get all program steps for this program
        Dim PrefixedSteps As String : PrefixedSteps = .Parent.PrefixedSteps
        Dim lineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
        Dim ProgramSteps() As String : ProgramSteps = PrefixedSteps.Split(lineFeed.ToCharArray)
        Dim Steps() As String, Command As String

        'Loop through each step starting from the beginning of the program and count the number hot drops
        Dim i As Integer
        Dim StartChecking As Boolean : StartChecking = False
        For i = ProgramSteps.GetLowerBound(0) To ProgramSteps.GetUpperBound(0)
          Steps = ProgramSteps(i).Split(Convert.ToChar(255))
          If Steps.GetUpperBound(0) >= 1 Then
            If (Steps(0) = ProgNum.ToString) And (Steps(1) = (StepNum - 1).ToString) Then StartChecking = True
            If StartChecking Then
              MinutesTillEndOfCycle = MinutesTillEndOfCycle + CInt(Steps(3))
            End If
            If (Steps(0) = ProgNum.ToString) And (Steps(1) = (StepNum - 1).ToString) Then StartChecking = True
          End If
        Next i
      End With
      GetEndTime = MinutesTillEndOfCycle

      'Avoid Error Handler
      Exit Function

Err_getendtime:
      MinutesTillEndOfCycle = 0
    End Function


  End Class
End Class