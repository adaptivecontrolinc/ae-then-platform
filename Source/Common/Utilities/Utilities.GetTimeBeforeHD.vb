Imports System.Globalization
Imports System.Xml

Namespace Utilities.Conversions
  Public Module GetTimeBeforeHD

    Public Function GetTimeBeforeHDTime(ByVal controlCode As ControlCode) As Integer
      On Error GoTo Err_gettimebeforeHDtime
      Dim TimeBeforeHD As Integer : TimeBeforeHD = 0 'Return this value

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
              Command = Steps(5).Trim().Substring(0, 2)
              If Not (Command = "HD") Then TimeBeforeHD = TimeBeforeHD + CInt(Steps(3))
              If (Command = "HD") Then Exit For
            End If
          End If
          If i = ProgramSteps.GetUpperBound(0) Then TimeBeforeHD = 0
        Next i
      End With
      GetTimeBeforeHDTime = TimeBeforeHD

      'Avoid Error Handler
      Exit Function

Err_gettimebeforeHDtime:
      GetTimeBeforeHDTime = 0
    End Function


  End Module
End Namespace