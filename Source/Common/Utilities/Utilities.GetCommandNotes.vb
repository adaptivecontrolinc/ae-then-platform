Namespace Utilities.Conversions
  Public Module GetCommandNotes

    Public Function GetNotes(ByVal Controlcode As ControlCode) As String
      Try

        'These separators are used by programs and prefixed steps
        Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
        Dim separator1 As String = Convert.ToChar(255)
        Dim separator2 As String = ","

        Dim ProgNum, StepNum As Integer
        Dim Notes As String
        Notes = " " 'Return this value

        Dim ProgramSteps(), checkingfornotes() As String
        Dim i As Integer
        Dim StartChecking As Boolean

        Dim [Step]() As String
        Dim Command As String

        With Controlcode
          'Get current program and step number
          ProgNum = .Parent.ProgramNumber : StepNum = .Parent.StepNumber
          'Get all program steps for this program
          ProgramSteps = .Parent.PrefixedSteps.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
          'Loop through each step starting from the beginning of the program and count the number of adds
          StartChecking = False

          For i = ProgramSteps.GetLowerBound(0) To ProgramSteps.GetUpperBound(0)
            'Ignore step 0
            If i > 0 Then
              [Step] = ProgramSteps(i).Split(separator1.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

              If [Step].GetUpperBound(0) >= 1 Then
                If (CDbl([Step](0)) = ProgNum) AndAlso (CDbl([Step](1)) = StepNum - 1) Then StartChecking = True
                Command = [Step](5).Substring(0, 2)

                If StartChecking Then
                  checkingfornotes = [Step](5).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
                  Dim t As Integer = checkingfornotes.GetUpperBound(0)
                  If t >= 1 Then
                    Notes = checkingfornotes(1)
                  Else
                    Notes = ""
                  End If
                  Exit For
                End If
              End If
            End If


            If i = ProgramSteps.GetUpperBound(0) Then Notes = ""
          Next i
        End With

        Return Notes

      Catch ex As Exception

      End Try

      Return ""

    End Function

  End Module
End Namespace
