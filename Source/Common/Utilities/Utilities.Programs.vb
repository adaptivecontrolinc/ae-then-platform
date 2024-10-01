Namespace Utilities
  Public Module Programs

    Public Function CheckFirstUseOfCommand(ByVal commandCodeToCheck As String, ByVal currentStep As Integer, ByVal programs As String) As Boolean
      'Check to see if this is the first use of this command in the running program
      Try
        'These separators are used by programs and prefixed steps
        Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
        Dim separator1 As String = Convert.ToChar(255)
        Dim separator2 As String = ","

        'Split the program steps string into individual steps
        Dim programSteps() As String = programs.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

        'Loop through the steps and see what is the step number of the first turbo fill / fill 
        Dim CommandString As String
        Dim CommandCode As String
        For i As Integer = 0 To programSteps.GetUpperBound(0)
          CommandString = programSteps(i)
          If CommandString.Length >= 2 Then
            CommandCode = CommandString.Substring(0, 2).ToUpper
            If (CommandCode = commandCodeToCheck.ToUpper) Then
              'We've found the first use of this command in the program - now we have to check the step number
              If currentStep <= (i + 1) Then Return True
              'Current step is not the first use of this command
              Return False
            End If
          End If
        Next i
      Catch ex As Exception
        'Ignore errors
      End Try
      Return False
    End Function

    Public Function GetProgramSteps(ByVal program As String) As String()
      Try
        'These separators are used by programs and prefixed steps
        Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
        Dim separator1 As String = Convert.ToChar(255)
        Dim separator2 As String = ","

        Return program.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
      Catch ex As Exception
        'Ignore errors
      End Try
      Return Nothing
    End Function

    Public Function GetStepParameters(ByVal programStep As String) As String()
      Try
        Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
        Dim separator1 As String = Convert.ToChar(255)
        Dim separator2 As String = ","
        Dim helpText As String = "'"

        'Split off the help text (if any)
        Dim helpSplit() As String = programStep.Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)

        'Make sure we have something we can use
        If helpSplit Is Nothing OrElse helpSplit.GetUpperBound(0) < 0 Then Return Nothing

        'Split out command parameters
        Dim stepParameters() As String = helpSplit(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries)

        'Return what we have
        Return stepParameters
      Catch ex As Exception
        'Ignore errors
      End Try
      Return Nothing
    End Function

    Public Function GetBatchParameters(ByVal parameters As String) As String()
      Try
        'These separators are used by programs and prefixed steps
        Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
        Dim separator1 As String = Convert.ToChar(255)
        Dim separator2 As String = ","

        Return parameters.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
      Catch ex As Exception
        'Ignore errors
      End Try
      Return Nothing
    End Function

    Public Function GetCallOffNumber(ByVal currentStep As Integer, ByVal programs As String) As Integer
      'Check to see what calloff is active based on add prepare commands
      Try

        'Default value for Calloff
        Dim CallOff As Integer = 0

        'These separators are used by programs and prefixed steps
        Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
        Dim separator1 As String = Convert.ToChar(255)
        Dim separator2 As String = ","

        'Split the program steps string into individual steps
        Dim programSteps() As String = programs.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

        'Loop through each step starting from the beginning of the program and count the number of adds
        Dim CommandString As String
        Dim CommandCode As String
        For i As Integer = 0 To programSteps.GetUpperBound(0)
          CommandString = programSteps(i)
          If CommandString.Length >= 2 Then
            CommandCode = CommandString.Substring(0, 2).ToUpper
            If (CommandCode = "AP") OrElse (CommandCode = "KP") OrElse (CommandCode = "DS") OrElse (CommandCode = "KF") Then
              'We've found a matching command in the program
              CallOff += 1
              'Is the the current step
              If currentStep <= (i + 1) Then Exit For
            End If
          End If
        Next i

        'Return the Calloff Calculated
        Return CallOff

      Catch ex As Exception
        'Ignore errors
      End Try
      'Failsafe Return Value
      Return 0
    End Function

    Public Function GetCallOffNumber(ByVal Controlcode As ControlCode) As Integer
      'Check to see what calloff is active based on add prepare commands
      Try

        'Default value for Calloff
        Dim CallOff As Integer = 0

        'These separators are used by programs and prefixed steps
        Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
        Dim separator1 As String = Convert.ToChar(255)
        Dim separator2 As String = ","

        Dim QQfound As Boolean
        Dim Command As String

        With Controlcode
          'Get current program and step number
          Dim ProgNum = .Parent.ProgramNumberAsString, StepNum = .Parent.StepNumber

          'Get all program steps for this program
          Dim ProgramSteps = .Parent.PrefixedSteps.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
          'Loop through each step starting from the beginning of the program and count the number of adds
          Dim StartChecking = False

          For i = ProgramSteps.GetLowerBound(0) To ProgramSteps.GetUpperBound(0)
            'Ignore step 0
            If i > 0 Then
              Dim [Step] = ProgramSteps(i).Split(separator1.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

              If [Step].GetUpperBound(0) >= 1 Then
                If [Step](0) = ProgNum AndAlso (CDbl([Step](1)) = StepNum - 1) Then StartChecking = True
                Command = [Step](5).Substring(0, 2)

                If StartChecking Then
                  If ((Command = "KP") OrElse (Command = "PT") OrElse (Command = "AP")) AndAlso (Not QQfound) Then
                    CallOff = CallOff + 1
                    Exit For
                  Else
                    CallOff = 0
                    Exit For
                  End If
                Else
                  If (Command = "QQ") Then
                    QQfound = True
                  End If
                  If (Command = "EI") Then
                    QQfound = False
                  End If
                  If ((Command = "KP") OrElse (Command = "PT") OrElse (Command = "AP")) AndAlso (Not QQfound) Then CallOff = CallOff + 1

                End If
              End If
            End If
            If i = ProgramSteps.GetUpperBound(0) Then CallOff = 0
          Next i
        End With

        Return CallOff

      Catch ex As Exception
        'Ignore errors
      End Try
      'Failsafe Return Value
      Return 0
    End Function

  End Module
End Namespace
