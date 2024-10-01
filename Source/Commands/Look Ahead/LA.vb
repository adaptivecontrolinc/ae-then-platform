'American & Efird - MX Package
' Version 2022-05-23

' NOTE: ParallelCommand Instructs this Start Function to be called with the Command directly before it in the Program Command structure.
'       - As an example, Program Steps: TM 05, LA, TM 05 will cause the LA.Start to be called as soon as the first TM.Start is called

Imports Utilities.Translations

<Command("Look Ahead", "", "", "", "", CommandType.ParallelCommand), _
Description("Begins looking ahead to the next tank preparation command."), _
Category("Look Ahead Functions")> _
Public Class LA
  Inherits MarshalByRefObject
  Implements ACCommand

  ' Command States
  Public Enum EState
    Off
    CheckExistingDyelot
    CheckNextDyelot
    Start
    Run
    Complete
  End Enum
  Property State As EState
  Property StateString As String

  ' Look Ahead Properties
  'Public KP1 As New Tank1Prepare(Me.ControlCode)
  Public KP1 As Tank1Prepare
  Public KP1Calloff As Integer

  Public Property Dyelot As String
  Public Property ReDye As Integer
  Public Property Job As String
  Public Property Blocked As Boolean
  Public Property Batched As Boolean
  Public Property ProgramNumber As Integer

  ' KP CommandStep parameters
  '   ("Prepare", "|0-99|% |0-180|F |0-99|mins Mix:|0-1| D:|1-2|", " ", "", "")
  Public Property Tank1Enabled As Boolean
  Public Property Tank1FillLevel As Integer
  Public Property Tank1TempDesired As Integer
  Public Property Tank1MixTime As Integer
  Public Property Tank1MixRequested As Integer
  Public Property Tank1Destination As EKitchenDestination
  Public Property Tank1Calloff As Integer

  ' LookAhead properties
  Public Property LookAheadFound As Boolean
  Public Property LookAheadStopFound As Boolean
  Public Property LookAheadRepeatFound As Boolean
  Public Property LookAheatRepeat As String
  Public Property WaitReady1() As Boolean
  Public Property IPKP1() As String
  Public Property IpNumber As Integer
  Public Property ProgramHasAnotherTank1 As Boolean

  ' Timers
  Public ReadOnly Timer As New Timer
  Friend TimerCheckNextDyelot As New Timer
  Public ReadOnly Property TimeCheckNextDyelot As Integer
    Get
      Return TimerCheckNextDyelot.Seconds
    End Get
  End Property


  'Keep a local reference to the control code object for convenience
  Private ReadOnly ControlCode As ControlCode
  Sub New(ByVal controlCode As ControlCode)
    Me.ControlCode = controlCode
    Me.ControlCode.Commands.Add(Me)
    KP1 = New Tank1Prepare(Me.ControlCode)
  End Sub

  Public Sub ParametersChanged(ParamArray param() As Integer) Implements ACCommand.ParametersChanged
    'If the command is not running then do nothing the changes will be picked up when the command starts
    If Not IsOn Then Exit Sub
  End Sub

  ' NOTE: ParallelCommand Instructs this Start Function to be called with the Command directly before it in the Program Command structure.
  '       - As an example, Program Steps: TM 05, LA, TM 05 will cause the LA.Start to be called as soon as the first TM.Start is called
  Public Function Start(ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With ControlCode

      ' Reset States & Variables
      Dyelot = ""
      ReDye = 0
      Blocked = False
      ProgramNumber = 0
      IPKP1 = ""
      Tank1Calloff = 0
      Tank1Enabled = False
      ProgramHasAnotherTank1 = False

      LookAheadFound = False
      LookAheadStopFound = False
      LookAheadRepeatFound = False
      LookAheatRepeat = ""


      ' Initialize command
      State = EState.Off
      Timer.Seconds = 5

      ' LookAhead enabled - begin checking existing dyelot first
      If .Parameters.LookAheadEnabled = 1 Then
        State = EState.CheckExistingDyelot
        Tank1Calloff = .GetNextCalloff
      End If

      'This is a background command so step on
      Return True
    End With
  End Function

  Public Function Run() As Boolean Implements ACCommand.Run
    'If the command is not active then don't run this code
    If Not IsOn Then Exit Function


    With ControlCode
      ' Note State
      Static StartState As EState
      Do
        'Remember state and loop until state does not change
        'This makes sure we go through all the state changes before proceeding and setting IO
        StartState = State

        Select Case State

          Case EState.Off
            StateString = (" ")


          Case EState.CheckExistingDyelot
            StateString = Translate("Searching Existing Job")
            ' Check to see if there is another kp in this program.
            ' TODO look for LS (Look Ahead Stop)
            If CheckExistingProgram() Then
              If ProgramHasAnotherTank1 Then
                .LARequest = False
                State = EState.Start
                Timer.Seconds = 5
              Else
                .LARequest = True
                State = EState.CheckNextDyelot
              End If
            End If
            ' Stop looking if we've found a LA Repeat
            If LookAheadRepeatFound Then Cancel()


          Case EState.CheckNextDyelot
            StateString = Translate("Search Next Job") & (" ") & TimerCheckNextDyelot.ToString
            ' Check next dyelot scheduled on 60-second interval
            If TimerCheckNextDyelot.Finished Then
              ' Reset timer
              TimerCheckNextDyelot.Seconds = 60
              ' If wasBlocked and parameter for 'IgnoreBlocked' disabled, recheck
              If Blocked Then
                LookAheadRepeatFound = False
                LookAheatRepeat = ""
              End If
              ' Check next dyelot
              If CheckNextDyelot() Then
                'If the next lot is not blocked then start preparing the tanks
                If (Not Blocked) Or (.Parameters.LookAheadIgnoreBlocked = 1) Then
                  If LookAheadRepeatFound Then
                    'stop trying to check next dyelot if we've found a LA repeat
                    Cancel()
                  Else
                    .LARequest = False
                    State = EState.Start
                    KP1Calloff = Tank1Calloff
                    .KitchenJob = Job
                    .LASetSqlString = True
                    .LASqlString = "UPDATE Dyelots SET Committed=1 WHERE Dyelot='" & Dyelot & "'"
                  End If
                End If
              Else
                If LookAheadRepeatFound Then
                  If (Not Blocked) OrElse (.Parameters.LookAheadIgnoreBlocked = 1) Then
                    .LARequest = False
                    Cancel()
                  End If
                End If
              End If
            End If


          Case EState.Start
            StateString = Translate("Starting") & Timer.ToString(1)
            If Timer.Finished Then
              If Tank1Enabled Then
                KP1.Start(Tank1Calloff, Tank1Destination, Tank1FillLevel, Tank1TempDesired, Tank1MixTime, Tank1MixRequested, .Parameters.StandardTimeAddPrepare)
              End If
              State = EState.Run
            End If

          Case EState.Run
            StateString = Translate("Active") & Timer.ToString(1)
            If Tank1Enabled Then KP1.Run()
            If Timer.Finished Then
              If Not KP1.IsOn Then State = EState.Complete
              Timer.Seconds = 5
            End If

          Case EState.Complete
            StateString = Translate("Completing") & Timer.ToString(1)
            If Timer.Finished Then Cancel()

        End Select
      Loop Until (StartState = State)

      ' Cancel this command if no longer enabled
      If .Parameters.LookAheadEnabled <> 1 Then
        Cancel()
        Exit Function
      End If

    End With

  End Function

  Public Sub Cancel() Implements ACCommand.Cancel
    'Looks a bit odd but this function is called when a lot finishes
    ' - LA command must remain active from lot to lot so we only cancel when both tanks are inactive
    If Not (KP1.IsOn) Then
      KP1.Cancel()
      State = EState.Off
      StateString = ""
      Timer.Cancel()
      TimerCheckNextDyelot.Cancel()
    End If
  End Sub

  '***********************************************************************************************
  '******                  Check current Program for the next KP/LA/LS                      ******
  '******               Get values from database instead of from Parent.Job                 ******
  '***********************************************************************************************
  Public Function CheckExistingProgram() As Boolean
    Dim sql As String = ("SELECT * FROM Dyelots WHERE State=2 ORDER BY StartTime DESC")
    Dim returnValue As Boolean = False
    Try
      With ControlCode

        ' Get currently active job's Dyelot & Redye values from dbo.Dyelots
        Dim dt As System.Data.DataTable = .Parent.DbGetDataTable(sql)
        If (dt IsNot Nothing) AndAlso (dt.Rows.Count > 0) Then
          For Each dr As System.Data.DataRow In dt.Rows
            'Get updated Dyelot/Redye values from Datatable
            Dyelot = Utilities.Null.NullToEmptyString(dr("Dyelot"))
            ReDye = Utilities.Null.NullToZeroInteger(dr("ReDye"))
          Next
        End If

        ' These separators are used by programs and prefixed steps
        Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
        Dim separator1 As String = Convert.ToChar(255)
        Dim separator2 As String = ","

        ' Get current program and step number
        Dim programNumber As String = .Parent.ProgramNumberAsString
        Dim stepNumber As Integer : stepNumber = .Parent.StepNumber

        ' Get all program steps for this program
        Dim PrefixedSteps As String : PrefixedSteps = .Parent.PrefixedSteps
        Dim ProgramSteps() As String : ProgramSteps = PrefixedSteps.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

        ' Variables for Command Step parameters & Notes
        Dim stepParameters() As String
        Dim stepNotes() As String

        'Loop through each step starting from the beginning of the program
        Dim StartChecking As Boolean = False
        Dim KitchenPrepareFound As Boolean = False
        Dim CommandCode() As String
        Dim KP1 As String = ""

        ' Look for the next KP,LA, or LS command
        For i = ProgramSteps.GetLowerBound(0) To ProgramSteps.GetUpperBound(0)
          'Ignore step 0
          If i > 0 Then
            ' Current Step
            Dim [Step] = ProgramSteps(i).Split(separator1.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            If [Step].GetUpperBound(0) >= 1 Then
              If ([Step](0)) = programNumber AndAlso (CDbl([Step](1)) = stepNumber - 1) Then StartChecking = True
              CommandCode = [Step](5).Split(separator2.ToCharArray)
              If StartChecking AndAlso Not KitchenPrepareFound Then
                Select Case CommandCode(0).ToUpper

                  Case "KP"
                    ' Only find first KP after LA
                    KitchenPrepareFound = True
                    KP1 = [Step](5)

                  Case "LA"
                    ' Do nothing if another LA is found before the next KP
                    If LookAheadRepeatFound Then
                      LookAheadRepeatFound = True
                      LookAheatRepeat = "Program: " & programNumber.Trim & " Step: " & stepNumber.ToString("#0")
                    End If
                    LookAheadFound = True

                  Case "LS"
                    LookAheadStopFound = True
                    Exit For

                End Select
              End If
            End If
          End If
          ' Reached the upper bound of the program steps array
          If i = ProgramSteps.GetUpperBound(0) Then returnValue = False
        Next i

        ' Do nothing if second LA is found before next KP
        If LookAheadRepeatFound Then
          Return False
          Exit Function
        End If

        ' Set next TankPrepare parameters
        Tank1Enabled = False
        ProgramHasAnotherTank1 = False
        Tank1FillLevel = 0
        Tank1TempDesired = 0
        Tank1MixTime = 0
        Tank1MixRequested = 0
        Tank1Destination = EKitchenDestination.Drain

        ' KP1 was set with something, now break it apart
        If KP1.Length > 0 Then
          ProgramHasAnotherTank1 = True
          Tank1Enabled = True

          ' Get step notes, if any
          stepNotes = KP1.Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
          stepParameters = stepNotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries)

          ' 5 parameters defined for KP command: ["|0-99|% |0-180|F |0-99|mins Mix:|0-1| D:|1-2|"]
          If stepParameters.GetUpperBound(0) >= 5 Then
            '0 = command '1 = Level, 2 = desired temp, 3 = mixing time, 4 = mixer on or off, 5 = Destination

            ' Update local variables with command step parameters
            Tank1FillLevel = CInt(stepParameters(1))
            Tank1TempDesired = CInt(stepParameters(2))
            Tank1MixTime = CInt(stepParameters(3))
            Tank1MixRequested = CInt(stepParameters(4))
            Tank1Destination = CType((CInt(stepParameters(5))), EKitchenDestination)

          End If

        End If

        ' Everything ok
        returnValue = True

      End With
    Catch ex As Exception
      'Clear if there's a problem

      Utilities.Log.LogError(ex.Message, sql)
    End Try

    'Return Default Value 
    Return returnValue
  End Function

  '***********************************************************************************************
  '******                          Check next scheduled dyelot                              ******
  '******                            Get values from database                               ******
  '***********************************************************************************************
  Public Function CheckNextDyelot() As Boolean
    Dim sql As String = ("SELECT TOP(1) * FROM Dyelots WHERE State IS NULL ORDER BY StartTime DESC")
    Dim returnValue As Boolean = False
    Try
      With ControlCode

        ' Get next scheduled job's Dyelot & Redye values from dbo.Dyelots (Confirm Only 1 Row)
        Dim dt As System.Data.DataTable = .Parent.DbGetDataTable(sql)
        If (dt IsNot Nothing) AndAlso (dt.Rows.Count = 1) Then
          For Each dr As System.Data.DataRow In dt.Rows
            'Get updated Dyelot/Redye values from Datatable
            Dyelot = Utilities.Null.NullToEmptyString(dr("Dyelot"))
            ReDye = Utilities.Null.NullToZeroInteger(dr("ReDye"))
            Blocked = Not dr.IsNull("Blocked")
            Batched = Not dr.IsNull("Batched")
            ProgramNumber = Utilities.Null.NullToZeroInteger(dr("Program"))

            ' The first batch ready to run on this machine (if any) must be not blocked and must have the Batched time set to be considered
            If Blocked Or Not Batched Then Return False

            ' Get Program Steps
            sql = "SELECT Steps FROM Programs WHERE ProgramNumber='" & ProgramNumber & "'"
            Dim ProgramSteps As String
            Dim dtProgram As System.Data.DataTable = .Parent.DbGetDataTable(sql)
            If dtProgram Is Nothing OrElse dtProgram.Rows.Count <> 1 Then Return False

            ' Now look through the program steps
            Dim drProgram As System.Data.DataRow = dtProgram.Rows(0)
            ProgramSteps = Utilities.Null.NullToEmptyString(drProgram("Steps"))

            returnValue = CheckNextProgram(ProgramSteps)
          Next
        End If

      End With

      ' Everything ok
      returnValue = returnValue

    Catch ex As Exception
      'Clear if there's a problem

      Utilities.Log.LogError(ex.Message, sql)
    End Try

    'Return Default Value 
    Return returnValue
  End Function

  '***********************************************************************************************
  '******                   Check next scheduled program for IP/KP/LA/LS                    ******
  '***********************************************************************************************
  Private Function CheckNextProgram(ProgramSteps As String) As Boolean
    Dim returnValue As Boolean = False
    Try
      With ControlCode

        'Make sure we've got something to check
        If ProgramSteps.Length <= 0 Then Return False

        'These separators are used by programs and prefixed steps
        Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
        Dim separator1 As String = Convert.ToChar(255)
        Dim separator2 As String = ","

        ' Variables for Command Step parameters & Notes
        Dim stepParameters() As String
        Dim stepNotes() As String

        'Loop through each step starting from the beginning of the program
        Dim StartChecking As Boolean = False
        Dim KitchenPrepareFound As Boolean = False
        Dim CommandCode() As String
        Dim CommandPrompt As Integer
        Dim KP1 As String = ""
        Dim IPFound As Integer

        ' Split the program string into an array of program steps - one step per array element
        Dim [Steps] As String() = ProgramSteps.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

        ' Make sure we've have something to check
        If Steps.GetUpperBound(0) <= 0 Then Return False

        ' Look for the next KP,LA, or LS command
        For i As Integer = 0 To Steps.GetUpperBound(0)
          If LookAheadStopFound = True Then Exit For
          If Tank1Enabled = True Then Exit For
          ' Current step command code
          CommandCode = Steps(i).Split(separator2.ToCharArray)

          ' look for the commands we want
          Select Case CommandCode(0)

            Case "IP"
              ' Get the program number
              stepNotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
              stepParameters = stepNotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
              ' Confirm necessary parameters are set
              If stepParameters.GetUpperBound(0) >= 2 Then
                IPFound = CInt(stepParameters(2)) * 256 + CInt(stepParameters(1))
              Else
                IPFound = CInt(stepParameters(1))
              End If
              ' Check for matching IP program
              CheckIPProgram(IPFound)

            Case "KP"
              ' get the program number
              stepNotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
              stepParameters = stepNotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
              ' Confirm necessary parameters are set
              If stepParameters.GetUpperBound(0) >= 1 Then
                CommandPrompt = CInt(stepParameters(1))
              End If
              If CommandPrompt <> 99 Then
                ' If its not 99 then change it.
                Tank1Calloff = Tank1Calloff + 1
              End If
              Tank1Enabled = True
              Exit For 'we found it we are done.

            Case "LA"
              ' Do nothing if another LA is found before the next KP
              LookAheadRepeatFound = True
              LookAheatRepeat = "Program: " ' & ProgramNumber.Trim & " Step: " & stepNumber.ToString("#0")
              Exit For

            Case "LS"
              LookAheadStopFound = True
              Exit For

          End Select
        Next i
      End With

      ' Everything ok
      Return returnValue = True

    Catch ex As Exception
      Utilities.Log.LogError(ex.Message)
    End Try

    'Return Default Value 
    Return False
  End Function

  '***********************************************************************************************
  '******                   1st IP found within next program - Get Steps                    ******
  '***********************************************************************************************
  Private Function CheckIPProgram(ProgramNumber As Integer) As Boolean
    Dim sql As String = "SELECT * FROM Programs WHERE ProgramNumber=" & ProgramNumber
    Dim returnValue As Boolean = False
    Try
      With ControlCode

        Dim dt As System.Data.DataTable = .Parent.DbGetDataTable(sql)
        If dt Is Nothing OrElse dt.Rows.Count <> 1 Then Return False

        Dim dr As System.Data.DataRow = dt.Rows(0)
        Dim ProgramSteps As String = Utilities.Null.NullToEmptyString(dr("Steps"))

        returnValue = CheckIPProgramSteps(ProgramSteps)

      End With

      ' Everything ok
      Return returnValue

    Catch ex As Exception
      Utilities.Log.LogError(ex.Message)
    End Try

    ' Return Default Value
    Return returnValue
  End Function

  '***********************************************************************************************
  '******                         1st IP - check for IP/KP/LA/LS                            ******
  '***********************************************************************************************
  Private Function CheckIPProgramSteps(ProgramSteps As String) As Boolean
    Dim returnValue As Boolean = False
    Try

      'Make sure we've got something to check
      If ProgramSteps.Length <= 0 Then Return False

      ' These separators are used by programs and prefixed steps
      Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
      Dim separator1 As String = Convert.ToChar(255)
      Dim separator2 As String = ","

      ' Use this to split out steps in this program
      Dim Steps() As String
      Dim stepParameters() As String
      Dim stepNotes() As String

      ' Loop through the program steps to find the first KP command for each tank
      Dim StartChecking As Boolean = False
      Dim KitchenPrepareFound As Boolean = False
      Dim CommandCode() As String
      Dim CommandPrompt As Integer
      Dim KP1 As String = ""
      Dim IPFound As Integer

      ' Split the program string into an array of program steps - one step per array element
      Steps = ProgramSteps.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

      ' Make sure we've have something to check
      If Steps.GetUpperBound(0) <= 0 Then Return False

      ' Look for the first KP command
      For i As Integer = 0 To Steps.GetUpperBound(0)
        If LookAheadStopFound = True Then Exit For
        If Tank1Enabled = True Then Exit For
        ' Current step command code
        CommandCode = Steps(i).Split(separator2.ToCharArray)

        ' look for the commands we want
        Select Case CommandCode(0)

          Case "IP"
            ' get the program number
            stepNotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            stepParameters = stepNotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            ' Confirm necessary parameters are set
            If stepParameters.GetUpperBound(0) >= 2 Then
              IPFound = CInt(stepParameters(2)) * 256 + CInt(stepParameters(1))
            Else
              IPFound = CInt(stepParameters(1))
            End If
            CheckIP2Program(IPFound)


          Case "KP"
            ' get the program number
            stepNotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            stepParameters = stepNotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            ' Confirm necessary parameters are set
            If stepParameters.GetUpperBound(0) >= 1 Then
              CommandPrompt = CInt(stepParameters(1))
            End If
            If CommandPrompt <> 99 Then
              ' If its not 99 then change it.
              Tank1Calloff = Tank1Calloff + 1
            End If
            Tank1Enabled = True
            Exit For 'we found it we are done.


          Case "LA"
            ' Do nothing if another LA is found before the next KP
            LookAheadRepeatFound = True
            LookAheatRepeat = "Program: " ' & ProgramNumber.Trim & " Step: " & stepNumber.ToString("#0")
            Exit For


          Case "LS"
            LookAheadStopFound = True
            Exit For

        End Select
      Next i

      ' Everything ok
      Return returnValue = True

    Catch ex As Exception
      Utilities.Log.LogError(ex.Message)
    End Try

    ' Return Default Value
    Return returnValue
  End Function

  '  Remainder is designed to Check Sub IP's within IP -
  '    IP contains IP which contains IP which contains IP
  '***********************************************************************************************
  '******              IP found within 1st IP program steps - Get it's steps                ******
  '***********************************************************************************************
  'this is for a IP inside the first IP
  Private Function CheckIP2Program(ProgramNumber As Long) As Boolean
  Dim sql As String = "SELECT * FROM Programs WHERE ProgramNumber=" & ProgramNumber
    Dim returnValue As Boolean = False
    Try
      With ControlCode

        Dim dt As System.Data.DataTable = .Parent.DbGetDataTable(sql)
        If dt Is Nothing OrElse dt.Rows.Count <> 1 Then Return False

        Dim dr As System.Data.DataRow = dt.Rows(0)
        Dim ProgramSteps As String = Utilities.Null.NullToEmptyString(dr("Steps"))

        returnValue = CheckIP2ProgramSteps(ProgramSteps)

      End With

      ' Everything ok
      Return returnValue

    Catch ex As Exception
      Utilities.Log.LogError(ex, "CheckIP2Program")
    End Try

    ' Return Default Value
    Return returnValue
  End Function

  '***********************************************************************************************
  '******                      Check IP within IP's Steps - KP/LA/LS                        ******
  '***********************************************************************************************
  Private Function CheckIP2ProgramSteps(ProgramSteps As String) As Boolean
    Dim returnValue As Boolean = False
    Try

      'Make sure we've got something to check
      If ProgramSteps.Length <= 0 Then Return False

      ' These separators are used by programs and prefixed steps
      Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
      Dim separator1 As String = Convert.ToChar(255)
      Dim separator2 As String = ","

      ' Use this to split out steps in this program
      Dim Steps() As String
      Dim stepParameters() As String
      Dim stepNotes() As String

      ' Loop through the program steps to find the first KP command for each tank
      Dim StartChecking As Boolean = False
      Dim KitchenPrepareFound As Boolean = False
      Dim CommandCode() As String
      Dim CommandPrompt As Integer
      Dim KP1 As String = ""
      Dim IPFound As Integer

      ' Split the program string into an array of program steps - one step per array element
      Steps = ProgramSteps.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

      ' Make sure we've have something to check
      If Steps.GetUpperBound(0) <= 0 Then Return False

      ' Look for the first KP command
      For i As Integer = 0 To Steps.GetUpperBound(0)
        If LookAheadStopFound = True Then Exit For
        If Tank1Enabled = True Then Exit For
        ' Current step command code
        CommandCode = Steps(i).Split(separator2.ToCharArray)

        ' look for the commands we want
        Select Case CommandCode(0)

          Case "IP"
            ' get the program number
            stepNotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            stepParameters = stepNotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            ' Confirm necessary parameters are set
            If stepParameters.GetUpperBound(0) >= 2 Then
              IPFound = CInt(stepParameters(2)) * 256 + CInt(stepParameters(1))
            Else
              IPFound = CInt(stepParameters(1))
            End If
            CheckIP3Program(IPFound)


          Case "KP"
            ' get the program number
            stepNotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            stepParameters = stepNotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            ' Confirm necessary parameters are set
            If stepParameters.GetUpperBound(0) >= 1 Then
              CommandPrompt = CInt(stepParameters(1))
            End If
            If CommandPrompt <> 99 Then
              ' If its not 99 then change it.
              Tank1Calloff = Tank1Calloff + 1
            End If
            Tank1Enabled = True
            Exit For 'we found it we are done.


          Case "LA"
            ' Do nothing if another LA is found before the next KP
            LookAheadRepeatFound = True
            LookAheatRepeat = "Program: " ' & ProgramNumber.Trim & " Step: " & stepNumber.ToString("#0")
            Exit For


          Case "LS"
            LookAheadStopFound = True
            Exit For

        End Select
      Next i

      ' Everything ok
      Return returnValue = True

    Catch ex As Exception
      Utilities.Log.LogError(ex, "CheckIP2ProgramSteps")
    End Try

    ' Return Default Value
    Return returnValue
  End Function

  '***********************************************************************************************
  '******                           IP found within Sub IP - Get Steps                      ******
  '***********************************************************************************************
  Private Function CheckIP3Program(ProgramNumber As Long) As Boolean
    Dim sql As String = "SELECT * FROM Programs WHERE ProgramNumber=" & ProgramNumber
    Dim returnValue As Boolean = False
    Try
      With ControlCode

        Dim dt As System.Data.DataTable = .Parent.DbGetDataTable(sql)
        If dt Is Nothing OrElse dt.Rows.Count <> 1 Then Return False

        Dim dr As System.Data.DataRow = dt.Rows(0)
        Dim ProgramSteps As String = Utilities.Null.NullToEmptyString(dr("Steps"))

        returnValue = CheckIP3ProgramSteps(ProgramSteps)

      End With

      ' Everything ok
      Return returnValue

    Catch ex As Exception
      Utilities.Log.LogError(ex.Message)
    End Try

    ' Return Default Value
    Return returnValue
  End Function

  '***********************************************************************************************
  '******                          Check Sub IP's steps for IP/LA/LS                        ******
  '***********************************************************************************************
  Private Function CheckIP3ProgramSteps(ProgramSteps As String) As Boolean
    Dim returnValue As Boolean = False
    Try

      'Make sure we've got something to check
      If ProgramSteps.Length <= 0 Then Return False

      ' These separators are used by programs and prefixed steps
      Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
      Dim separator1 As String = Convert.ToChar(255)
      Dim separator2 As String = ","

      ' Use this to split out steps in this program
      Dim Steps() As String
      Dim stepParameters() As String
      Dim stepNotes() As String

      ' Loop through the program steps to find the first KP command for each tank
      Dim StartChecking As Boolean = False
      Dim KitchenPrepareFound As Boolean = False
      Dim CommandCode() As String
      Dim CommandPrompt As Integer
      Dim KP1 As String = ""
      Dim IPFound As Integer

      ' Split the program string into an array of program steps - one step per array element
      Steps = ProgramSteps.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

      ' Make sure we've have something to check
      If Steps.GetUpperBound(0) <= 0 Then Return False

      ' Look for the first KP command
      For i As Integer = 0 To Steps.GetUpperBound(0)
        If LookAheadStopFound = True Then Exit For
        If Tank1Enabled = True Then Exit For
        ' Current step command code
        CommandCode = Steps(i).Split(separator2.ToCharArray)

        ' look for the commands we want
        Select Case CommandCode(0)

          Case "IP"
            ' get the program number
            stepNotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            stepParameters = stepNotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            ' Confirm necessary parameters are set
            If stepParameters.GetUpperBound(0) >= 2 Then
              IPFound = CInt(stepParameters(2)) * 256 + CInt(stepParameters(1))
            Else
              IPFound = CInt(stepParameters(1))
            End If
            CheckIP4Program(IPFound)


          Case "KP"
            ' get the program number
            stepNotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            stepParameters = stepNotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            ' Confirm necessary parameters are set
            If stepParameters.GetUpperBound(0) >= 1 Then
              CommandPrompt = CInt(stepParameters(1))
            End If
            If CommandPrompt <> 99 Then
              ' If its not 99 then change it.
              Tank1Calloff = Tank1Calloff + 1
            End If
            Tank1Enabled = True
            Exit For 'we found it we are done.


          Case "LA"
            ' Do nothing if another LA is found before the next KP
            LookAheadRepeatFound = True
            LookAheatRepeat = "Program: " ' & ProgramNumber.Trim & " Step: " & stepNumber.ToString("#0")
            Exit For


          Case "LS"
            LookAheadStopFound = True
            Exit For

        End Select
      Next i

      ' Everything ok
      Return returnValue = True

    Catch ex As Exception
      Utilities.Log.LogError(ex, "CheckIP3ProgramSteps")
    End Try

    ' Return Default Value
    Return returnValue
  End Function

  '***********************************************************************************************
  '******                    Sub IP found within this IP - Get Program Steps                ******
  '***********************************************************************************************
  'this is for a IP inside the third IP
  Private Function CheckIP4Program(ProgramNumber As Long) As Boolean
  Dim sql As String = "SELECT * FROM Programs WHERE ProgramNumber=" & ProgramNumber
    Dim returnValue As Boolean = False
    Try
      With ControlCode

        Dim dt As System.Data.DataTable = .Parent.DbGetDataTable(sql)
        If dt Is Nothing OrElse dt.Rows.Count <> 1 Then Return False

        Dim dr As System.Data.DataRow = dt.Rows(0)
        Dim ProgramSteps As String = Utilities.Null.NullToEmptyString(dr("Steps"))

        returnValue = CheckIP4ProgramSteps(ProgramSteps)

      End With

      ' Everything ok
      Return returnValue

    Catch ex As Exception
      Utilities.Log.LogError(ex.Message)
    End Try

    ' Return Default Value
    Return returnValue
  End Function

  '***********************************************************************************************
  '******                   Check this Sub IP's steps - KP/LA/LS (No More IP's)             ******
  '***********************************************************************************************
  Private Function CheckIP4ProgramSteps(ProgramSteps As String) As Boolean
    Dim returnValue As Boolean = False
    Try

      'Make sure we've got something to check
      If ProgramSteps.Length <= 0 Then Return False

      ' These separators are used by programs and prefixed steps
      Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
      Dim separator1 As String = Convert.ToChar(255)
      Dim separator2 As String = ","

      ' Use this to split out steps in this program
      Dim Steps() As String
      Dim stepParameters() As String
      Dim stepNotes() As String

      ' Loop through the program steps to find the first KP command for each tank
      Dim StartChecking As Boolean = False
      Dim KitchenPrepareFound As Boolean = False
      Dim CommandCode() As String
      Dim CommandPrompt As Integer
      Dim KP1 As String = ""
      Dim IPFound As Integer

      ' Split the program string into an array of program steps - one step per array element
      Steps = ProgramSteps.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

      ' Make sure we've have something to check
      If Steps.GetUpperBound(0) <= 0 Then Return False

      ' Look for the first KP command
      For i As Integer = 0 To Steps.GetUpperBound(0)
        If LookAheadStopFound = True Then Exit For
        If Tank1Enabled = True Then Exit For
        ' Current step command code
        CommandCode = Steps(i).Split(separator2.ToCharArray)

        ' look for the commands we want
        Select Case CommandCode(0)

          Case "IP"
            ' get the program number
            stepNotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            stepParameters = stepNotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            ' Confirm necessary parameters are set
            If stepParameters.GetUpperBound(0) >= 2 Then
              IPFound = CInt(stepParameters(2)) * 256 + CInt(stepParameters(1))
            Else
              IPFound = CInt(stepParameters(1))
            End If
            ' Do not go any further, set out
            Exit For

          Case "KP"
            ' get the program number
            stepNotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            stepParameters = stepNotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            ' Confirm necessary parameters are set
            If stepParameters.GetUpperBound(0) >= 1 Then
              CommandPrompt = CInt(stepParameters(1))
            End If
            If CommandPrompt <> 99 Then
              ' If its not 99 then change it.
              Tank1Calloff = Tank1Calloff + 1
            End If
            Tank1Enabled = True
            Exit For 'we found it we are done.


          Case "LA"
            ' Do nothing if another LA is found before the next KP
            LookAheadRepeatFound = True
            LookAheatRepeat = "Program: " ' & ProgramNumber.Trim & " Step: " & stepNumber.ToString("#0")
            Exit For


          Case "LS"
            LookAheadStopFound = True
            Exit For

        End Select
      Next i

      ' Everything ok
      Return returnValue = True

    Catch ex As Exception
      Utilities.Log.LogError(ex, "CheckIP4ProgramSteps")
    End Try

    ' Return Default Value
    Return returnValue
  End Function

#Region " PROPERTIES "

  Public ReadOnly Property IsOn As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property

  Public ReadOnly Property IsActive() As Boolean
    Get
      Return Not ((State = EState.Off) OrElse (State = EState.CheckNextDyelot))
    End Get
  End Property

#End Region

#Region " I/O PROPERTIES "

#End Region

End Class

