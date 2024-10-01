'American & Efird - Mt Holly
' Version 2024-09-03

Imports Utilities.Translations

<Command("Prepare", "|0-99|% |0-180|F |0-99|mins Mix:|0-1| D:|1-2|", " ", "", ""),
Description("Fills the ColorRoom tank to the low level. Signals the operator to prepare the tank to be transferred to the desired destination [1=Add Tank, 2=Reserve Tank]."),
Category("Kitchen Functions")>
Public Class KP : Inherits MarshalByRefObject : Implements ACCommand
  Private Const commandName_ As String = ("KP: ")
  Private ReadOnly controlCode As ControlCode

  'Command States
  Public Enum Estate
    Off
    Active
    LookAheadInJob
  End Enum
  Property State As Estate
  Property Status As String

  ' Look Ahead Properties
  Public KP1 As Tank1Prepare

  Public Dyelot As String
  Public ReDye As Integer
  Public Job As String

  ' KP CommandStep parameters
  Property Tank1Enabled As Boolean
  Property Tank1FillLevel As Integer
  Property Tank1TempDesired As Integer
  Property Tank1MixTime As Integer
  Property Tank1MixRequested As Integer
  Property Tank1Destination As EKitchenDestination
  Property Tank1Calloff As Integer

  ' LookAhead properties
  Property LookAheadFound As Boolean
  Property LookAheadStopFound As Boolean
  Property LookAheadRepeatFound As Boolean
  Property LookAheatRepeat As String
  Property ProgramHasAnotherTank1 As Boolean

  ' Timers
  Property Timer As New Timer
  Property TimerOverrun As New Timer

  Public Sub New(ByVal controlCode As ControlCode)
    Me.controlCode = controlCode
    Me.controlCode.Commands.Add(Me)
    Me.KP1 = New Tank1Prepare(Me.controlCode)
  End Sub

  Sub ParametersChanged(ByVal ParamArray param() As Integer) Implements ACCommand.ParametersChanged
    'If the command is not running then do nothing the changes will be picked up when the command starts
    If Not IsOn Then Exit Sub

    'Enable command parameter adjustments, if parameter set
    '     If controlCode.Parameters.EnableCommandChg = 1 Then
    '     Start(param)
    '     End If

  End Sub

  Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With controlCode

      'If look ahead is active then copy details and cancel look ahead
      'We need to verify that the LA job that is either active or completed, needs to correspond to the Current Parent.Job that's running
      '  A&E has had issues where they've dispensed the wrong chemicals/dyes to the tank due to the LA being active for the wrong batch
      '  PE's committed is being set, but they're jumping through a procedure to the next batch and the LA states are carrying over:

      'Reset working values
      Job = ""
      Dyelot = ""
      ReDye = 0

      ' Confirm LookAhead status
      If (.LA.KP1.IsOn) OrElse (.LAActive) Then
        ' LookAhead Active, confirm settings and correct if necessary
        If (.LA.Job = "") OrElse (.LA.Job <> Job) Then
          ' Jobs do not match:

          ' Reset Tank1 status
          .Tank1Ready = False
          '  .KitchenJob1 = Job
          ' Get active calloff - TODO - need to grab the correct job associated with the next batch scheduled if looking ahead to next batch

          Tank1Calloff = GetNextCallOff()

          ' Check Command Parameters
          If param.GetUpperBound(0) >= 1 Then Tank1FillLevel = param(1) * 10
          If param.GetUpperBound(0) >= 2 Then Tank1TempDesired = MinMax(param(2) * 10, 0, 1800)
          If param.GetUpperBound(0) >= 3 Then Tank1MixTime = param(3) * 60
          If param.GetUpperBound(0) >= 4 Then Tank1MixRequested = param(4)
          If param.GetUpperBound(0) >= 5 Then Tank1Destination = CType(param(5), EKitchenDestination)

          ' Fix for MixTime and MixRequest parameters- Elid doesn't need the second parameter - if MixTime > 0, then need to mix, otherwise, don't mix
          If (Tank1MixTime = 0) Then Tank1MixRequested = 0
          If (Tank1MixTime > 0) Then Tank1MixRequested = 1


          ' Start KitchenPrepare
          ' KP1.Start(Tank1FillLevel, Tank1TempDesired, Tank1MixTime, Tank1MixRequested = 1, Tank1Calloff, .Parameters.StandardTimeAddPrepare, Tank1Destination)
          KP1.Start(Tank1Calloff, Tank1Destination, Tank1FillLevel, Tank1TempDesired, Tank1MixTime, Tank1MixRequested, .Parameters.StandardTimeAddPrepare)



          ' Set destination ready flags false b/c jobs do not match
          If Tank1Destination = EKitchenDestination.Add Then
            .AddReady = False
            .AddControl.DrainManual()
          End If
          If Tank1Destination = EKitchenDestination.Reserve Then
            .ReserveReady = False
            .ReserveControl.ManualDrain()
          End If

          ' Redye issue is possible, alarm to prevent repeat product preparation
          If (Dyelot = .LA.Dyelot) AndAlso (ReDye = .LA.ReDye - 1) Then
            .Alarms.RedyeIssueCheckMachineTank = True
          End If

          ' Cancel LookAhead command
          .LA.Cancel()
          .LAActive = False
          .LARequest = False
        Else
          ' LookAhead jobs match, copy to current KP
          '  .LA.KP1.CopyTo(KP1)      ' TODO UPdate if using VB6 AddPrepare
          ' Cancel LookAhead
          .LA.KP1.Cancel()
          .LAActive = False
          .LARequest = False
        End If
      Else
        ' LookAhead not active or recently completed
        ' Reset Tank1 status
        .Tank1Ready = False

        ' Get updated Job details
        .LoadRecipeDetailsActiveJob()

        ' Get active calloff
        Tank1Calloff = GetCurrentCallOff()

        ' Check Command Parameters
        If param.GetUpperBound(0) >= 1 Then Tank1FillLevel = param(1) * 10
        If param.GetUpperBound(0) >= 2 Then Tank1TempDesired = MinMax(param(2) * 10, 0, 1800)
        If param.GetUpperBound(0) >= 3 Then Tank1MixTime = param(3) * 60
        If param.GetUpperBound(0) >= 4 Then Tank1MixRequested = param(4)
        If param.GetUpperBound(0) >= 5 Then Tank1Destination = CType(param(5), EKitchenDestination)

        ' Fix for MixTime and MixRequest parameters- Elid doesn't need the second parameter - if MixTime > 0, then need to mix, otherwise, don't mix
        If (Tank1MixTime = 0) Then Tank1MixRequested = 0
        If (Tank1MixTime > 0) Then Tank1MixRequested = 1

        ' Start KitchenPrepare
        ' KP1.Start(Tank1FillLevel, Tank1TempDesired, Tank1MixTime, Tank1MixRequested = 1, Tank1Calloff, .Parameters.StandardTimeAddPrepare, Tank1Destination)
        KP1.Start(Tank1Calloff, Tank1Destination, Tank1FillLevel, Tank1TempDesired, Tank1MixTime, Tank1MixRequested, .Parameters.StandardTimeAddPrepare)

        ' Set destination ready flags false b/c jobs do not match
        If Tank1Destination = EKitchenDestination.Add Then .AddReady = False
        If Tank1Destination = EKitchenDestination.Reserve Then .ReserveReady = False

      End If

      ' Set default start state
      State = Estate.Active

    End With

    'Tell the control system to step on
    Return True
  End Function

  Function Run() As Boolean Implements ACCommand.Run
    With controlCode

      ' Run the tank control module
      KP1.Run()

      Select Case State
        Case Estate.Off
          Status = (" ")

        Case Estate.Active
          Status = Translate("Active")

      End Select

    End With
  End Function

  Sub Cancel() Implements ACCommand.Cancel
    State = Estate.Off
    KP1.Cancel()
    Timer.Cancel()
    TimerOverrun.Cancel()
  End Sub

#Region " PROPERTIES "

  ReadOnly Property IsOn As Boolean Implements ACCommand.IsOn
    Get
      Return State <> Estate.Off
    End Get
  End Property

  ReadOnly Property IsOverrun As Boolean
    Get
      Return IsOn AndAlso TimerOverrun.Finished
    End Get
  End Property

#End Region

#Region " I/O PROPERTIES "

#End Region

#Region " FUNCTIONS "
  Public Function StartLookAhead() As Boolean
    ' Used by active program to automatically begin looking for next KP command
    ' Doesn't require LA command to look to next KP in same job.  
    ' LA used to look to next scheduled dyelot.
    With controlCode
      ' Check safe to start conditions
      If IsOn Then Return False
      If .Parameters.LookAheadEnabled <> 1 Then Return False
      If Not OkayToStartLookAhead() Then Return False

      ' Reset calloff and update
      Tank1Calloff = 0
      Tank1Calloff = GetNextCallOff()

      ' Check to see if there is another kp in this program.
      If CheckExistingProgram() Then
        If ProgramHasAnotherTank1 Then
          .LARequest = False
          State = Estate.LookAheadInJob
          Timer.Seconds = 5
        Else
          Return False
        End If
      End If

    End With
    Return True
  End Function

  Private Function OkayToStartLookAhead() As Boolean
    Try
      With controlCode
        ' We are likely to be called from a transfer or stop step so always start checking from the next step
        Dim startStep = .Parent.CurrentStep
        Dim startStepNumber = .Parent.ProgramStep(.Parent.CurrentStep).StepNumber ' cos of the silly look forward parallel commands

        ' ProgramStep() is a zero based array...
        For i As Integer = startStep To .Parent.ProgramStepCount - 1
          Dim programStep = .Parent.ProgramStep(i)
          If programStep.StepNumber > startStepNumber Then
            Select Case programStep.Command
              Case "LS"
                ' We hit a LookAhead Stop command (LS), do not look any further
                Return False

              Case "KA"
                ' If we hit a transfer command before the prepare command do not start look ahead
                Return False

              Case "PH", "SA"
                ' If we hit a ph check or salt check before the prepare command do not start look ahead
                Return False

              Case "KP"
                ' Found the next prepare command okay to start
                Return True

            End Select
          End If
        Next
      End With
    Catch ex As Exception
      Utilities.Log.LogError(ex)
    End Try
    Return False
  End Function

  Private Function GetCurrentCallOff() As Integer
    With controlCode
      ' Parent.CurrentStep represents the current programStep in relation to all the program steps, regardless of IP's
      ' StepNumber below is calculated as the current programStep within the actual IP active 
      Dim stepNumber = .Parent.ProgramStep(.Parent.CurrentStep).StepNumber
      Return GetCallOff(stepNumber)
    End With
  End Function

  Private Function GetNextCallOff() As Integer
    With controlCode
      Dim stepNumber = .Parent.ProgramStep(.Parent.CurrentStep).StepNumber + 1
      Return GetCallOff(stepNumber)
    End With
  End Function

  Private Function GetCallOff(programStepNumber As Integer) As Integer
    Try
      Dim calloff As Integer = 0
      With controlCode

        ' These separators are used by programs and prefixed steps
        Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
        Dim separator1 As String = Convert.ToChar(255)
        Dim separator2 As String = ","

        ' Get current program and step number - step number corresponds to IP's internal step number
        Dim programStepCount As Integer = .Parent.ProgramStepCount - 1
        Dim programNumber As String = .Parent.ProgramNumberAsString
        Dim stepNumber As Integer : stepNumber = .Parent.StepNumber

        ' Get all program steps for this program
        Dim ProgramSteps = .Parent.PrefixedSteps.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

        ' ProgramStep() is a zero based array...
        ' .Parent.ProgramStepCount is the total number of steps in the current program, including IP's
        Dim QQFound As Boolean = False
        For i As Integer = 0 To programStepCount
          Dim programStep = .Parent.ProgramStep(i)

          If programStep.Command = "QQ" Then QQFound = True
          If programStep.Command = "EI" Then QQFound = False

          ' Don't count preps inside a QQ statement
          If Not QQFound Then
            Select Case programStep.Command
              Case "IP"
                ' TODO Look inside IP to count each KP
                '  calloff += 1
                If (programStep.StepNumber >= programStepNumber) Then Exit For
              Case "KP"
                'TODO
                calloff += 1
                If (programStep.StepNumber >= (programStepNumber - 1)) Then Exit For
            End Select
          End If
        Next i

        ' Test MW's format from CC - SchollCharge
        Dim startChecking As Boolean = False
        Dim command() As String
        Dim testCalloff As Integer = 0
        QQFound = False
        'Loop through each step starting from the beginning of the program and count the number of adds
        For j As Integer = ProgramSteps.GetLowerBound(0) To ProgramSteps.GetUpperBound(0)
          'Ignore step 0
          If j > 0 Then
            Dim StepArray = ProgramSteps(j).Split(separator1.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            ' Confirm StepArray is made of something (>=1)
            If StepArray.GetUpperBound(0) >= 1 Then
              Dim currentStepNumber = CDbl(StepArray(1))
              Dim currentProgramNumber = StepArray(0)
              If currentProgramNumber = programNumber.ToString AndAlso (currentStepNumber = stepNumber - 1) Then
                startChecking = True
              End If
              command = StepArray(5).Split(separator2.ToCharArray)
              If command(0) = "QQ" Then QQFound = True
              If command(0) = "EI" Then QQFound = False

              If startChecking Then
                If (command(0) = "KP") AndAlso Not (QQFound) Then
                    testCalloff += 1
                  Exit For
              End If
              Else
                ' StartCheck=false >> Loop active prior to the current step, count any adds
                If (command(0) = "KP") AndAlso Not (QQFound) Then
                  testCalloff += 1
            End If

          End If
            End If

          End If
          ' Hit upper bounds without finding a KP
          If j = ProgramSteps.GetUpperBound(0) Then testCalloff = 0
        Next j

        ' Testing
        calloff = testCalloff
      End With

      Return calloff

    Catch ex As Exception
      Utilities.Log.LogError(ex)
    End Try
    Return 0
  End Function


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
          stepParameters = stepNotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0-based array

          ' 5 parameters defined for KP command: ["|0-99|% |0-180|F |0-99|mins Mix:|0-1| D:|1-2|"]
          If stepParameters.GetUpperBound(0) >= 5 Then
            '0 = command '1 = Level, 2 = desired temp, 3 = mixing time, 4 = mixer on or off, 5 = Destination

            ' Update local variables with command step parameters
            Tank1FillLevel = CInt(stepParameters(1)) * 10
            Tank1TempDesired = CInt(stepParameters(2)) * 10
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

#End Region

#If 0 Then
'===============================================================================================
'KP - Kitchen tank prepare command
'===============================================================================================
Option Explicit
Implements ACCommand
Public KP1 As New acAddPrepare
Public Calloff As Long

'Private storage for properties
Private pJob As String
Private pDyelot As String
Private pRedye As Long

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Parameters=|0-99|% |0-180|F |0-99|M Mix?|0-1| D:|1-2|\r\nName=Kitchen Prepare \r\nHelp=Fills the drugroom tank to the specified level with the specified water type."
Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  
  With ControlCode
    'If look ahead is active then copy details and cancel look ahead
    'We need to verify that the LA job that is either active or completed, needs to correspond to the Current Parent.Job that's running
    '  A&E has had issues where they've dispensed the wrong chemicals/dyes to the tank due to the LA being active for the wrong batch
    '  PE's committed is being set, but they're jumping through a procedure to the next batch and the LA states are carrying over:
    
    ResetAll

    'Check to see if there is another kp in this program.
    If GetDyelotName(ControlCode) Then
      'dyelot/redy/job determined correctly
    Else
      pDyelot = "Error"
    End If
  
    If .LA.KP1.IsOn Then
      If (.LA.Job = "") Or (.LA.Job <> Job) Then
        .ManualAddDrainPB = True
        .ManualReserveDrainPB = True
        .Tank1Ready = False
        .DrugroomDisplayJob1 = Job
        Calloff = GetKPCalloff(ControlObject)
        'Make sure the parameters are what we expect (Destination 1=AddTank / 2=ReserveTank)
        If UBound(Param) >= 5 Then
          If IsNumeric(Param(1)) And IsNumeric(Param(2)) And IsNumeric(Param(3)) And IsNumeric(Param(4)) And IsNumeric(Param(5)) Then
            KP1.Start ((CLng(Param(1)) * 10) - .Parameters.Tank1FillLevelDeadBand), ((CLng(Param(2)) * 10) - .Parameters.Tank1HeatDeadband), _
            CLng(Param(3)) * 60, CLng(Param(4)), Calloff, 300, CLng(Param(5))
          End If
        End If
        'set destination tank false if jobs do not match
        If CLng(Param(5) = 1) Then
          .AddReady = False
        ElseIf CLng(Param(5) = 2) Then
          .ReserveReady = False
        End If
        'If @1 issue is possible, alarm for it to prevent repeat product preparation
        If (pDyelot = .LA.Dyelot) And (pRedye = .LA.Redye - 1) Then
          KP1.AlarmRedyeIssue = True
        End If
        .LA.Cancel
        .LAActive = False
        .LARequest = False
      Else
        .LA.KP1.CopyTo KP1
        .LA.KP1.Cancel
        .LAActive = False
        .LARequest = False
      End If
    ElseIf .LAActive Then
      If (.LA.Job = "") Or (.LA.Job <> Job) Then
        .ManualAddDrainPB = True
        .ManualReserveDrainPB = True
        .Tank1Ready = False
        .DrugroomDisplayJob1 = Job
        Calloff = GetKPCalloff(ControlObject)
        'Make sure the parameters are what we expect (Destination 1=AddTank / 2=ReserveTank)
        If UBound(Param) >= 5 Then
          If IsNumeric(Param(1)) And IsNumeric(Param(2)) And IsNumeric(Param(3)) And IsNumeric(Param(4)) And IsNumeric(Param(5)) Then
            KP1.Start ((CLng(Param(1)) * 10) - .Parameters.Tank1FillLevelDeadBand), ((CLng(Param(2)) * 10) - .Parameters.Tank1HeatDeadband), _
            CLng(Param(3)) * 60, CLng(Param(4)), Calloff, 300, CLng(Param(5))
          End If
        End If
        'set destination tank false if jobs do not match
        If CLng(Param(5) = 1) Then
          .AddReady = False
        ElseIf CLng(Param(5) = 2) Then
          .ReserveReady = False
        End If
        'If @1 issue is possible, alarm for it to prevent repeat product preparation
        If (pDyelot = .LA.Dyelot) And (pRedye = .LA.Redye - 1) Then
          KP1.AlarmRedyeIssue = True
        End If
        .LA.Cancel
        .LAActive = False
        .LARequest = False
      Else
        .LAActive = False
        .LARequest = False
      End If
    Else
      'If LA Command is not active then set ready to false
      .Tank1Ready = False
      .DrugroomDisplayJob1 = Job
      Calloff = GetKPCalloff(ControlObject)
      'Make sure the parameters are what we expect (Destination 1=AddTank / 2=ReserveTank)
      If UBound(Param) >= 5 Then
        If IsNumeric(Param(1)) And IsNumeric(Param(2)) And IsNumeric(Param(3)) And IsNumeric(Param(4)) And IsNumeric(Param(5)) Then
          KP1.Start ((CLng(Param(1)) * 10) - .Parameters.Tank1FillLevelDeadBand), ((CLng(Param(2)) * 10) - .Parameters.Tank1HeatDeadband), _
          CLng(Param(3)) * 60, CLng(Param(4)), Calloff, 300, CLng(Param(5))
        End If
      End If
    End If
  End With
  
'Carry on my wayward son...
  StepOn = True
  
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode
     KP1.Run ControlCode
  End With
  End Sub

Public Sub ACCommand_Cancel()
  KP1.Cancel
End Sub

Private Sub ResetAll()
'Reset command state & variables
  pJob = ""
  pDyelot = ""
  pRedye = 0
End Sub

'this function determines the dyelot Name and redye due to issue with Parent.Job not being accurate
Public Function GetDyelotName(ByVal ControlObject As Object) As Boolean
On Error GoTo Err_GetDyelotName

'Added this to get the current job (dyelot and redye) from the database instead of from parent.job and splitting due to exception
'added 8/5/2009

'Create an ADO connection to the local database
  Dim con As ADODB.Connection15
  Set con = New ADODB.Connection
  con.Open "BatchDyeing"
  
  Dim rsDyelot As ADODB.Recordset15
  Set rsDyelot = con.Execute("SELECT * FROM Dyelots WHERE State=2 order by StartTime")
  rsDyelot.MoveFirst

  If Not rsDyelot.EOF Then
    'Set Dyelot and Redye - these must be non-null because they are the primary key
    pDyelot = CStr(rsDyelot!Dyelot)
    pRedye = CLng(rsDyelot!Redye)
  End If

'everything ok
  GetDyelotName = True

'Avoid Error Handler
  Exit Function
  
Err_GetDyelotName:
  GetDyelotName = False
End Function

Public Property Get Job() As String
  Job = Dyelot
  If Redye > 0 Then Job = Dyelot & "@" & Redye
End Property
Public Property Get Dyelot() As String
  Dyelot = pDyelot
End Property
Public Property Get Redye() As Long
  Redye = pRedye
End Property
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property
Friend Property Get IsOn() As Boolean
  If KP1.IsOn Then IsOn = True
End Property

'Used By Drugroom Preview (Do Not Delete)
Public Property Get Tank1DispenseError() As Long
 Tank1DispenseError = KP1.AddDispenseError
End Property

#End If
End Class
