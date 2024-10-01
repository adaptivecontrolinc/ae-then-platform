'American & Efird - Mt. Holly Then Platform 
' Version 2024-08-23

Imports Utilities.Translations

Public Class AddControl
  Inherits MarshalByRefObject

  ' For convenience
  Private Property parent As ACParent
  Private Property controlCode As ControlCode

  Private _timerState As New Timer
  Private _alarmTimer As New Timer
  Private _timerMix As New Timer

  ' Public Timer Value Display
  Public Property StateTimerSecs As Integer = _timerState.Seconds
  Public Property AlarmTimerSecs As Integer = _alarmTimer.Seconds

  Public Property FillType As EFillType
  Public Property FillLevel As Integer
  Public Property MixState As EMixState
  Public Property NumberOfRinses As Integer
  Public Property NumberOfDrains As Integer

  Public Enum EState
    Off
    Ready

    ' DISPENSER STATES
    '    WaitIdle
    '    PreFill
    '    DispenseWaitTurn
    '    DispenseWaitReady
    '    DispenseWaitProducts
    '    DispenseWaitResponse

    ' FILL STATES
    FillPause
    FillStart
    FillInterlock
    FillWithWater
    FillRunbackPause
    FillRunback
    FillHeatToTemp
    FillComplete

    ' TODO code for circulation mix
    '   If AdditionLevel > Parameters_AddMixOnLevel Then AdditionMixOn = True
    '   If AdditionLevel <= Parameters_AddMixOffLevel Then AdditionMixOn = False
    MixingPause
    Mixing

    ' TRANSFER STATES - DO NOT INCLUDE TRANSFER STATES IN THIS CONTROL - USE AT/AC FOR THAT FUNCTION
    '    TransferPause
    '    TransferStart
    '    TransferEmpty1
    '    TransferRinse
    '    TransferEmpty2
    '    TransferComplete


    ' DRAIN, RINSE, DRAIN - MANUAL/AUTO DRAINS START HERE - OPERATOR CAN CANCEL AT ANYTIME
    DrainPause
    DrainStart
    DrainInterlock
    DrainEmpty1
    DrainRinseToDrain
    DrainDrainEmpty2
    DrainComplete

    ' FILL, MIX, AND DRAIN
    DrainFlushPause
    DrainFlushStart
    DrainFlushFill
    DrainFlushCirculate
    DrainFlushCloseMixForTime
    DrainFlushTransfer1
    DrainFlushRinseToDrain
    DrainFlushTransfer2
    DrainFlushComplete

  End Enum
  Public State As EState
  Public StateRestart As EState
  Public Status As String

  Private stateRestartTime_ As Integer



  Public Sub New(ByVal parent As ACParent, ByVal controlCode As ControlCode)
    Me.parent = parent
    Me.controlCode = controlCode
    Me._timerState.Seconds = 10
  End Sub

  Public Sub Run()
    Try
      With controlCode

        ' Pause the state machine under these conditions
        Dim pauseFill As Boolean = .Parent.IsPaused OrElse .EStop
        '   Dim pauseTransfer As Boolean = .Parent.IsPaused OrElse .EStop OrElse (Not .TempSafe) OrElse (Not .MachineClosed)
        Dim pauseDrainEmpty As Boolean = .Parent.IsPaused OrElse .EStop
        Dim pauseDrainFlush As Boolean = .Parent.IsPaused OrElse .EStop

        'Remember the start state so we can loop until all state changes have completed
        '  prevents IO flicker as we move through the state machine
        Static startState As EState
        Do
          ' Set default restart state & time
          If pauseFill AndAlso (State >= EState.FillStart) AndAlso (State < EState.FillComplete) Then
            StateRestart = State
            stateRestartTime_ = _timerState.Seconds
            State = EState.FillPause
          End If
          If pauseDrainEmpty AndAlso (State >= EState.DrainStart) AndAlso (State < EState.DrainComplete) Then
            StateRestart = State
            stateRestartTime_ = _timerState.Seconds
            State = EState.DrainPause
          End If
          If pauseDrainFlush AndAlso (State >= EState.DrainFlushStart) AndAlso (State <= EState.DrainFlushTransfer2) Then
            StateRestart = State
            stateRestartTime_ = _timerState.Seconds
            State = EState.DrainFlushPause
          End If

          'Save start state so we can see if there has been a state change at the end
          '  normally we keep looping until there are no state changes
          startState = State
          Select Case State

            Case EState.Off
              If _timerState.Finished Then State = EState.Ready
              Status = Translate("Startup") & _timerState.ToString(1)

            Case EState.Ready
              Status = Translate("Idle")


              '********************************************************************************************
              '******   FILL THE ADD TANK
              '********************************************************************************************
            Case EState.FillPause
              Status = Translate("Paused") & _timerState.ToString(1)
              If pauseFill Then
                If .EStop Then Status = ("EStop")
                If .Parent.IsPaused Then Status = ("Paused")
                _timerState.Seconds = 2
              End If
              If _timerState.Finished Then
                State = EState.FillStart
              End If

            Case EState.FillStart
              If MixState = EMixState.Active Then
                ' Make sure that Fill Level is greater than the minimum add mix level
                If FillLevel < (.Parameters.AddMixOnLevel + 50) Then FillLevel = (.Parameters.AddMixOnLevel + 50)
              End If
              If _timerState.Finished Then
                State = EState.FillInterlock
              End If
              Status = Translate("Fill Starting") & _timerState.ToString(1)

            Case EState.FillInterlock
              Status = Translate("Fill Interlock")
              If .KA.IsOn AndAlso (.KA.Destination = EKitchenDestination.Add) Then
                _timerState.Seconds = 2
                Status = Translate("Kitchen") & (" ") & .KA.Status
                'ElseIf .KP.IsOn AndAlso (.KP.KP1.Param_Destination = EKitchenDestination.Add) AndAlso (.KP.KP1.IsForeground) Then
              ElseIf .KP.IsOn AndAlso (.KP.KP1.DispenseTank = EKitchenDestination.Add) Then ' TODO AndAlso (.KP.KP1.IsForeground) Then
                _timerState.Seconds = 2
                Status = Translate("Kitchen") & (" ") & .KP.KP1.Status
              ElseIf .LA.IsOn AndAlso (.LA.KP1.Param_Destination = EKitchenDestination.Add) Then ' TODO AndAlso (.KP.KP1.IsForeground) Then
                _timerState.Seconds = 2
                Status = Translate("Kitchen") & (" ") & .LA.KP1.Status
              Else
                ' Kitchen complete or idle
                If FillType = EFillType.Vessel Then
                  If Not (.MachineClosed OrElse .TempSafe) Then
                    _timerState.Seconds = 2
                    If Not .TempSafe Then Status = Translate("Wait for TempSafe") & _timerState.ToString(1)
                    If Not .MachineClosed Then Status = Translate("Machine Not Closed")
                  End If
                Else
                  Status = Translate("Starting") & _timerState.ToString(1)
                End If
              End If
              ' Reset alarm timer
              _alarmTimer.Minutes = 5
              ' Next state
              If _timerState.Finished Then
                If FillType = EFillType.Vessel Then
                  State = EState.FillRunback
                  _timerState.Seconds = .Parameters.AddRunbackPulseTime
                Else
                  State = EState.FillWithWater
                End If
              End If

            Case EState.FillWithWater
              ' Deadband parameter to assist with possible overshoot
              ' Level Reached
              If .AddLevel >= (FillLevel - .Parameters.AdditionFillDeadband) Then
                ' Mix Requested - Confirm Mix Level reached
                If MixState = EMixState.Active Then
                  ' Confirm level above MixOnLevel
                  If .AddLevel > .Parameters.AddMixOnLevel Then
                    State = EState.Mixing
                  End If
                Else
                  ' Mixing Not Requested
                  State = EState.Ready
                End If
              End If
              Status = Translate("Filling") & (" ") & (.AddLevel / 10).ToString("#0.0") & " / " & (FillLevel / 10).ToString("#0.0") & "%"

            Case EState.FillRunbackPause
              ' Runback Pause Timer Finished
              If _timerState.Finished OrElse (.Parameters.AddRunbackPulseTime = 0) Then
                State = EState.FillRunback
                _timerState.Seconds = .Parameters.AddRunbackPulseTime
              End If
              ' Level Reached
              If .AddLevel >= (FillLevel - .Parameters.AdditionFillDeadband) Then
                ' Mix Requested - Confirm Mix Level reached
                If MixState = EMixState.Active Then
                  If .AddLevel > .Parameters.AddMixOnLevel Then
                    State = EState.Mixing
                  Else
                    State = EState.Ready
                  End If
                End If
              End If
              Status = Translate("Runback") & (" ") & Translate("Paused") & (" ") & (.AddLevel / 10).ToString("#0.0") & " / " & (FillLevel / 10).ToString("#0.0") & "% "

            Case EState.FillRunback
              ' Runback Pause Timer
              If _timerState.Finished AndAlso (.Parameters.AddRunbackPulseTime > 0) Then
                State = EState.FillRunbackPause
                _timerState.Seconds = .Parameters.AddRunbackPulseTime
              End If
              ' Level Reached
              If .AddLevel >= (FillLevel - .Parameters.AdditionFillDeadband) Then
                ' Mix Requested - Confirm Mix Level reached
                If MixState = EMixState.Active Then
                  If .AddLevel > .Parameters.AddMixOnLevel Then
                    State = EState.Mixing
                  Else
                    State = EState.Ready
                  End If
                End If
              End If
              Status = Translate("Filling") & (" ") & (.AddLevel / 10).ToString("#0.0") & " / " & (FillLevel / 10).ToString("#0.0") & "% "


            Case EState.FillHeatToTemp


            Case EState.FillComplete
              Status = Translate("Completing") & _timerState.ToString(1)
              If _timerState.Finished Then State = EState.Ready


              '********************************************************************************************
              '******   MIX THE ADD TANK CONTENTS
              '********************************************************************************************
            Case EState.MixingPause
              Status = Translate("Mixing") & (" ") & Translate("Paused")

            Case EState.Mixing
              Status = Translate("Mixing")
              If MixState < EMixState.Active Then State = EState.Ready


              '********************************************************************************************
              '******   TRANSFER THE ADD TANK CONTENTS
              '********************************************************************************************
              '            Case EState.TransferPause
              '
              '            Case EState.TransferStart
              '              StateString = Translate("Transfer") & (" ") & _timerState.ToString
              'If pauseTransfer Then
              '  If Not .MachineClosed Then StateString = Translate("Transfer") & (" ") & Translate("Machine Not Closed")
              '  If Not .SampleClosed Then StateString = Translate("Transfer") & (" ") & Translate("Sample Not Closed")
              '  If Not .TempSafe Then StateString = Translate("Transfer") & (" ") & Translate("Temp Not Safe")
              '  If .EStop Then StateString = "ESTOP"
              'Else
              '  If _timerState.Finished Then
              '    State = EState.TransferEmpty1
              '    _timerState.Seconds = .Parameters.AddTimeBeforeRinse
              '  End If
              'End If

              'Case EState.TransferEmpty1
              '  If .AddLevel > 10 Then
              '    StateString = Translate("Transferring") & (" ") & (.AddLevel / 10).ToString("#0.0") & "%"
              '    _timerState.Seconds = .Parameters.AddTimeBeforeRinse
              '  Else
              '    StateString = Translate("Transferring") & (" ") & _timerState.ToString
              '  End If
              '  If _timerState.Finished Then
              '    State = EState.TransferRinse
              '    _timerState.Seconds = .Parameters.AddTimeToRinseToMachine
              '  End If

              'Case EState.TransferRinse
              '  StateString = Translate("Rinse to machine") & (" ") & _timerState.ToString
              '  If _timerState.Finished Then
              '    State = EState.TransferEmpty2
              '    _timerState.Seconds = .Parameters.AddTransferSettleTime
              '  End If

              'Case EState.TransferEmpty2
              '  If .AddLevel > 10 Then
              '    StateString = Translate("Rinse to machine") & (" ") & (.AddLevel / 10).ToString("#0.0") & "%"
              '    _timerState.Seconds = .Parameters.AddTimeAfterRinse
              '  Else
              '    StateString = Translate("Rinse to machine") & (" ") & _timerState.ToString
              '  End If
              '  If _timerState.Finished Then
              '    State = EState.Ready
              '  End If

              'Case EState.TransferComplete
              '  ' Just used as a common placeholder for property logic 
              '  ' IsActive = (state>start) andalso (state<complete)


              '********************************************************************************************
              '******   DRAIN THE ADD TANK CONTENTS
              '********************************************************************************************
            Case EState.DrainPause
              Status = Translate("Draining") & (" ") & Translate("Paused")
              If .EStop Then Status = Translate("Estop")
              If Not pauseDrainEmpty Then
                State = EState.DrainEmpty1
                _timerState.Seconds = .Parameters.AddTransferTimeBeforeRinse
                _alarmTimer.Seconds = .Parameters.AddTransferTimeOverrun
              End If

            Case EState.DrainStart
              Status = Translate("Draining") & (" ") & Translate("Starting") & _timerState.ToString(1)
              If _timerState.Finished Then
                State = EState.DrainInterlock
              End If

            Case EState.DrainInterlock
              Status = Translate("Draining") & (" ") & Translate("Interlocked")
              If .EStop Then Status = Translate("Estop")
              If .Parent.IsPaused Then Status = Translate("Paused")
              If Not pauseDrainEmpty Then
                State = EState.DrainEmpty1
                _timerState.Seconds = .Parameters.AddTransferTimeBeforeRinse
                _alarmTimer.Seconds = .Parameters.AddTransferTimeOverrun
              End If

            Case EState.DrainEmpty1
              Status = Translate("Draining")
              If .AddLevel > 50 Then
                Status = Status & (" ") & (.AddLevel / 10).ToString("#0.0") & ("%")
                _timerState.Seconds = .Parameters.AddTransferTimeBeforeRinse
                'Continuously check current level with previous level to make sure where transferring
                CheckLevelDecreasing(.AddLevel, .Parameters.AddPumpPauseTime)
              Else
                Status = Status & (" ") & _timerState.ToString
                ResetLevelCheck(.AddLevel)
              End If
              ' Drain Timer finishes OR Alarm Delay finished
              If _timerState.Finished OrElse (_alarmTimer.Finished AndAlso (.Parameters.AddTransferTimeOverrun > 0)) Then
                State = EState.DrainRinseToDrain
                _timerState.Seconds = MinMax(.Parameters.AddTimeRinseToDrain, 5, 120)
                _alarmTimer.Seconds = .Parameters.AddTransferTimeOverrun
              End If

            Case EState.DrainRinseToDrain
              Status = Translate("Rinse To Drain") & _timerState.ToString(1)
              ' Transferring completed OR Alarm Delay finished
              If _timerState.Finished OrElse (_alarmTimer.Finished AndAlso (.Parameters.AddTransferTimeOverrun > 0)) Then
                State = EState.DrainDrainEmpty2
                _timerState.Seconds = .Parameters.AddTimeToDrain
              End If

            Case EState.DrainDrainEmpty2
              Status = Translate("Draining")
              If .AddLevel > 10 Then
                Status = Status & (" ") & (.AddLevel / 10).ToString("#0.0") & "%"
                _timerState.Seconds = .Parameters.AddTimeToDrain
                'Continuously check current level with previous level to make sure where transferring
                CheckLevelDecreasing(.AddLevel, .Parameters.AddPumpPauseTime)
              Else
                Status = Status & _timerState.ToString(1)
                ResetLevelCheck(.AddLevel)
              End If
              ' Transferring completed
              If _timerState.Finished Then
                State = EState.DrainComplete
                _timerState.Seconds = 2
              End If
              ' Drain too long enabled
              If _alarmTimer.Finished AndAlso (.Parameters.AddTransferTimeOverrun > 0) Then
                State = EState.DrainComplete
                _timerState.Seconds = 2
              End If

            Case EState.DrainComplete
              Status = Translate("Draining") & (" ") & Translate("Complete") & _timerState.ToString(1)
              ' Delay completed
              If _timerState.Finished Then
                State = EState.DrainFlushFill
              End If


              '********************************************************************************************
              '******   FILL, CIRCULATE, AND SEND TO DRAIN
              '********************************************************************************************
            Case EState.DrainFlushPause
              Status = Translate("Paused")
              If .EStop Then Status = Translate("EStop")
              If Not (.EStop OrElse .Parent.IsPaused) Then
                State = StateRestart
                _timerState.Seconds = stateRestartTime_
              End If

            Case EState.DrainFlushStart
              Status = Translate("Flush Starting") & _timerState.ToString(1)
              If pauseDrainFlush Then
                Status = Translate("Flush Paused") & _timerState.ToString(1)
                _timerState.Seconds = 2
              End If
              If _timerState.Finished Then
                State = EState.DrainFlushFill
              End If

            Case EState.DrainFlushFill
              Dim fillLevel As Integer = MinMax(.Parameters.AddRinseFillLevel, .Parameters.AddMixOnLevel, 500)
              Status = Translate("Filling") & (" ") & (.AddLevel / 10).ToString("#0.0") & (" / ") & (fillLevel / 10).ToString("#0.0") & "% "
              If .AddLevel > fillLevel Then
                State = EState.DrainFlushCirculate
                _timerState.Seconds = MinMax(.Parameters.AddRinseMixTime, 5, 120)
                _timerMix.Seconds = MinMax(.Parameters.AddMixPulseTime, 5, 60)
              End If
              ' Update Restart State
              StateRestart = EState.DrainFlushFill

            Case EState.DrainFlushCirculate
              ResetLevelCheck(.AddLevel)
              If _timerMix.Finished Then
                State = EState.DrainFlushCloseMixForTime
                _timerMix.Seconds = 1
              End If
              If _timerState.Finished Then
                State = EState.DrainFlushTransfer1
                NumberOfDrains -= 1
                _timerState.Seconds = .Parameters.AddTimeToDrain
                _alarmTimer.Seconds = .Parameters.AddTransferTimeOverrun
              End If
              Status = Translate("Mixing") & _timerState.ToString(1)

            Case EState.DrainFlushCloseMixForTime
              If _timerMix.Finished Then
                State = EState.DrainFlushCirculate
                _timerMix.Seconds = .Parameters.AddMixPulseTime
              End If
              If _timerState.Finished Then
                State = EState.DrainFlushTransfer1
                NumberOfDrains -= 1
                _timerState.Seconds = .Parameters.AddTimeToDrain
                _alarmTimer.Seconds = .Parameters.AddTransferTimeOverrun
              End If
              Status = Translate("Mixing") & _timerState.ToString(1)

            Case EState.DrainFlushTransfer1
              Status = Translate("Draining")
              If .AddLevel > 10 Then
                Status = Status & (" ") & (.AddLevel / 10).ToString("#0.0") & "%"
                _timerState.Seconds = .Parameters.AddTimeToDrain
                'Continuously check current level with previous level to make sure where transferring
                CheckLevelDecreasing(.AddLevel, .Parameters.AddPumpPauseTime)
              Else
                Status = Status & _timerState.ToString(1)
                ResetLevelCheck(.AddLevel)
              End If
              ' Transfer Timer finished OR Drain Too Long & Enabled
              If _timerState.Finished OrElse (_alarmTimer.Finished AndAlso (.Parameters.AddTransferTimeOverrun > 0)) Then
                If NumberOfDrains > 0 Then
                  ' Repeat Flush sequence: Fill/Circulate/Drain 
                  State = EState.DrainFlushFill
                Else
                  State = EState.DrainFlushRinseToDrain
                  NumberOfRinses = MinMax(.Parameters.AddRinses, 0, 10)
                  _timerState.Seconds = MinMax(.Parameters.AddTimeRinseToDrain, 5, 120)
                End If
              End If
              ' Update Restart State
              StateRestart = EState.DrainFlushRinseToDrain


              '********************************************************************************************
              '******   WALL WASHING RINSE TO DRAIN
              '********************************************************************************************
            Case EState.DrainFlushRinseToDrain
              Status = Translate("Rinse To Drain") & _timerState.ToString(1)
              If _timerState.Finished Then
                NumberOfRinses -= 1
                State = EState.DrainFlushTransfer2
                _timerState.Seconds = .Parameters.AddTransferTimeBeforeRinse
                _alarmTimer.Seconds = .Parameters.AddTransferTimeOverrun
              End If
              ' Update Restart State
              StateRestart = EState.DrainFlushTransfer2

            Case EState.DrainFlushTransfer2
              If .AddLevel > 10 Then
                Status = Translate("Draining") & (" ") & (.AddLevel / 10).ToString("#0.0") & "%" & _timerState.ToString(1)
                _timerState.Seconds = .Parameters.AddTransferTimeBeforeRinse
              Else : Status = Translate("Draining") & _timerState.ToString(1)
              End If
              ' Transfer completed
              If _timerState.Finished Then
                If (NumberOfRinses > 0) Then
                  State = EState.DrainFlushRinseToDrain
                  _timerState.Seconds = MinMax(.Parameters.AddTimeRinseToDrain, 5, 120)
                Else
                  State = EState.DrainFlushComplete
                  _timerState.Seconds = 5
                End If
              End If
              ' Drain too long enabled - Separate from stateTimer so to prevent Repeat Rinse w/ Error
              If _alarmTimer.Finished AndAlso (.Parameters.AddTransferTimeOverrun > 0) Then
                State = EState.DrainFlushComplete
                _timerState.Seconds = 5
              End If
              ' Update Restart State
              StateRestart = EState.DrainFlushTransfer2

            Case EState.DrainFlushComplete
              Status = Translate("Completing") & _timerState.ToString(1)
              If _timerState.Finished Then State = EState.Ready

          End Select
        Loop Until (State = startState)

      End With

    Catch ex As Exception
      Dim message As String = ex.Message
    End Try
  End Sub

#Region " CANCEL LOGIC "

  Public Sub Cancel()
    State = EState.Ready
    StateRestart = EState.Ready
    NumberOfRinses = 0
    NumberOfDrains = 0
    MixState = EMixState.Off
    FillType = EFillType.Cold
    FillLevel = 0
  End Sub

  Public ReadOnly Property CancelActive As Boolean
    Get
      Return (State = EState.FillComplete) OrElse (State = EState.DrainComplete)
    End Get
  End Property

  Public ReadOnly Property IsReady() As Boolean
    Get
      Return (State = EState.Ready)
    End Get
  End Property

  Public ReadOnly Property IsActive As Boolean
    Get
      Return FillIsActive OrElse MixIsActive OrElse DrainIsActive
    End Get
  End Property

#End Region

#Region " FILL PROPERTIES / SUBROUTINES "

  Public Sub FillAuto(ByVal fillType As EFillType, ByVal fillLevel As Integer, ByVal mixState As EMixState)
    With controlCode

      ' Update Local Parameters
      Me.FillType = fillType
      Me.FillLevel = fillLevel
      Me.MixState = mixState

      ' Restrict the Fill Type [2016-08-31]
      'If (.Parameters.FillEnableBlend = 0) AndAlso fillType = EFillType.Mix Then fillType = EFillType.Hot
      'If (.Parameters.FillEnableHot = 0) Then fillType = EFillType.Cold

      ' Set default state
      State = EState.FillStart
      _timerState.Seconds = 2

    End With
  End Sub

  Public Sub FillCancel()
    If FillIsOn Then
      State = EState.FillComplete
      _timerState.Seconds = 1
    End If
  End Sub

  Public Sub FillManual(ByVal fillType As EFillType, ByVal level As Integer)
    With controlCode

      Me.FillType = fillType
      Me.FillLevel = MinMax(level, 0, 1000)

      ' Restrict the Fill Type [2016-08-31]
      'If (.Parameters.FillEnableBlend = 0) AndAlso fillType = EFillType.Mix Then fillType = EFillType.Hot
      'If (.Parameters.FillEnableHot = 0) Then fillType = EFillType.Cold

      ' Set default state
      State = EState.FillStart
      _timerState.Seconds = 2

    End With
  End Sub

  Public ReadOnly Property FillIsActive As Boolean
    Get
      Return (State >= EState.FillStart) AndAlso (State <= EState.FillWithWater)
    End Get
  End Property

  Friend ReadOnly Property FillIsOn As Boolean
    Get
      Return (State >= EState.FillPause) AndAlso (State <= EState.FillComplete)
    End Get
  End Property

  Friend ReadOnly Property FillIsForeground As Boolean
    Get
      Return (FillType = EFillType.Vessel) AndAlso FillIsActive
    End Get
  End Property

  Friend ReadOnly Property FillIsPaused As Boolean
    Get
      Return (State = EState.FillPause)
    End Get
  End Property

#End Region

#Region " MIXING LOGIC "


  Public ReadOnly Property MixIsActive As Boolean
    Get
      Return (State = EState.Mixing)
    End Get
  End Property


#End Region

#Region " TRANSFER LOGIC "

  'Public Sub TransferManual()
  '  State = EState.TransferStart
  '  _timerState.Seconds = 2
  'End Sub

  'Public ReadOnly Property TransferIsActive As Boolean
  '  Get
  '    Return (State >= EState.TransferPause) AndAlso (State < EState.TransferComplete)
  '  End Get
  'End Property


#End Region

#Region " DRAIN LOGIC "

  Public Sub DrainAuto(ByVal numberOfDrains As Integer)
    Me.NumberOfDrains = numberOfDrains
    State = EState.DrainFlushStart
  End Sub

  Public Sub DrainManual()
    With controlCode
      State = EState.DrainStart
      NumberOfDrains = 0
      NumberOfRinses = MinMax(.Parameters.AddRinses, 0, 10)
      _timerState.Seconds = 5
    End With
  End Sub

  Public Sub DrainCancel()
    '   If FillIsActive Then State = EState.DrainFlushRinseToDrain
    State = EState.Ready
    StateRestart = EState.Ready

    _pumpPause = False

    NumberOfDrains = 0
    NumberOfRinses = 0

    ' TODO - add a state test logic if Currently Draining, and if so - perform final drain empty (without rinse)

  End Sub

  Public ReadOnly Property DrainIsActive As Boolean
    Get
      'TODO - Check
      '  Return (State >= EState.DrainStart) AndAlso (State < EState.DrainComplete)
      Return (State >= EState.DrainPause) AndAlso (State <= EState.DrainFlushComplete)

    End Get
  End Property

  Friend ReadOnly Property DrainIsForeground As Boolean
    Get
      Return (State >= EState.DrainInterlock) AndAlso (State <= EState.DrainDrainEmpty2)
    End Get
  End Property

  Friend ReadOnly Property DrainIsPaused As Boolean
    Get
      Return (State = EState.DrainPause)
    End Get
  End Property

#End Region

#Region " LEVEL MONITORING LOGIC "

  Private _timerPump As New Timer
  Private _pumpPause As Boolean
  Private _levelLast As Integer

  Private Sub ResetLevelCheck(ByVal currentlevel As Integer)
    _timerPump.TimeRemaining = 10
    _pumpPause = False
    _levelLast = currentlevel
  End Sub

  Private Sub CheckLevelDecreasing(ByVal currentlevel As Integer, ByVal pumpPauseTime As Integer)
    'Setup a pump pausing timer to ensure that we're losing level as expected - if not, pause pump
    If _timerPump.Finished Then
      If _pumpPause Then
        'we were paused, so restart the add pump and count down from the pause time
        _pumpPause = False
        _levelLast = currentlevel
        _timerPump.TimeRemaining = 15
      Else
        'we we running the pump, so make sure level hasn't
        If Not (currentlevel < (_levelLast - 25)) Then
          _pumpPause = True
          _timerPump.TimeRemaining = MinMax(pumpPauseTime, 5, 10)
        Else
          _levelLast = currentlevel
          _timerPump.TimeRemaining = 15
        End If
      End If
    End If
  End Sub

  Private Sub CheckLevelMaintaining(ByVal currentlevel As Integer, ByVal pumpPauseTime As Integer)
    'Setup a pump pausing timer to ensure that we're not increasing in level due to pump not primed? <level increases while pump is running>
    If _timerPump.Finished Then
      If _pumpPause Then
        'we were paused, so restart the add pump and count down from the pause time
        _pumpPause = False
        _levelLast = currentlevel
        _timerPump.TimeRemaining = 15
      Else
        'we we running the pump, so make sure level hasn't increased
        If currentlevel > (_levelLast + 25) Then
          _pumpPause = True
          _timerPump.TimeRemaining = MinMax(pumpPauseTime, 5, 10)
        Else
          _levelLast = currentlevel
          _timerPump.TimeRemaining = 15
        End If
      End If
    End If
  End Sub

#End Region

#Region " I/O PROPERTIES "

  Public ReadOnly Property IoAddLamp As Boolean
    Get
      Return False
    End Get
  End Property

  ReadOnly Property IoFill As Boolean
    Get
      Return (State = EState.FillWithWater) OrElse (State = EState.DrainRinseToDrain) OrElse
              (State = EState.DrainFlushFill) OrElse (State = EState.DrainFlushRinseToDrain)
    End Get
  End Property

  ReadOnly Property IoFillCold As Boolean
    Get
      Return ((State = EState.FillWithWater) AndAlso (FillType = EFillType.Cold OrElse FillType = EFillType.Mix)) OrElse
              (State = EState.DrainRinseToDrain) OrElse
              (State = EState.DrainFlushFill) OrElse (State = EState.DrainFlushRinseToDrain)
    End Get
  End Property

  ReadOnly Property IoFillHot As Boolean
    Get
      Return ((State = EState.FillWithWater) AndAlso (FillType = EFillType.Hot OrElse FillType = EFillType.Mix))
    End Get
  End Property



  ReadOnly Property IoFillRunback As Boolean
    Get
      Return (State = EState.FillRunback)
    End Get
  End Property

  ReadOnly Property IoTransfer() As Boolean
    Get
      Return False
    End Get
  End Property

  ReadOnly Property IoAddPump As Boolean
    Get
      Return (State = EState.Mixing) OrElse (IoTransfer AndAlso Not _pumpPause) OrElse
             (State = EState.DrainFlushCirculate) OrElse (State = EState.DrainFlushCloseMixForTime)
    End Get
  End Property

  ReadOnly Property IoAddMix As Boolean
    Get
      Return (State = EState.Mixing) OrElse (State = EState.DrainFlushCirculate)
    End Get
  End Property

  ReadOnly Property IoDrain() As Boolean
    Get
      Return ((State >= EState.DrainEmpty1) AndAlso (State <= EState.DrainDrainEmpty2)) OrElse
             ((State >= EState.DrainFlushTransfer1) AndAlso (State <= EState.DrainFlushTransfer2))
    End Get
  End Property



#End Region




#If 0 Then
  ' TODO - From AEThenPlatform - Check
'===============================================================================================
'General purpose add tank prepare command
'===============================================================================================
Option Explicit
Public Enum AddPrepareState
  Off
  WaitIdle
  PreFill
  DispenseWaitTurn
  DispenseWaitReady
  DispenseWaitProducts
  DispenseWaitResponse
  Fill
  Heat
  Slow
  Fast
  MixForTime
  Ready
  KAInterlock
  KATransfer1
  KARinse
  KATransfer2
  KARinseToDrain
  KATransferToDrain
  InManual
End Enum

Public State As AddPrepareState
Public StatePrev As AddPrepareState
Public StateString As String

Public AddPreFillLevel As Long
Public AddFillLevel As Long
Public DesiredTemperature As Long
Public AddMixTime As Long
Public AddMixing As Boolean
Public OverrunTime As Long
Public OverrunTimer As New acTimer
Public MixTimer As New acTimer
Public HeatOn As Boolean, FillOn As Boolean
Public AddCallOff As Long
Public Timer As New acTimer
Public AdvanceTimer As New acTimer
Public HeatPrepTimer As New acTimer
Public DispenseTimer As New acTimer
Private NumberOfRinses As Long
Public WaitReadyTimer As New acTimerUp
 
'variables for alarms on manual dispense or error. does nothing
Public ManualAdd As Boolean
Public AddDispenseError As Boolean
Public AlarmRedyeIssue As Boolean

'For display only
Public AddRecipeStep As String
Private RecipeSteps(1 To 64, 1 To 8) As String

'For Display in drugroom preview
Public DrugroomDisplay As String

'Dispense States
Public DispenseCalloff As Long            'This filled in by us
Public DispenseTank As Long               'This filled in by us
Public Dispensestate As Long              'This filled in by AutoDispenser
Public Dispenseproducts As String         'This filled in by Autodispenser

Public DispenseDyesOnly As Boolean        'Used to determine dispenser delay type
Public DispenseChemsOnly As Boolean
Public DispenseDyesChems As Boolean
Private DispenseDyes As Boolean
Private DispenseChems As Boolean

'Dispense States
Private Const DispenseReady As Long = 101
Private Const DispenseBusy As Long = 102
Private Const DispenseAuto As Long = 201
Private Const DispenseScheduled As Long = 202
Private Const DispenseComplete As Long = 301
Private Const DispenseManual As Long = 302
Private Const DispenseError As Long = 309

Public Sub Start(FillLevel As Long, DesiredTemp As Long, MixTime As Long, MixingOn As Long, Calloff As Long, StandardTime As Long, DTank As Long)
  AddFillLevel = FillLevel
  DesiredTemperature = DesiredTemp
  If DesiredTemperature > 1800 Then DesiredTemperature = 1800
  AddMixTime = MixTime
  AddMixing = MixingOn
  AddCallOff = Calloff
  OverrunTime = StandardTime
  DispenseTank = DTank
  
  State = WaitIdle
  WaitReadyTimer.Pause
  Timer = 5
  AddDispenseError = False
  
  DispenseDyes = False
  DispenseChems = False
  DispenseDyesOnly = False
  DispenseChemsOnly = False
  DispenseDyesChems = False
  
End Sub

Public Sub Run(ByVal ControlObject As Object)
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode
    Dispensestate = .Dispensestate
    Dispenseproducts = .Dispenseproducts
    
    'Run an advance timer to reset alarms, where necessary
    If Not .IO_RemoteRun Then AdvanceTimer.TimeRemaining = 2
    If AdvanceTimer.Finished Then AlarmRedyeIssue = False
  
  Do
    'Remember state and loop until state does not change
    'This makes sure we go through all the state changes before proceeding and setting IO
     StatePrev = State
  
  Select Case State
  
    Case Off
      StateString = ""
      DrugroomDisplay = "Tank 1 Idle"
      
    Case WaitIdle    'Wait for destination tank to be idle - no active drains...
      DrugroomDisplay = "Wait For Destination Idle "
      If AlarmRedyeIssue And (.Parameters.EnableRedyeIssueAlarm = 1) Then
        Timer = 5
        StateString = "Check Tank (Redye Issue), Hold Run to reset " & TimerString(AdvanceTimer.TimeRemaining)
        DrugroomDisplay = "Check Tank (Redye Issue), Hold Run to Reset " & TimerString(AdvanceTimer.TimeRemaining)
      ElseIf (DispenseTank = 1) And (.AD.IsOn Or .ManualAddDrain.IsActive) Then
        Timer = 5
        StateString = "Wait to Add - " & .AD.StateString
      ElseIf (DispenseTank = 2) And (.RD.IsOn Or .ManualReserveDrain.IsActive) Then
        Timer = 5
        StateString = "Wait to Reserve - " & .RD.StateString
      Else
        StateString = "Wait for Destination Idle "
      End If
      If Timer.Finished Then State = PreFill

    Case PreFill
      StateString = "Tank 1 PreFill to " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters.Tank1PreFillLevel / 10, "0", 3) & "%"
      DrugroomDisplay = "Tank 1 PreFill to " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters.Tank1PreFillLevel / 10, "0", 3) & "%"
      If (.Tank1Level >= .Parameters.Tank1PreFillLevel) Then
         State = DispenseWaitTurn
      End If
      
    Case DispenseWaitTurn    'Wait for other tank to be finished with dispensing before new dispense
      StateString = "Wait for Turn "
      DrugroomDisplay = "Wait for Turn"
      If (AddCallOff = 0) Or (.Parameters.DispenseEnabled <> 1) Then
        HeatPrepTimer = .Parameters.Tank1HeatPrepTimer
        State = Fill
        ManualAdd = True
      Else
        DispenseTimer = .Parameters.DispenseReadyDelayTime * 60
        State = DispenseWaitReady
        DispenseCalloff = 0
      End If
      
    Case DispenseWaitReady    'Wait to make sure dispenser has completed previous dispense and is ready
      StateString = "Wait for Dispenser Ready "
      DrugroomDisplay = "Wait for Dispenser Ready "
      If Dispensestate = DispenseReady Then                                   '101
         State = DispenseWaitProducts
         DispenseTimer = .Parameters.DispenseResponseDelayTime * 60
         DispenseCalloff = AddCallOff
         .DispenseTank = DispenseTank
      End If
      If (AddCallOff = 0) Or (.Parameters.DispenseEnabled <> 1) Then
        HeatPrepTimer = .Parameters.Tank1HeatPrepTimer
        State = Fill
        DispenseCalloff = 0
        ManualAdd = True
      End If
          
    Case DispenseWaitProducts     'Wait for Dispensestate & DispenseProducts String to be set to determine how to signal delays
      StateString = "Wait for Dispense Products "
      DrugroomDisplay = "Wait for Dispense Products "
      DispenseTimer = .Parameters.DispenseResponseDelayTime * 60
      If Dispensestate <> DispenseReady Then                '(DispenseBusy = 102) but if these no recipe, we'll get a (DispenseManual = 302)
        If Dispenseproducts <> "" Then
          'split the products
          'products() = "Step <calloff>| <Ingredient_id> : <Amount> <Units> <DResult> | <Ingredient_Desc> "
          Dim ProductsArray() As String, i As Long
          ProductsArray = Split(Dispenseproducts, "|")
          For i = 1 To UBound(ProductsArray)                'Disregard the 0 row "Step <calloff>
            Dim position As Long
            position = InStr(ProductsArray(i), ":")         'Need to verify ":" is in the current row due to second split "|"
            If position > 0 Then
              Dim test As String
              test = Mid(ProductsArray(i), 6, 1)
              If test = ":" Then                            'If ":" is at 5th position, we have a chemical => "| 1004: ..."
                DispenseChems = True
              Else
                DispenseDyes = True
              End If
            End If
          Next i
          If DispenseChems And DispenseDyes Then
            DispenseDyesChems = True
          Else
            If DispenseChems Then
              DispenseChemsOnly = True
            Else: DispenseDyesOnly = True
            End If
          End If
        End If
        'Proceed to the next state
        State = DispenseWaitResponse
        DispenseCalloff = AddCallOff
        .DispenseTank = DispenseTank
      End If
      If (AddCallOff = 0) Or (.Parameters.DispenseEnabled <> 1) Then
        HeatPrepTimer = .Parameters.Tank1HeatPrepTimer
        State = Fill
        DispenseCalloff = 0
        ManualAdd = True
      End If
      
    Case DispenseWaitResponse     'Wait for response from dispenser
      StateString = "Wait For Reponse From Dispenser "
      DrugroomDisplay = "Wait for Response From Dispenser "
      AddRecipeStep = Dispenseproducts
      Select Case Dispensestate
         Case DispenseComplete                                                '301
          If Not .WK.IsOn Then
            If DispenseTank = 2 Then
             .ReserveReady = True
            ElseIf DispenseTank = 1 Then
             .AddReady = True
            End If
          End If
          .KR.ACCommand_Cancel
          If .LA.KP1.IsOn Then .LAActive = True
          State = Off
          Cancel
          .DispenseTank = 0
          DispenseTank = 0
          DispenseCalloff = 0

        Case DispenseManual                                                  '302
          HeatPrepTimer = .Parameters.Tank1HeatPrepTimer
          State = Fill
          DispenseCalloff = 0
          ManualAdd = True

        Case DispenseError                                                   '309
          AddDispenseError = True
          HeatPrepTimer = .Parameters.Tank1HeatPrepTimer
          State = Fill
          DispenseCalloff = 0

      End Select
     
    Case Fill
      StateString = "Tank 1 Fill to " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(AddFillLevel / 10, "0", 3) & "%"
      DrugroomDisplay = "Fill to " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(AddFillLevel / 10, "0", 3) & "%"
      If .IO_Tank1Manual_SW Then State = InManual
      If AddFillLevel = 0 Then
        WaitReadyTimer.Start
        State = Slow
      End If
      If .Tank1Level > AddFillLevel Then State = Heat
      
    Case Heat
      If Not .DrugroomMixerOn Then
        StateString = "Tank 1 Level Too Low to Heat " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters.Tank1MixerOnLevel / 10, "0", 3) & "%"
        DrugroomDisplay = "Level Too Low to Heat " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters.Tank1MixerOnLevel / 10, "0", 3) & "%"
      Else
        StateString = "Tank 1 Heat to " & Pad(.IO_Tank1Temp / 10, "0", 3) & "F / " & Pad(DesiredTemperature / 10, "0", 3)
        DrugroomDisplay = "Heat to " & Pad(.IO_Tank1Temp / 10, "0", 3) & "F / " & Pad(DesiredTemperature / 10, "0", 3)
      End If
      If .IO_Tank1Manual_SW Then State = InManual
      If AddFillLevel = 0 Then State = Slow
      ' Heat Start/Stop
      If .IO_Tank1Temp < DesiredTemperature Then
        HeatOn = True
        AddMixing = True ' Must mix if heating
      End If
      If .IO_Tank1Temp > DesiredTemperature Then
        HeatOn = False
        State = Slow
        WaitReadyTimer.Start
        OverrunTimer = OverrunTime
      End If
      
    Case Slow
      StateString = "Waiting on Tank 1 " & TimerString(WaitReadyTimer.TimeElapsed)
      DrugroomDisplay = "Prepare Tank 1"
      If .IO_Tank1Manual_SW Then State = InManual
      ' Heat Start/Stop
      If .IO_Tank1Temp < DesiredTemperature Then
        HeatOn = True
        AddMixing = True ' Must mix if heating
      End If
      If .IO_Tank1Temp >= DesiredTemperature Then HeatOn = False
      If DispenseTank = 2 Then
        'Reserve Tank
        If .RT.IsWaitReady Then State = Fast
      ElseIf DispenseTank = 1 Then
        'Add Tank
        If .AC.IsWaitReady Then State = Fast
        If .AT.IsWaitReady Then State = Fast
      End If
      If .Tank1Ready Then
        WaitReadyTimer.Pause
        State = MixForTime
        MixTimer = AddMixTime
      End If
      If WaitReadyTimer.IsPaused Then WaitReadyTimer.Restart 'should cover LA copy to KP function
      If OverrunTimer.Finished Then
         State = Fast
      End If
    
    Case Fast
      StateString = "Waiting on Tank 1 " & TimerString(WaitReadyTimer.TimeElapsed)
      DrugroomDisplay = "Prepare Tank 1"
      If .IO_Tank1Manual_SW Then State = InManual
      ' Heat Start/Stop
      If .IO_Tank1Temp < DesiredTemperature Then
        HeatOn = True
        AddMixing = True ' Must mix if heating
      End If
      If .IO_Tank1Temp >= DesiredTemperature Then HeatOn = False
      If WaitReadyTimer.IsPaused Then WaitReadyTimer.Restart 'should cover LA copy to KP function
      If .Tank1Ready Then
        If .Tank1Level Then
          WaitReadyTimer.Pause
          State = MixForTime
          MixTimer = AddMixTime
        Else
          StateString = "Tank 1 Level low " & Pad(.Tank1Level / 10, "0", 3) & "%"
        End If
      End If

      
    Case MixForTime
      If Not .DrugroomMixerOn Then
        StateString = "Tank 1 Level Too Low to Mix " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters.Tank1MixerOnLevel / 10, "0", 3) & "%"
        DrugroomDisplay = "Level Too Low to Mix " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters.Tank1MixerOnLevel / 10, "0", 3) & "%"
      Else
        StateString = "Tank 1 Mix For Time " & TimerString(MixTimer.TimeRemaining)
        DrugroomDisplay = "Mix For Time " & TimerString(MixTimer.TimeRemaining)
      End If
      If .IO_Tank1Manual_SW Then State = InManual
      If AddFillLevel = 0 Then State = Ready
      ' Heat Start/Stop
      If .IO_Tank1Temp < DesiredTemperature Then
        HeatOn = True
        AddMixing = True ' Must mix if heating
      End If
      If .IO_Tank1Temp >= DesiredTemperature Then HeatOn = False
      If MixTimer.Finished Then State = Ready
      If Not .Tank1Ready Then
        State = Fast
        WaitReadyTimer.Restart
      End If
      
    Case Ready
      StateString = "Tank 1 Ready "
      DrugroomDisplay = "Ready "
      If .IO_Tank1Manual_SW Then State = InManual
      If (DispenseTank = 1) And (.AdditionLevel > .Parameters.AdditionMaxTransferLevel) Then
         State = KAInterlock
      Else
        State = KATransfer1
      End If
      FillOn = False
      HeatOn = False
      AddMixing = False
      AddDispenseError = False
      ManualAdd = False

    Case KAInterlock
      StateString = "Add Level High for Tank 1 Transfer " & Pad(.AdditionLevel / 10, "0", 3) & "%"
      DrugroomDisplay = "Add Level High for Tank 1 Transfer " & Pad(.AdditionLevel / 10, "0", 3) & "%"
      If .IO_Tank1Manual_SW Then State = InManual
      If (.AdditionLevel <= .Parameters.AdditionMaxTransferLevel) Then
        State = KATransfer1
      End If
    
    Case KATransfer1
      If (.Tank1Level > 10) Then
        Timer = .Parameters.Tank1TimeBeforeRinse
        StateString = "Tank 1 Transfer Empty " & Pad(.Tank1Level / 10, "0", 3) & "%"
        DrugroomDisplay = "Transfer Empty " & Pad(.Tank1Level / 10, "0", 3) & "%"
      Else
        StateString = "Tank 1 Transfer Empty " & TimerString(Timer.TimeRemaining)
        DrugroomDisplay = "Transfer Empty " & TimerString(Timer.TimeRemaining)
      End If
      If .IO_Tank1Manual_SW Then State = InManual
      If Timer.Finished Then
        NumberOfRinses = .Parameters.DrugroomRinses
        If .KR.IsOn Then
          If (.KR.RinseMachine = 0) And (.KR.RinseDrain = 0) Then
            If Not .WK.IsOn Then
              If (DispenseTank = 1) Then .AddReady = True
              If (DispenseTank = 2) Then .ReserveReady = True
            End If
            If .LA.KP1.IsOn Then .LAActive = True
            DispenseTank = 0
            .DispenseTank = 0
            .Tank1Ready = False
            .KR.ACCommand_Cancel
            Cancel
          ElseIf .KR.RinseMachine = 0 Then
            State = KARinseToDrain
            Timer = .Parameters.Tank1RinseToDrainTime
          End If
        Else
          State = KARinse
          Timer = .Parameters.Tank1RinseTime
        End If
      End If
            
    Case KARinse
      StateString = "Tank 1 Rinse " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters.Tank1RinseLevel / 10, "0", 3) & "%"
      DrugroomDisplay = "Rinse " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters.Tank1RinseLevel / 10, "0", 3) & "%"
      If .IO_Tank1Manual_SW Then State = InManual
      If (.Tank1Level > .Parameters.Tank1RinseLevel) Then
        State = KATransfer2
        Timer = .Parameters.Tank1TimeAfterRinse
      End If
      
    Case KATransfer2
      If (.Tank1Level > 10) Then
        Timer = .Parameters.Tank1TimeAfterRinse
        StateString = "Tank 1 Transfer Rinse " & Pad(.Tank1Level / 10, "0", 3) & "%"
        DrugroomDisplay = "Transfer Rinse " & Pad(.Tank1Level / 10, "0", 3) & "%"
      Else
        StateString = "Tank 1 Transfer Rinse " & TimerString(Timer.TimeRemaining)
        DrugroomDisplay = "Transfer Rinse " & TimerString(Timer.TimeRemaining)
      End If
      If .IO_Tank1Manual_SW Then State = InManual
      If Timer.Finished Then
        If Not .WK.IsOn Then
          If (DispenseTank = 1) Then .AddReady = True
          If (DispenseTank = 2) Then .ReserveReady = True
        End If
        If .LA.KP1.IsOn Then .LAActive = True
        DispenseTank = 0
        .DispenseTank = 0
        .Tank1Ready = False
        State = KARinseToDrain
        Timer = .Parameters.Tank1RinseToDrainTime
      End If
      
    Case KARinseToDrain
      StateString = "Tank 1 Rinse To Drain " & TimerString(Timer.TimeRemaining)
      DrugroomDisplay = "Rinse To Drain " & TimerString(Timer.TimeRemaining)
      If .IO_Tank1Manual_SW Then State = InManual
      If Timer.Finished Or (.Tank1Level >= 500) Then
        State = KATransferToDrain
        Timer = .Parameters.Tank1DrainTime
        NumberOfRinses = NumberOfRinses - 1
      End If
    
    Case KATransferToDrain
      If (.Tank1Level > 10) Then
        Timer = .Parameters.Tank1DrainTime
        StateString = "Tank 1 To Drain " & Pad(.Tank1Level / 10, "0", 3) & "%"
        DrugroomDisplay = "Transfer To Drain " & Pad(.Tank1Level / 10, "0", 3) & "%"
      Else
        StateString = "Tank 1 To Drain " & TimerString(Timer.TimeRemaining)
        DrugroomDisplay = "Transfer To Drain " & TimerString(Timer.TimeRemaining)
      End If
      If .IO_Tank1Manual_SW Then State = InManual
      If Timer.Finished Then
        If NumberOfRinses > 0 Then
          State = KARinseToDrain
          Timer = .Parameters.Tank1RinseToDrainTime
        Else
          If Not .WK.IsOn Then
            If (DispenseTank = 1) Then .AddReady = True
            If (DispenseTank = 2) Then .ReserveReady = True
          End If
          If .LA.KP1.IsOn Then .LAActive = True
          DispenseTank = 0
          .DispenseTank = 0
          .Tank1Ready = False
          .KR.ACCommand_Cancel
          Cancel
        End If
      End If
         
    Case InManual
      StateString = "Drugroom Switch In Manual "
      DrugroomDisplay = "Switch In Manual"
      .Tank1Ready = False
      HeatPrepTimer = .Parameters.Tank1HeatPrepTimer
      If Not .IO_Tank1Manual_SW Then
        HeatPrepTimer = .Parameters.Tank1HeatPrepTimer
        State = Fill
      End If
    
  End Select
  
  Loop Until (StatePrev = State)    'Loop until state does not change
  
  End With
End Sub

Public Sub Cancel()
  FillOn = False
  State = Off
  HeatOn = False
  AddMixing = False
  AddDispenseError = False
  WaitReadyTimer.Pause
  AlarmRedyeIssue = False
  ManualAdd = False

  'This is to clear prepare properties and hopefully resolve issues in LA where wrong calloff is used
  AddCallOff = 0
  AddFillLevel = 0
  AddMixTime = 0
  AddRecipeStep = ""
  DesiredTemperature = 0
  DispenseCalloff = 0
  DispenseTank = 0
  Dispenseproducts = ""
  Dispensestate = 0
        
  DispenseDyes = False
  DispenseChems = False
  DispenseDyesOnly = False
  DispenseChemsOnly = False
  DispenseDyesChems = False
  
End Sub

Friend Sub ProgramStart()
On Error Resume Next
  Dispenseproducts = ""
  DispenseCalloff = 0
 ' DispenseTank = 0
  AddCallOff = 0
  AddRecipeStep = ""
  Dim i1 As Long, i2 As Long
  For i1 = 1 To 64
    For i2 = 1 To 8
      RecipeSteps(i1, i2) = ""
    Next i2
  Next i1
End Sub

Public Sub CopyTo(Target As acAddPrepare)
  With Target
    .State = State
    .AddFillLevel = AddFillLevel
    .AddCallOff = AddCallOff
    .OverrunTimer.TimeRemaining = OverrunTimer.TimeRemaining
    .AddMixing = AddMixing
    .DesiredTemperature = DesiredTemperature
    .AddMixTime = AddMixTime
    .OverrunTime = OverrunTime
    .MixTimer.TimeRemaining = MixTimer.TimeRemaining
    .WaitReadyTimer.TimeElapsed = WaitReadyTimer.TimeElapsed
        
    .ManualAdd = ManualAdd
    .AlarmRedyeIssue = AlarmRedyeIssue
    .AddDispenseError = AddDispenseError
    .AddRecipeStep = AddRecipeStep
    .DispenseCalloff = DispenseCalloff
    .DispenseTank = DispenseTank
    .Dispensestate = Dispensestate
    .Dispenseproducts = Dispenseproducts
    .Timer.TimeRemaining = Timer.TimeRemaining
    
    .DrugroomDisplay = DrugroomDisplay
    .DispenseDyesChems = DispenseDyesChems
    .DispenseDyesOnly = DispenseDyesOnly
    .DispenseChemsOnly = DispenseChemsOnly
    .DispenseTimer.TimeRemaining = DispenseTimer.TimeRemaining
    .MixTimer.TimeRemaining = MixTimer.TimeRemaining
        
  End With
End Sub

Public Property Get IsOn() As Boolean
  IsOn = (State <> Off)
End Property
Public Property Get IsWaitingToDispense() As Boolean
  IsWaitingToDispense = (State = DispenseWaitTurn)
End Property
Public Property Get IsDispensing() As Boolean
  IsDispensing = (State = DispenseWaitReady) Or (State = DispenseWaitProducts) Or (State = DispenseWaitResponse)
End Property
Friend Property Get IsDispenseWaitReady() As Boolean
  IsDispenseWaitReady = (State = DispenseWaitReady)
End Property
Friend Property Get IsDispenseReadyOverrun() As Boolean
  IsDispenseReadyOverrun = IsDispenseWaitReady And DispenseTimer.Finished
End Property
Friend Property Get IsDispenseWaitResponse() As Boolean
  IsDispenseWaitResponse = (State = DispenseWaitResponse)
End Property
Friend Property Get IsDispenseResponseOverrun() As Boolean
  IsDispenseResponseOverrun = IsDispenseWaitResponse And DispenseTimer.Finished
End Property

Public Property Get IsFill() As Boolean
  IsFill = ((State = Fill) Or FillOn) And (Not State = InManual)
End Property
Public Property Get IsHeating() As Boolean
  IsHeating = (HeatOn And ((State = Heat) Or (State = Slow) Or (State = Fast))) And _
                (Not State = InManual)
End Property
Friend Property Get IsHeatTankOverrun() As Boolean
  IsHeatTankOverrun = ((State = Fill) Or (State = Heat)) And HeatPrepTimer.Finished
End Property
Public Property Get IsSlow() As Boolean
  IsSlow = (State = Slow)
End Property
Public Property Get IsFast() As Boolean
  IsFast = (State = Fast)
End Property
Friend Property Get IsWaitReady() As Boolean
  IsWaitReady = (State = Slow) Or (State = Fast) Or (State = MixForTime)
End Property
Public Property Get IsMixing() As Boolean
  IsMixing = (State = MixForTime)
End Property
Public Property Get IsReady() As Boolean
  IsReady = (State = Ready)
End Property
Public Property Get IsMixerOn() As Boolean
  IsMixerOn = AddMixing And Not ((State = InManual) Or IsDispensing Or IsWaitingToDispense Or IsWaitIdle)
End Property
Public Property Get IsOverrun() As Boolean
  IsOverrun = IsWaitReady And OverrunTimer.Finished
End Property
Friend Property Get IsInterlocked() As Boolean
 IsInterlocked = (State = KAInterlock)
End Property
Friend Property Get IsTransfer() As Boolean
  IsTransfer = (State = KATransfer1) Or (State = KATransfer2) Or (State = KARinseToDrain) Or (State = KATransferToDrain)
End Property
Friend Property Get IsTransferToAddition() As Boolean
 IsTransferToAddition = ((State = KATransfer1) Or (State = KARinse) Or _
                       (State = KATransfer2)) And (DispenseTank = 1)
End Property
Friend Property Get IsTransferToReserve() As Boolean
 IsTransferToReserve = ((State = KATransfer1) Or (State = KARinse) Or _
                       (State = KATransfer2)) And (DispenseTank = 2)
End Property
Friend Property Get IsTransferToDrain() As Boolean
 IsTransferToDrain = IsRinseToDrain Or (State = KATransferToDrain)
End Property
Friend Property Get IsRinse() As Boolean
  If (State = KARinse) Then IsRinse = True
End Property
Friend Property Get IsRinseToDrain() As Boolean
  If (State = KARinseToDrain) Then IsRinseToDrain = True
End Property
Friend Property Get IsPaused() As Boolean
  IsPaused = (State = KAPause)
End Property
Friend Property Get IsWaitIdle() As Boolean
  IsWaitIdle = (State = WaitIdle)
End Property
Friend Property Get IsDelayed() As Boolean
  IsDelayed = (IsWaitReady And IsOverrun)
End Property
Friend Property Get IsInManual() As Boolean
  IsInManual = (State = InManual)
End Property

#End If


End Class

