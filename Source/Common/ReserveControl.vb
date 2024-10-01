
Imports Utilities.Translations

Public Class ReserveControl
  Inherits MarshalByRefObject

  ' For convenience
  Private Property Parent As ACParent
#Disable Warning IDE1006 ' Naming Styles
  Private Property controlCode As ControlCode
#Enable Warning IDE1006 ' Naming Styles


  Private _stateTimer As New Timer
  Private _alarmTimer As New Timer
  ' Public Timer Value Display
  Public StateTimerSecs As Integer = _stateTimer.Seconds
  Public AlarmTimerSecs As Integer = _alarmTimer.Seconds

  Public FillType As EFillType
  Public FillLevel As Integer
  Public TempDesired As Integer
  Public HeatOn As Boolean

  Public Enum EStateMixer
    Off
    [On]
  End Enum
  Public Property StateMixer As EStateMixer

  Public Enum EState
    Off
    Idle

    Pause

    FillStart
    FillInterlock
    FillDepressurize
    FillFromMachine
    FillWithWater

    HeatStart
    HeatToSetpoint
    HeatMaintain

    TransferStart
    TransferInterlock
    TransferEmpty1
    TransferRinse
    TransferEmpty2

    DrainStart
    DrainEmpty1
    DrainFillForMix
    DrainHeatToMix
    DrainEmpty2
    DrainRinseTime
    DrainEmpty3
    DrainComplete
  End Enum
  Public State As EState
  Public RestartState As EState
  Public Status As String

  Public LevelActual As Integer

  Public Timer As New Timer
  Public RestartTime As Integer

  Public Sub New(ByVal parent As ACParent, ByVal controlCode As ControlCode)
    Me.Parent = parent
    Me.controlCode = controlCode
    Me.Timer.Seconds = 10
  End Sub



#Region " STATE MACHINE "

  Public Sub Run()
    Try
      With controlCode

        ' Update local Reserve Level reference
        LevelActual = .ReserveLevel

        ' Manage Reserve Mixer
        If LevelActual > .Parameters.ReserveMixerOnLevel Then StateMixer = EStateMixer.On
        If LevelActual <= .Parameters.ReserveMixerOffLevel Then StateMixer = EStateMixer.Off

        ' Pause the state machine under these conditions
        Dim pauseFillDrain As Boolean = .Parent.IsPaused OrElse .EStop
        Dim pauseTransfer As Boolean = pauseFillDrain OrElse (Not .TempSafe) OrElse (Not .MachineClosed)

        ' Remember the start state so we can loop until all state changes have completed
        '   prevents IO flicker as we move through the state machine
        Static startState As EState
        Do
          ' Set default restart state & time
          If (Not IsPaused) Then
            RestartState = State
            RestartTime = 5
          End If

          'Save start state so we can see if there has been a state change at the end
          '  normally we keep looping until there are no state changes
          startState = State
          Select Case State
            Case EState.Off
              Status = Translate("Startup") & Timer.ToString(1)
              If Timer.Finished Then
                State = EState.Idle
              End If

            Case EState.Idle
              Status = Translate("Idle")

            Case EState.Pause
              Status = Translate("Paused")
              If .EStop Then
                Status = Status & Translate("EStop")
              Else
                If Not .PumpControl.PumpRunning Then Status = Translate("Paused") & (", ") & Translate("Pump Off")
                If Not .TempSafe Then Status = Translate("Paused") & (", ") & Translate("Temp Not Safe")
                If Not .MachineClosed Then Status = Translate("Paused") & (", ") & Translate("Machine Not Closed")
                '   If Not .SampleClosed Then StateString = Translate("Paused") & (", ") & Translate("Sample Not Closed")
              End If
              If Not pauseTransfer Then State = EState.Idle


              '====================================== FILL ===========================================

            Case EState.FillStart
              Status = Translate("Fill") & (" ") & Translate("Starting") & Timer.ToString(1)
              If Timer.Finished Then
                State = EState.FillWithWater
              End If

            Case EState.FillInterlock
              Status = Translate("Wait Idle") & Timer.ToString(1)
              If .KA.IsOn AndAlso (.KA.Destination = EKitchenDestination.Reserve) Then Status = .KA.Status : Timer.Seconds = 2
              If .KP.IsOn AndAlso (.KP.KP1.DispenseTank = 2) Then Status = .KP.KP1.DrugroomDisplay : Timer.Seconds = 2
              If .KP.IsOn AndAlso (.KP.KP1.Param_Destination = EKitchenDestination.Reserve) AndAlso (.KP.KP1.IsForeground) Then Status = .KP.KP1.DrugroomDisplay : Timer.Seconds = 2
              If .LA.IsOn AndAlso (.LA.KP1.Param_Destination = EKitchenDestination.Reserve) AndAlso (.LA.KP1.IsForeground) Then Status = .LA.KP1.DrugroomDisplay : Timer.Seconds = 2
              If Not .TempSafe Then Status = Translate("Temp Not Safe") & Timer.ToString(1) : Timer.Seconds = 2
              If .TempSafe AndAlso Timer.Finished Then
                If FillType = EFillType.Vessel Then
                  ' Check Machine Safe conditions
                  If .TempSafe AndAlso .MachineClosed Then
                    .CO.Cancel() : .HE.Cancel() : .TP.Cancel()
                    .TemperatureControl.Cancel()
                    .AirpadOn = False
                    .PR.Cancel()
                    .PumpControl.StopMainPump()
                    State = EState.FillDepressurize
                    Timer.Seconds = 2
                  End If
                Else
                  State = EState.FillWithWater
                End If
              End If

            Case EState.FillDepressurize
              Status = Translate("Depressuzing") & Timer.ToString(1)
              If Not (.PressSafe) Then Status = Translate("Depressuzing") & (" ") & .MachinePressureDisplay : Timer.Seconds = 2
              If .PressSafe AndAlso Timer.Finished Then
                State = EState.FillFromMachine
                .GetRecordedLevel = False
              End If
              If Not (.TempSafe AndAlso .MachineClosed) Then State = EState.FillInterlock

            Case EState.FillFromMachine
              ' Deadband parameter to assist with possible overshoot
              Status = Translate("Filling from machine") & (" ") & (LevelActual / 10).ToString("#0.0") & " / " & (FillLevel / 10).ToString("#0.0") & "%"
              If LevelActual >= FillLevel Then
                State = EState.Idle
              End If

            Case EState.FillWithWater
              ' Deadband parameter to assist with possible overshoot
              Status = Translate("Filling") & (" ") & (LevelActual / 10).ToString("#0.0") & " / " & (FillLevel / 10).ToString("#0.0") & "%"
              If LevelActual >= FillLevel Then
                State = EState.Idle
              End If

              '====================================== MIXING ========================================
              ' No Mixer Installed
              '   Case EState.Mixing
              '    StateString = Translate("Mixing")


              '====================================== HEATING ========================================
            Case EState.HeatStart
              If TempDesired > 0 AndAlso (LevelActual > 100) Then ' TODO HeatEnable Level
                If Timer.Finished Then
                  State = EState.HeatToSetpoint
                End If
                Status = Translate("Heat Start") & Timer.ToString(1)
              Else : Status = Translate("Level Too Low to Heat")
              End If

            Case EState.HeatToSetpoint
              If TempDesired > 0 Then
                State = EState.HeatMaintain
              End If
              Status = Translate("Heating") & (" ") & (.IO.ReserveTemp / 10).ToString("#0.0") & " / " & (TempDesired / 10).ToString("#0.0") & " F"

            Case EState.HeatMaintain
              If (.IO.ReserveTemp < (TempDesired - .Parameters.ReserveHeatDeadband)) Then HeatOn = True
              If (.IO.ReserveTemp >= (TempDesired - .Parameters.ReserveHeatDeadband)) Then HeatOn = False
              Status = Translate("Heating") & (" ") & (.IO.ReserveTemp / 10).ToString("#0.0") & " / " & (TempDesired / 10).ToString("#0.0") & " F"



              '===================================== TRANSFER =======================================

            Case EState.TransferStart
              Status = Translate("Transfer") & (" ") & Translate("Starting") & Timer.ToString(1)
              If pauseTransfer Then
                If Not .MachineClosed Then Status = Translate("Transfer") & (" ") & Translate("Machine Not Closed")
                If Not .TempSafe Then Status = Translate("Transfer") & (" ") & Translate("Temp Not Safe")
                If .EStop Then Status = "ESTOP"
              Else
                If Timer.Finished Then
                  .PumpControl.StartAuto()
                  State = EState.TransferEmpty1
                End If
              End If

            Case EState.TransferEmpty1
              Status = Translate("Transferring")
              If .ReserveLevel > 10 Then
                Status = Status & (" ") & (.ReserveLevel / 10).ToString("#0.0") & "%"
                Timer.Seconds = .Parameters.AddTransferTimeBeforeRinse
              Else
                Status = Status & Timer.ToString(1)
              End If
              If Timer.Finished Then
                State = EState.TransferRinse
                Timer.Seconds = .Parameters.ReserveRinseTime
              End If

            Case EState.TransferRinse
              Status = Translate("Rinse to machine") & Timer.ToString(1)
              If Timer.Finished Then
                State = EState.TransferEmpty2
                Timer.Seconds = .Parameters.ReserveTimeAfterRinse
              End If

            Case EState.TransferEmpty2
              Status = Translate("Transferring")
              If .ReserveLevel > 10 Then
                Status = Status & (" ") & (.ReserveLevel / 10).ToString("#0.0") & "% " & Timer.ToString
                Timer.Seconds = .Parameters.ReserveRinseToDrainTime
              Else
                Status = Status & Timer.ToString(1)
              End If
              If Timer.Finished Then
                State = EState.DrainFillForMix
                Timer.Seconds = .Parameters.ReserveRinseToDrainTime
              End If


              '===================================== DRAIN =======================================

            Case EState.DrainStart
              Status = Translate("Drain Start") & Timer.ToString(1)
              If Timer.Finished Then
                State = EState.DrainEmpty1
                Timer.Seconds = .Parameters.ReserveTimeBeforeRinse
              End If

            Case EState.DrainEmpty1
              Status = Translate("Draining")
              If .ReserveLevel > 10 Then
                Status = Status & (" ") & (.ReserveLevel / 10).ToString("#0.0") & "%"
                Timer.Seconds = .Parameters.ReserveTimeBeforeRinse
              Else
                Status = Status & Timer.ToString(1)
              End If
              If Timer.Finished Then
                State = EState.DrainFillForMix
                Timer.Seconds = .Parameters.ReserveRinseToDrainTime
              End If

            Case EState.DrainFillForMix
              ' Always fill above the Heat Enable parameter, live steam in the reserve tank
              Dim reserveFillLevel As Integer = MinMax(.Parameters.ReserveRinseLevel, 450, 750)
              If reserveFillLevel < .Parameters.ReserveHeatEnableLevel Then reserveFillLevel = (.Parameters.ReserveHeatEnableLevel + 50)
              Status = Translate("Filling") & (" ") & (.ReserveLevel / 10).ToString("#0.0") & (" / ") & (reserveFillLevel / 10).ToString("#0.0") & ("%")
              If .ReserveLevel > reserveFillLevel Then
                Timer.Seconds = .Parameters.ReserveHeatTime
                State = EState.DrainHeatToMix
              End If

            Case EState.DrainHeatToMix
              Dim reserveTempLimit As Integer = MinMax(.Parameters.ReserveTankHighTempLimit, 0, 1300)
              Status = Translate("Heating") & (" ") & (.IO.ReserveTemp / 10).ToString("#0.0") & "F " & Timer.ToString
              If Timer.Finished OrElse (.IO.ReserveTemp >= reserveTempLimit) Then
                State = EState.DrainEmpty2
                Timer.Seconds = .Parameters.ReserveDrainTime
              End If

            Case EState.DrainEmpty2
              Status = Translate("Draining")
              If .ReserveLevel > 10 Then
                Status = Status & (" ") & (.ReserveLevel / 10).ToString("#0.0") & "%"
                Timer.Seconds = .Parameters.ReserveDrainTime
              Else
                Status = Status & Timer.ToString(1)
              End If
              If Timer.Finished Then
                State = EState.DrainRinseTime
                Timer.Seconds = MinMax(.Parameters.ReserveRinseToDrainTime, 10, 300)
              End If

            Case EState.DrainRinseTime
              Status = Translate("Rinsing") & Timer.ToString(1)
              If Timer.Finished Then
                State = EState.DrainEmpty3
                Timer.Seconds = .Parameters.ReserveDrainTime
              End If

            Case EState.DrainEmpty3
              Status = Translate("Draining")
              If .ReserveLevel > 10 Then
                Status = Status & (" ") & (.ReserveLevel / 10).ToString("#0.0") & "% " & Timer.ToString
                Timer.Seconds = .Parameters.ReserveDrainTime
              Else
                Status = Status & Timer.ToString(1)
              End If
              If Timer.Finished Then State = EState.Idle

          End Select
        Loop Until (State = startState)

      End With

    Catch ex As Exception
      Dim message As String = ex.Message
    End Try
  End Sub

  Public Sub Cancel()
    State = EState.Idle
  End Sub

  Public Sub ManualFill(ByVal fillType As EFillType, ByVal level As Integer)
    With controlCode

      Me.FillType = fillType
      Me.FillLevel = MinMax(level, 0, 1000)

      ' NOTE: During first onsite startups - Fill Hot valves were not installed as Heat Reclaim system is not yet onsite.  
      '       Use Setting to determine when to allow filling with hot water
      '     If (.Parameters.FillEnableHot = 0) AndAlso fillType = EFillType.Hot Then fillType = EFillType.Cold

      ' Set default state
      State = EState.FillStart
      Timer.Seconds = 2

    End With
  End Sub

  Public Sub ManualHeat(ByVal setpoint As Integer)
    With controlCode

      Me.TempDesired = MinMax(setpoint, 0, 1350)

      ' Set default state
      State = EState.HeatStart
      Timer.Seconds = 2

    End With
  End Sub

  Public Sub ManualDrain()
    State = EState.DrainEmpty2
    Timer.Seconds = MinMax(controlCode.Parameters.ReserveDrainTime, 10, 300)
  End Sub

  Public Sub ManualTransfer()
    State = EState.TransferStart
    Timer.Seconds = 2
  End Sub

#End Region

#Region " PROPERTIES "

  Public ReadOnly Property IsReady() As Boolean
    Get
      Return (State = EState.Idle)
    End Get
  End Property

  Public ReadOnly Property IsPaused() As Boolean
    Get
      Return (State = EState.Pause)
    End Get
  End Property

  Public ReadOnly Property IsActive As Boolean
    Get
      Return IsFillActive OrElse IsTransferActive OrElse IsDrainActive
    End Get
  End Property

  Public ReadOnly Property IsFillActive As Boolean
    Get
      Return (State >= EState.FillStart) AndAlso (State <= EState.FillWithWater)
    End Get
  End Property

  Public ReadOnly Property IsHeatActive As Boolean
    Get
      Return (State >= EState.HeatStart) AndAlso (State <= EState.HeatMaintain)
    End Get
  End Property

  Public ReadOnly Property IsTransferActive As Boolean
    Get
      Return (State = EState.TransferStart) OrElse (State = EState.TransferEmpty1) OrElse
             (State = EState.TransferRinse) OrElse (State = EState.TransferEmpty2)
    End Get
  End Property

  Public ReadOnly Property IsDrainActive As Boolean
    Get
      Return (State >= EState.DrainStart) AndAlso (State <= EState.DrainEmpty3)
    End Get
  End Property

#End Region


#Region " I/O PROPERTIES "

  Public ReadOnly Property IoReserveFillCold As Boolean
    Get
      Return ((State = EState.FillWithWater) AndAlso ((FillType = EFillType.Cold) OrElse (FillType = EFillType.Mix))) OrElse
              State = EState.TransferRinse OrElse State = EState.DrainFillForMix OrElse State = EState.DrainRinseTime
    End Get
  End Property

  Public ReadOnly Property IoReserveFillHot As Boolean
    Get
      Return ((State = EState.FillWithWater) AndAlso ((FillType = EFillType.Hot) OrElse (FillType = EFillType.Mix)))
    End Get
  End Property

  Public ReadOnly Property IoReserveHeat As Boolean
    Get
      Return (State = EState.HeatToSetpoint) OrElse ((State = EState.HeatMaintain) AndAlso HeatOn) OrElse _
             (State = EState.DrainHeatToMix)
    End Get
  End Property

  Public ReadOnly Property IoReserveTransfer() As Boolean
    Get
      Return (State = EState.TransferEmpty1) OrElse (State = EState.TransferEmpty2)
    End Get
  End Property

  Public ReadOnly Property IoReserveDrain() As Boolean
    Get
      Return (State = EState.DrainEmpty1) OrElse (State = EState.DrainEmpty2) OrElse (State = EState.DrainRinseTime) OrElse (State = EState.DrainEmpty3)
    End Get
  End Property

  Public ReadOnly Property IoReserveMixerOn As Boolean
    Get
      Return False
    End Get
  End Property

#End Region

#If 0 Then
  ' TODO vb6 remove

'===============================================================================================
'manual fill from add buttons
'===============================================================================================
'
  Option Explicit
  Public Enum ManualReserveTransferState
    Off
    Interlock
    Transfer1
    Rinse
    Transfer2
    RinseToDrain
    Drain
    TurnOff
  End Enum
  Public State As ManualReserveTransferState
  Public Timer As New acTimer
  

Public Sub Run(ByVal ControlObject As Object)
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode:


    Select Case State
       
       Case Off
          If .ManualReserveTransferPB Then
            State = Interlock
          End If
          
       Case Interlock
          If Not .ManualReserveTransferPB Then State = Off
          If .MachineSafe And .LidLocked Then
            State = Transfer1
            Timer = .Parameters.ReserveTimeBeforeRinse
          End If
          
        Case Transfer1
          If Not .ManualReserveTransferPB Then State = Off
          If Not .LidLocked Then State = Interlock
          If .ReserveLevel > 10 Then Timer = .Parameters.ReserveTimeBeforeRinse
          If Timer.Finished Then
             State = Rinse
             Timer = .Parameters.ReserveRinseTime
          End If
          
        Case Rinse
          If Not .ManualReserveTransferPB Then State = Off
          If Timer.Finished Then
             State = Transfer2
             Timer = .Parameters.ReserveTimeAfterRinse
          End If
          
        Case Transfer2
          If Not .ManualReserveTransferPB Then State = Off
          If Not .LidLocked Then State = Interlock
          If .ReserveLevel > 10 Then Timer = .Parameters.ReserveTimeAfterRinse
          If Timer.Finished Then
             State = RinseToDrain
             Timer = .Parameters.ReserveRinseToDrainTime
          End If
        
        Case RinseToDrain
          If Not .ManualReserveTransferPB Then State = Off
          If Timer.Finished Then
             State = Drain
             Timer = .Parameters.ReserveDrainTime
          End If
          
        Case Drain
          If Not .ManualReserveTransferPB Then State = Off
          If .ReserveLevel > 10 Then Timer = .Parameters.ReserveDrainTime
          If Timer.Finished Then
             State = TurnOff
             .ManualReserveTransferFinished = True
          End If
          
        Case TurnOff
           If .ManualReserveTransferPB = False Then
             .ManualReserveTransferFinished = False
             State = Off
           End If
    
    End Select
    End With
End Sub
Public Property Get IsTransferring() As Boolean
 If (State = Transfer1) Or (State = Rinse) Or (State = Transfer2) Then IsTransferring = True
End Property
Public Property Get IsRinsing() As Boolean
  If (State = Rinse) Or (State = RinseToDrain) Then IsRinsing = True
End Property
Public Property Get IsDraining() As Boolean
  If (State = Drain) Or (State = RinseToDrain) Then IsDraining = True
End Property


#End If
End Class