'American & Efird - Mt. Holly Then Platform 
' [2024-09-09] Mt. Holly
' [2022-05-06] Tina14 GMX
' [2017-08-11] Add parameter "FillColdOverrideTime" to enable opening FillCold valve when filling to cool with hot valve and machine not reached level in parameter time
'              Use to prevent delays when Hot Supply Water is not available.  Disable by setting parameter to '0'
' [2016-05-12] John Smith wants the machine to fill to the defined level, then hold to cool.  Do not step to the atmospheric drain if the TempSafe flag is made
' [2016-05-12] Due to some machines not having FillHot valve piped, using timer as insurance to fill with cold water after hardcoded delay (60seconds)

Imports Settings
Imports Utilities.Translations

<Command("Hot Drain", "", "", "70", "'StandardTimeDrainHot=10"),
  TranslateCommand("es", "Drenado en Caliente"),
  Description("Drain the machine under pressure."),
  TranslateDescription("es", "Drena la máquina por el tiempo del DrainTime del parámetro (s) después de que el nivel ma's kier sea menos de el 1%"),
  Category("Machine Functions"), TranslateCategory("es", "Machine Functions")>
Public Class HD : Inherits MarshalByRefObject : Implements ACCommand
  Private Const commandName_ As String = ("HD: ")

  'Command states
  Public Enum EState
    Off
    SelectHighTempDrain

    ' Cool Portion (If HD Not Enabled/Installed)
    Interlock
    RampToTempSafe
    HoldCoolDelay

    ' HD Portion
    HighTempDrainPause
    HighTempDrainStart
    HighTempDrain
    HoldForPress
    HoldForTime
    LevelFlush1
    LevelSettle1
    FillToCool
    HoldToCool

    ' DR Portion
    DrainPause
    DrainStart
    DrainEmpty
    LevelFlush2
    LevelSettle2
    DrainEmpty2
    Complete
  End Enum
  Public State As EState
  Public StateRestart As EState

  Public Status As String
  Public HDCompleted As Boolean
  Public FlushCount As Integer
  Public FillLevel As Integer
  Public MachineNotClosed As Boolean

  Public Timer As New Timer
  Public TimerAdvance As New Timer
  Public TimerBeforeFlush As New Timer
  Public TimerOverrun As New Timer

  Friend ReadOnly TimerVentAlarm As New Timer
  Public ReadOnly Property TimeVentAlarm As Integer
    Get
      Return TimerVentAlarm.Seconds
    End Get
  End Property
  Public ReadOnly Property HotDrainEnabled As Boolean
    Get
      Return Settings.HighTempDrainEnabled > 0
    End Get
  End Property

  'Keep a local reference to the control code object for convenience
  Private ReadOnly controlCode As ControlCode
  Public Sub New(ByVal controlCode As ControlCode)
    Me.controlCode = controlCode
    Me.controlCode.Commands.Add(Me)
  End Sub

  Public Sub ParametersChanged(ByVal ParamArray param() As Integer) Implements ACCommand.ParametersChanged
    'If the command is not running then do nothing the changes will be picked up when the command starts
    If Not IsOn Then Exit Sub
  End Sub

  Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With controlCode

      ' Cancel Commands
      .AC.Cancel() : .AT.Cancel()
      .WK.Cancel()
      .DR.Cancel() : .FI.Cancel() : .HC.Cancel() ': .HD.Cancel() 
      .RC.Cancel() : .RH.Cancel() : .RI.Cancel() : .TM.Cancel()
      .LD.Cancel() : .PH.Cancel() : .SA.Cancel() : .UL.Cancel()
      If .RT.IsForeground Then .RT.Cancel()
      .CO.Cancel() : .HC.Cancel() : .HE.Cancel() : .WT.Cancel()
      .TemperatureControl.Cancel()

      ' Cancel PressureOn command if active during Atmospheric Commands
      .PR.Cancel()

      'Set the overrun to the higher value
      TimerOverrun.Minutes = .Parameters.StandardTimeDrainHot
      TimerVentAlarm.Seconds = .Parameters.HDVentAlarmTime

      ' Set flow direction
      .PumpControl.SwitchInToOut()

      'Set default state
      State = EState.SelectHighTempDrain
      FlushCount = 0

      'This is a foreground command so do not step on
      Return False
    End With
  End Function

  Function Run() As Boolean Implements ACCommand.Run
    With controlCode

      Dim pauseHotDrop As Boolean = .EStop OrElse .Parent.IsPaused OrElse Not (.MachineClosed)
      Dim pauseAtmDrain As Boolean = .EStop OrElse .Parent.IsPaused OrElse Not (.MachineClosed) AndAlso .MachineSafe

      Select Case State

        Case EState.Off
          Status = ""

        Case EState.SelectHighTempDrain
          ' Check if HD is enabled (some machines do not have High Temp Drain hardware installed)
          If Settings.HighTempDrainEnabled > 0 Then
            HDCompleted = False
            State = EState.HighTempDrainStart
            Timer.Seconds = 5
          Else
            ' HD not enabled, must cool first
            State = EState.Interlock
            Timer.Seconds = 2
          End If
          Status = ""

          '******************************************************************************************************
          '****************                          COOL TO ATMOSPHERIC PORTION                 ****************
          '******************************************************************************************************
        Case EState.Interlock
          If Timer.Finished Then
            ' Start Temperature control to cool down
            .CO.Cancel() : .HE.Cancel() : .HC.Cancel() : .TP.Cancel()
            .TemperatureControl.Start(.Temp, 600, 0)
            State = EState.RampToTempSafe
            Timer.Seconds = 2
          End If
          Status = commandName_ & Translate("Starting") & Timer.ToString(1)

        Case EState.RampToTempSafe
          If Not .TempSafe Then Timer.Seconds = 2
          If .TemperatureControl.Pid.IsHold Then
            .TemperatureControl.Cancel()
            .AirpadOn = False
            State = EState.HoldCoolDelay
          End If
          Status = commandName_ & .TemperatureControl.Status

        Case EState.HoldCoolDelay
          If .MachineSafe Then
            .CO.Cancel() : .HE.Cancel() : .TP.Cancel()
            .TemperatureControl.Cancel()
            .PumpControl.StopMainPump()
            .GetRecordedLevel = False
            Timer.Seconds = 2
            .AirpadOn = False
            .PR.Cancel()
            State = EState.DrainStart
          End If
          Status = commandName_ & Translate("Wait For Safe")



          '******************************************************************************************************
          '****************                              HIGH DRAIN PORTION                      ****************
          '******************************************************************************************************
        Case EState.HighTempDrainPause
          If pauseHotDrop Then Timer.Seconds = 5
          If Timer.Finished Then
            State = StateRestart
            Timer.Seconds = 5
          End If
          Status = commandName_ & Translate("Paused") & Timer.ToString(1)

        Case EState.HighTempDrainStart
          Status = commandName_ & Translate("Starting") & Timer.ToString(1)
          If Not .MachineClosed Then
            MachineNotClosed = True
            Timer.Seconds = 5
            If Not .MachineClosed Then Status = commandName_ & Translate("Machine Not Closed")
          Else
            ' Machine was not closed
            If MachineNotClosed Then
              Timer.Seconds = 5
              ' Clear Flag to continue
              If Not .AdvancePb Then
                Status = commandName_ & Translate("Hold Advance") & TimerAdvance.ToString(1)
                TimerAdvance.Seconds = 2
              Else
                'Attempting to Advance
                If .Parent.Signal <> "" Then .Parent.Signal = ""
                Status = commandName_ & Translate("Advancing") & TimerAdvance.ToString(1)
                If TimerAdvance.Finished Then MachineNotClosed = False
              End If
            Else
              If .EStop Then Status = commandName_ & Translate("EStop") & Timer.ToString(1) : Timer.Seconds = 5
              If .Parent.IsPaused Then Status = commandName_ & Translate("Paused") & Timer.ToString(1) : Timer.Seconds = 5
              ' carry on...
              If Timer.Finished Then
                State = EState.HighTempDrain
                .AirpadOn = True
                .PumpControl.StopMainPump()
                .GetRecordedLevel = False
                TimerVentAlarm.Seconds = .Parameters.HDVentAlarmTime
                TimerBeforeFlush.Seconds = MinMax(.Parameters.HotDrainMachineFlushTime, 120, 600)
              End If
            End If
          End If
          ' Pause Command Conditions
          If pauseHotDrop Then
            State = EState.HighTempDrainPause
            StateRestart = EState.HighTempDrainStart
          End If

        Case EState.HighTempDrain
          If .MachineLevel > 50 Then Timer.Seconds = MinMax(.Parameters.HDDrainTime, 30, 300)
          .Alarms.LosingLevelInTheVessel = False
          If Timer.Finished Then
            .AirpadOn = False
            .PR.Cancel()
            State = EState.HoldForPress
            Timer.Seconds = 120
          End If
          ' 2021/04/20 - DP Level Transmitter issue seen during startup
          If (.Parameters.HotDrainMachineFlushTime > 0) Then
            If TimerBeforeFlush.Finished Then
              .AirpadOn = False
              .PR.Cancel()
              State = EState.HoldForPress
              Timer.Seconds = 120
            End If
          End If
          ' Pause Command Conditions
          If pauseHotDrop Then
            State = EState.HighTempDrainPause
            StateRestart = EState.HighTempDrainStart
          End If
          Status = commandName_ & Translate("Hot Draining") & (" ") & (.MachineLevel / 10).ToString("#0.0") & "%" & Timer.ToString(1)

        Case EState.HoldForPress
          Timer.Seconds = 120
          If .PressSafe Then
            State = EState.HoldForTime
          End If
          ' Pause Command Conditions
          If pauseHotDrop Then
            State = EState.HighTempDrainPause
            StateRestart = EState.HighTempDrainStart
          End If
          Status = commandName_ & Translate("Wait for PressSafe") & (" ") & .MachinePressureDisplay

        Case EState.HoldForTime
          If Timer.Finished Then
            HDCompleted = True
            State = EState.LevelFlush1
            Timer.Seconds = MinMax(.Parameters.LevelGaugeFlushTime, 10, 120)
          End If
          ' Double-Check
          If Not .PressSafe Then State = EState.HoldForPress
          ' Pause Command Conditions
          If pauseHotDrop Then
            State = EState.HighTempDrainPause
            StateRestart = EState.HighTempDrainStart
          End If
          Status = commandName_ & Translate("Holding") & Timer.ToString(1)



          '******************************************************************************************************
          '****************                          FILL TO COOL PORTION                        ****************
          '******************************************************************************************************

        Case EState.LevelFlush1
          ' Flush level before filling in the event that level remains due to a faulty dp transmitter
          If Timer.Finished Then
            State = EState.LevelSettle1
            Timer.Seconds = MinMax(.Parameters.LevelGaugeSettleTime, 10, 120)
          End If
          Status = commandName_ & Translate("Flush Level") & (" ") & Translate("To Cool") & Timer.ToString(1)

        Case EState.LevelSettle1
          If Timer.Finished Then
            State = EState.FillToCool
            Timer.Seconds = MinMax(.Parameters.OverFillTime, 1, 20)
          End If
          Status = commandName_ & Translate("Level Settle") & Timer.ToString(1)

        Case EState.FillToCool
          FillLevel = MinMax(.Parameters.HDHiFillLevel, .Parameters.FillLevelMinimum, .Parameters.FillLevelMaximum)
          If FillLevel > 999 Then FillLevel = 999
          ' [2016-05-12] John Smith wants the machine to fill to the defined level, then hold to cool.  Do not step to the atmospheric drain if the TempSafe flag is made
          If (.MachineLevel < FillLevel) Then
            Timer.Seconds = MinMax(.Parameters.OverFillTime, 1, 20)
          Else
            If Timer.Finished Then
              State = EState.HoldToCool
              .PumpControl.StartAuto()
              Timer.Seconds = MinMax(.Parameters.HDFillRunTime, 60, 300)
            End If
          End If
          ' Pause Command Conditions
          If pauseHotDrop Then
            State = EState.HighTempDrainPause
            StateRestart = EState.HighTempDrainStart
          End If
          Status = commandName_ & Translate("Filling") & (" ") & (.MachineLevel / 10).ToString("#0.0") & "% / " & (FillLevel / 10).ToString("#0.0") & "%"

        Case EState.HoldToCool
          If Not (.TempSafe AndAlso .PumpControl.IsRunning) Then
            Timer.Seconds = MinMax(.Parameters.HDFillRunTime, 60, 300)
            ' Restart pump if needed
            If Not .PumpControl.IsRunning Then
              .PumpControl.StartAuto()
              Status = commandName_ & Translate("Waiting for Pump") & Timer.ToString(1)
            End If
          End If
          If .TempSafe AndAlso .PumpControl.IsRunning AndAlso Timer.Finished Then
            State = EState.DrainEmpty
            .PumpControl.StopMainPump()
            Timer.Seconds = MinMax(.Parameters.DrainMachineTime, 30, 300)
          End If
          ' Pause Command Conditions
          If pauseHotDrop Then
            State = EState.HighTempDrainPause
            StateRestart = EState.HoldToCool
          End If
          Status = commandName_ & Translate("Hold To Cool") & Timer.ToString(1)


          '******************************************************************************************************
          '****************                          ATMOSPHERIC DRAIN PORTION                   ****************
          '******************************************************************************************************
        Case EState.DrainPause
          If pauseHotDrop Then Timer.Seconds = 5
          .AirpadOn = False
          If Timer.Finished Then
            State = StateRestart
            Timer.Seconds = 5
          End If
          Status = commandName_ & Translate("Paused") & Timer.ToString(1)

        Case EState.DrainStart
          If Not .PressSafe Then Timer.Seconds = 5
          If Timer.Finished Then
            .GetRecordedLevel = False
            State = EState.DrainEmpty
            Timer.Seconds = MinMax(.Parameters.DrainMachineTime, 30, 300)
            TimerBeforeFlush.Seconds = MinMax(.Parameters.DrainMachineFlushTime, 60, 600)
          End If
          ' Pause Command Conditions
          If pauseAtmDrain Then
            State = EState.DrainPause
            StateRestart = EState.DrainStart
          End If
          Status = commandName_ & Translate("Starting") & Timer.ToString(1)

        Case EState.DrainEmpty
          ' 1st Machine Drain based solely on time delay
          If .MachineLevel > 50 Then Timer.Seconds = MinMax(.Parameters.DrainMachineTime, 30, 300)
          ' Continue draining until delay timer has completed
          If Timer.Finished Then
            State = EState.LevelFlush2
            Timer.Seconds = MinMax(.Parameters.LevelGaugeFlushTime, 10, 60)
          End If
          ' Waiting too long to drain empty - may be issue with DP level transmitter
          If TimerBeforeFlush.Finished AndAlso (.Parameters.DrainMachineFlushTime > 0) Then
            .PumpControl.StopMainPump()
            Timer.Seconds = MinMax(.Parameters.LevelGaugeFlushTime, 10, 120)
            State = EState.LevelFlush2
          End If
          ' Pause Command Conditions
          If pauseAtmDrain Then
            State = EState.DrainPause
            StateRestart = EState.DrainStart
          End If
          Status = commandName_ & Translate("Draining") & (" ") & (.MachineLevel / 10).ToString("#0.0") & ("%") & Timer.ToString(1)

        Case EState.LevelFlush2
          If Timer.Finished Then
            FlushCount += 1
            State = EState.DrainEmpty2
            Timer.Seconds = MinMax(.Parameters.LevelGaugeSettleTime, 10, 120)
            TimerBeforeFlush.Seconds = MinMax(.Parameters.DrainMachineFlushTime, 60, 600)
          End If
          ' Pause Command Conditions
          If pauseAtmDrain Then
            State = EState.DrainPause
            StateRestart = EState.DrainStart
          End If
          Status = commandName_ & Translate("Flush Level") & (" ") & (.MachineLevel / 10).ToString("#0.0") & ("%") & Timer.ToString(1)

        Case EState.LevelSettle2
          If Timer.Finished Then
            State = EState.DrainEmpty2
            Timer.Seconds = MinMax(.Parameters.LevelGaugeSettleTime, 10, 120)
            TimerBeforeFlush.Seconds = MinMax(.Parameters.DrainMachineFlushTime, 60, 600)
          End If
          Status = Translate("Level Settle") & (" ") & (.MachineLevel / 10).ToString("#0.0") & ("%") & Timer.ToString(1)

        Case EState.DrainEmpty2
          If .MachineLevel > 50 Then Timer.Seconds = MinMax(.Parameters.LevelGaugeSettleTime, 10, 120)
          ' Continue draining until delay timer has completed
          If Timer.Finished Then
            State = EState.Complete
            Timer.Seconds = 5
          End If
          ' Waiting too long to drain empty - may be issue with DP level transmitter
          If TimerBeforeFlush.Finished AndAlso (.Parameters.DrainMachineFlushTime > 0) Then
            .PumpControl.StopMainPump()
            Timer.Seconds = MinMax(.Parameters.LevelGaugeFlushTime, 10, 60)
            State = EState.LevelFlush2
          End If
          ' Lost machine safe
          If Not .MachineSafe Then
            State = EState.Interlock
            Timer.Seconds = 2
          End If
          ' Pause Command Conditions
          If pauseAtmDrain Then
            State = EState.DrainPause
            StateRestart = EState.DrainStart
          End If
          Status = commandName_ & Translate("Draining") & (" ") & (.MachineLevel / 10).ToString("#0.0") & ("%") & Timer.ToString(1)

        Case EState.Complete
          If Timer.Finished Then
            '.MachineVolume = 0
            State = EState.Off
          End If
          Status = commandName_ & Translate("Completing") & Timer.ToString(1)

      End Select
    End With
  End Function

  Public Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
    StateRestart = EState.Off
  End Sub

  ReadOnly Property IsOn As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property

  ReadOnly Property IsFillingToCool As Boolean
    Get
      Return (State = EState.FillToCool)
    End Get
  End Property

  ReadOnly Property IsOverrun As Boolean
    Get
      Return IsOn AndAlso TimerOverrun.Finished
    End Get
  End Property

  ReadOnly Property IoAirpadCutoff As Boolean
    Get
      Return ((State >= EState.HoldForPress) AndAlso (State <= EState.FillToCool)) OrElse
             ((State >= EState.DrainStart) AndAlso (State <= EState.DrainEmpty2)) OrElse
             HDCompleted
    End Get
  End Property

  ReadOnly Property IoAirpadBleed As Boolean
    Get
      Return ((State >= EState.FillToCool) AndAlso (State <= EState.HoldToCool)) OrElse
            ((State >= EState.DrainStart) AndAlso (State <= EState.DrainEmpty2)) OrElse
              HDCompleted
    End Get
  End Property

  ReadOnly Property IoMachineDrain As Boolean
    Get
      Return (State >= EState.DrainEmpty) AndAlso (State <= EState.DrainEmpty2)
    End Get
  End Property

  ReadOnly Property IoFill As Boolean
    Get
      Return (State = EState.FillToCool) 'AndAlso ((Timer.Finished AndAlso (controlCode.Parameters.FillColdOverrideTime > 0)))
    End Get
  End Property

  ReadOnly Property IoHighTempBypass As Boolean
    Get
      Return (State >= EState.LevelFlush1) AndAlso (State <= EState.DrainEmpty2)
    End Get
  End Property

  ReadOnly Property IoHighTempDrain As Boolean
    Get
      Return (State >= EState.HighTempDrain) AndAlso (State <= EState.HoldForTime)
    End Get
  End Property

  ReadOnly Property IoLevelFlush As Boolean
    Get
      Return (State = EState.LevelFlush1) OrElse (State = EState.LevelFlush2)
    End Get
  End Property

  ReadOnly Property IoTopWash As Boolean
    Get
      Return (State = EState.FillToCool) OrElse IoMachineDrain
    End Get
  End Property



  ' TODO
#If 0 Then

'Hot Dain command
Option Explicit
Implements ACCommand
Public Enum HDState
  HDOff
  HDHiTempDrain
  HDLevelFlush1
  HDLevelSettle1
  HDWaitForPress
  HDHoldForTime
  HDFillToCool
  HDWaitingToCool
  HDAtmosDrain
  HDLevelFlush
  HDLevelSettle
End Enum
Public State As HDState
Public StateString As String
Public Timer As New acTimer, OverrunTimer As New acTimer
Public VentAlarmTimer As New acTimer
Public LevelGaugeTimer As New acTimer
Public LevelGaugeFlushCount As Integer
Public HDCompleted As Boolean

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Name=Hot Drain\r\nMinutes=8\r\nHelp=Drains the machine for parameter time DR Drain Time (s) after the expansion tank level float input is lost."
  Dim ControlCode As ControlCode: Set ControlCode = ControlObject
  With ControlCode
    .AC.ACCommand_Cancel
    .AT.ACCommand_Cancel: .CO.ACCommand_Cancel: .DR.ACCommand_Cancel
    .FI.ACCommand_Cancel: .HE.ACCommand_Cancel: .HC.ACCommand_Cancel
    .LD.ACCommand_Cancel: .PH.ACCommand_Cancel: .PR.ACCommand_Cancel
    .RC.ACCommand_Cancel: .RH.ACCommand_Cancel: .RI.ACCommand_Cancel
    If .RF.FillType = 86 Then .RF.ACCommand_Cancel:
    .RT.ACCommand_Cancel: .RW.ACCommand_Cancel: .SA.ACCommand_Cancel
    .TC.ACCommand_Cancel: .TM.ACCommand_Cancel: .TP.ACCommand_Cancel
    .UL.ACCommand_Cancel: .WK.ACCommand_Cancel: .WT.ACCommand_Cancel
    .TemperatureControl.Cancel: .TemperatureControlContacts.Cancel
    .PumpRequest = False
    VentAlarmTimer = .Parameters.HDVentAlarmTime
    LevelGaugeTimer = .Parameters.LevelFlushOverrideTime
    LevelGaugeFlushCount = 0
    State = HDHiTempDrain
    OverrunTimer = 600
  End With
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  
  With ControlCode
    Select Case State
    
      Case HDOff
        StateString = ""
        
      Case HDHiTempDrain
        If (.VesselLevel > 10) Then
          StateString = "HD: Hot Draining " & Pad(.VesselLevel / 10, "0", 3) & "%"
          Timer = .Parameters.DRDrainTime
        Else
          State = HDWaitForPress
          .AirpadOn = False
        End If
        .PumpOnCount = 0
        .GetRecordedLevel = False
        .Alarms_LosingLevelInTheVessel = False
        ' Level gauge flush override timer
        If Not .PressSafe Then LevelGaugeTimer = .Parameters.LevelFlushOverrideTime + 1
        If (.Parameters.LevelFlushOverrideTime > 0) And LevelGaugeTimer.Finished Then
          State = HDLevelFlush1
          Timer = .Parameters.LevelGaugeFlushTime
          LevelGaugeFlushCount = LevelGaugeFlushCount + 1
        End If
        
      Case HDLevelFlush1
        StateString = "HD: Level Flush Auto " & Pad(LevelGaugeFlushCount, "0", 1) & " & TimerString(Timer.TimeRemaining)"
        If Timer.Finished Then
          Timer = .Parameters.LevelGaugeSettleTime
          State = HDLevelSettle1
        End If
        
      Case HDLevelSettle1
        StateString = "HD: Level Settle Auto " & Pad(LevelGaugeFlushCount, "0", 1) & " " & TimerString(Timer.TimeRemaining)
        If Timer.Finished Then
          LevelGaugeTimer = .Parameters.LevelFlushOverrideTime
          State = HDHiTempDrain
        End If
                  
      Case HDWaitForPress
        StateString = "HD: Wait for PressSafe "
        If .PressSafe Then
          State = HDHoldForTime
          Timer = 120
        End If
          
      Case HDHoldForTime
        StateString = "HD: Hold For Time " & TimerString(Timer.TimeRemaining)
        If Timer.Finished Then
           State = HDLevelFlush
           Timer = .Parameters.LevelGaugeFlushTime
        End If
          
      Case HDFillToCool
        StateString = "HD: Fill To Cool " & Pad(.VesselLevel / 10, "0", 3) & "% / " & Pad(.Parameters.HDHiFillLevel / 10, "0", 3) & "%"
        If .TempSafe Then
           State = HDAtmosDrain
           Timer = .Parameters.HDDrainTime
        End If
        If .VesselLevel >= .Parameters.HDHiFillLevel Then State = HDWaitingToCool
          
      Case HDWaitingToCool
        StateString = "HD: Wait for Temp Below 190F "
        If .TempSafe And (.VesTemp <= 1900) Then
           State = HDAtmosDrain
           Timer = .Parameters.HDDrainTime
        End If
          
      Case HDAtmosDrain
        If (.VesselLevel > 10) Then
          StateString = "HD: Atmospheric Drain " & Pad(.VesselLevel / 10, "0", 3) & "%"
          Timer = .Parameters.HDDrainTime
        Else: StateString = "HD: Atmospheric Drain " & TimerString(Timer.TimeRemaining)
        End If
        If Timer.Finished Then
          State = HDLevelFlush
          Timer = .Parameters.LevelGaugeFlushTime
        End If
        ' Level gauge flush override timer
        If Not .PressSafe Then LevelGaugeTimer = .Parameters.LevelFlushOverrideTime + 1
        If (.Parameters.LevelFlushOverrideTime > 0) And LevelGaugeTimer.Finished Then
          State = HDLevelFlush
          Timer = .Parameters.LevelGaugeFlushTime
          LevelGaugeFlushCount = LevelGaugeFlushCount + 1
        End If
        
      Case HDLevelFlush
        StateString = "HD: Level Flush " & Pad(LevelGaugeFlushCount, "0", 1) & TimerString(Timer.TimeRemaining)
        If Timer.Finished Then
          Timer = .Parameters.LevelGaugeSettleTime
          State = HDLevelSettle
        End If
      
      Case HDLevelSettle
        StateString = "HD: Level Settle " & Pad(LevelGaugeFlushCount, "0", 1) & " " & TimerString(Timer.TimeRemaining)
        If Timer.Finished Then
          State = HDOff
          HDCompleted = True
        End If
        
    End Select
  End With
End Sub


Friend Property Get IsActive() As Boolean
  If (State >= HDHiTempDrain) And (State <= HDLevelSettle) Then IsActive = True
End Property

Friend Property Get IsWaitingForPressure() As Boolean
 If (State = HDWaitForPress) Then IsWaitingForPressure = True
End Property
Friend Property Get IsHoldingForTime() As Boolean
 If (State = HDHoldForTime) Then IsHoldingForTime = True
End Property
Friend Property Get IsFillingToCool() As Boolean
 If State = HDFillToCool Then IsFillingToCool = True
End Property
Friend Property Get HasFilledWaitingToCool() As Boolean
 If State = HDWaitingToCool Then HasFilledWaitingToCool = True
End Property
Friend Property Get IsAtmosDraining() As Boolean
 If State = HDAtmosDrain Then IsAtmosDraining = True
End Property


Friend Property Get IoHotDrain() As Boolean
  If (State = HDHiTempDrain) Or (State = HDLevelFlush1) Or (State = HDLevelSettle1) Or _
     (State = HDWaitForPress) Or (State = HDHoldForTime) Then IoHotDrain = True
End Property
Friend Property Get IoVent() As Boolean
  If (State = HDFillToCool) Or (State = HDAtmosDrain) Or (State = HDWaitingToCool) Or _
     (State = HDLevelFlush) Or (State = HDLevelSettle) Then IoVent = True
End Property



#End If
End Class

