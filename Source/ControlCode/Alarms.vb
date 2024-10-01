Public NotInheritable Class Alarms : Inherits MarshalByRefObject
  ' For convenience
  Private ReadOnly controlCode As ControlCode

  Private powerUpTimer As New Timer
  Private ReadOnly firstScanDone As Boolean


  ' Alphabetical Order
  Public AddPumpTripped As Boolean
  Public AddLevelTooHigh As Boolean 'LevelTooHighInAddTank
  Public AutoDispenseDelay As Boolean 'TODO

  Public BlendFillTempProbeError As Boolean

  Public CheckLevelTransmitter As Boolean ' TODO

  Private CondensateAlarmTimer As New Timer
  Public CondensateTempTooHigh As Boolean ' TODO

  Public ControlModeDebugOn As Boolean
  Public ControlModeOverrideOn As Boolean  ' OverrideModeSelected
  Public ControlModeTestOn As Boolean      ' TestModeSelected
  Public ControlPowerFailure As Boolean

  Public DiffPressTransmitter As Boolean
  Public DispenserCompleteDelay As Boolean  ' TODO
  Public DrainTooLongCheckLevel As Boolean ' TODO - Need?

  Public EmergencyStopPB As Boolean


  Public LevelTooLowForPump As Boolean
  Public LidNotClosed As Boolean
  Public LidNotLocked As Boolean
  Public LosingLevelInTheVessel As Boolean
  Public LowFlowRate As Boolean

  Public MainPumpTripped As Boolean
  Public MainPumpAlarm As Boolean
  Public MainPumpTurnedOnOffTooManyTimes As Boolean ' TODO Need?

  Public PlcComs1 As Boolean    'MainPLCNotResponding
  Public PlcComs2 As Boolean    'DrugroomPLCNotResponding
  Public PumpLevelLost As Boolean

  Public RedyeIssueCheckMachineTank As Boolean
  Public ReserveTempProbeError As Boolean
  Public ReserveTempTooHigh As Boolean

  Public SampleLidNotClosed As Boolean
  Public SysInStandby As Boolean

  Public Tank1DispenseError As Boolean    ' TODO

  Private Tank1TempPropeErrorTimer As New Timer
  Public Tank1TempProbeError As Boolean
  Public Tank1TempTooHigh As Boolean      ' TODO
  Public Tank1TooLongToHeat As Boolean


  Public TemperatureTooHigh As Boolean
  Public TemperatureHigh As Boolean
  Public TemperatureLow As Boolean
  '  Public TestValueTooHigh As Boolean 
  Public TooLongToHeat As Boolean

  Public TopWashBlocked As Boolean

  Public VentNotOpenOnHotDrain As Boolean
  Public VesselTempProbeError As Boolean

  ' Timers
  Public AddPumpAlarmTimer As New Timer
  Public PumpRunningTimer As New Timer
  Public PumpAlarmTimer As New Timer
  Public LidLockedTimer As New Timer


  Public Sub New(ByVal controlCode As ControlCode)
    ' Save control code reference locally
    Me.controlCode = controlCode

    powerUpTimer.Seconds = 8
  End Sub

  Public Sub Run()
    If Not powerUpTimer.Finished Then Exit Sub

    With controlCode

      ' Add Pump Alarm
      If (Not .IO.AddPumpStart) OrElse .IO.AddPumpRunning Then AddPumpAlarmTimer.TimeRemaining = 5
      AddPumpTripped = AddPumpAlarmTimer.Finished


      If (.IO.CondensateTemp > 320) AndAlso (.IO.CondensateTemp < .Parameters.TempCondensateLimit) Then CondensateAlarmTimer.Seconds = .Parameters.TempCondensateLimitTime
      CondensateTempTooHigh = CondensateAlarmTimer.Finished AndAlso (.Parameters.TempCondensateLimit > 0) AndAlso (.Parameters.TempCondensateLimitTime > 0)


      'Manual override modes
      ControlModeDebugOn = (.Parent.Mode = Mode.Debug) AndAlso (.Parameters.Demo = 0)
      ControlModeOverrideOn = (.Parent.Mode = Mode.Override)
      ControlModeTestOn = (.Parent.Mode = Mode.Test)

      ControlPowerFailure = Not .IO.PowerOn

      ' Pump is off: differential pressure should be 0 (analog input = 50%)
      ' At 60% output speed, differential pressure is around 600mb = 6000
      '   Signal Alarm if Pump Off and DP > 200mb or if pump Running and DP > 1.5bar (15000)
      DiffPressTransmitter = (Not .PumpControl.IsRunning AndAlso (.PackageDifferentialPressure > 2000)) OrElse
                                  (.PumpControl.IsRunning AndAlso (.PackageDifferentialPressure > 15000))



      '  Alarms_DrainTooLongCheckDrainFloat = DR.IoDrain AndAlso DR.IsOverrun AndAlso (Not IO.LevelDrain)
      '  Alarms_DrainTooLongCheckFillFloat = DR.IoDrain AndAlso DR.IsOverrun AndAlso (IO.Level1 OrElse IO.Level2)

      EmergencyStopPB = .IO.EmergencyStop


      BlendFillTempProbeError = (.IO.FillTemp < 320 OrElse .IO.FillTemp > 3000)

      If (.Parameters.MaxLevelLossAmount > 0) AndAlso .Parent.IsProgramRunning Then
        LosingLevelInTheVessel = .GotRecordedLevel AndAlso (.MachineLevelRecorded > 0) AndAlso (.MachineLevel <= (.MachineLevelRecorded - .Parameters.MaxLevelLossAmount))
      Else
        LosingLevelInTheVessel = False
      End If

      ' Main Pump Alarms
      Static PumpRunningTimer As New Timer
      If (Not .IO.MainPump) OrElse .IO.PumpRunning Then PumpRunningTimer.TimeRemaining = 5
      MainPumpTripped = PumpRunningTimer.Finished

      Static PumpAlarmTimer As New Timer
      If (Not .IO.PumpAlarm) Then PumpAlarmTimer.TimeRemaining = 5
      MainPumpAlarm = PumpAlarmTimer.Finished

      PlcComs1 = .IO.PLC1Timer.Finished AndAlso (.Parent.Mode <> Mode.Debug) AndAlso (.Parameters.PLCComsLossDisregard <> 1)
      PlcComs2 = .IO.PLC2Timer.Finished AndAlso (.Parent.Mode <> Mode.Debug) AndAlso (.Parameters.PLCComsLossDisregard <> 1)

      PumpLevelLost = .PumpControl.IsAlarmLevelLost

      ReserveTempProbeError = (.IO.ReserveTemp < 320) OrElse (.IO.ReserveTemp > 3000)

      SampleLidNotClosed = False ' Parent.IsProgramRunning AndAlso Not SampleClosed AndAlso Not (LD.IsOn OrElse SA.IsSampleIsolate OrElse UL.IsOn)

      SysInStandby = False 'todo Parent.IsProgramRunning AndAlso (PumpRequest OrElse FI.IsOn) AndAlso (Not SA.IsOn)

      Tank1DispenseError = False

      If ((.IO.Tank1Temp > 320) AndAlso (.IO.Tank1Temp < 3000)) Then Tank1TempPropeErrorTimer.Seconds = 5
      Tank1TempProbeError = (.Parameters.Tank1TempProbe > 0) AndAlso (Not PlcComs2) AndAlso Tank1TempPropeErrorTimer.Finished

      Tank1TooLongToHeat = False 'KP.KP1.IsHeatTankOverrun Or LA.KP1.IsHeatTankOverrun
      Tank1TempTooHigh = False ' AndAlso (Not Tank1TempProbeError) AndAlso (.IO.Tank1Temp > .Parameters.Tank1HighTempLimit) AndAlso (.Parameters.Tank1HighTempLimit > 0)


      TemperatureTooHigh = (.Temp > 2800) OrElse (.IO.ContactTemp)


      TemperatureHigh = (.TemperatureControl.TempHighAlarmTimer.Finished)
      TemperatureLow = (.TemperatureControl.TempLowAlarmTimer.Finished)

      TooLongToHeat = .HE.IsOverrun

      TopWashBlocked = (.RI.IsTopWashBlocked) OrElse (.RH.IsTopWashBlocked)

      ' Alarms_RedyeIssueCheckMachineTank managed in KP command

      VentNotOpenOnHotDrain = (.HD.IoFill AndAlso .HD.TimerVentAlarm.Finished AndAlso Not .IO.Vent)

      ' Alarm for Lid Not Locked
      If (Not .Parent.IsProgramRunning) OrElse (.MachineClosed) OrElse (.LD.IsOn OrElse .SA.IsActive OrElse .UL.IsOn) Then LidLockedTimer.Seconds = 2

      If .Parent.IsProgramRunning AndAlso .PumpControl.PumpRequest AndAlso Not (.SA.IsActive OrElse .LD.IsOn OrElse .UL.IsOn) AndAlso Not .LidLocked Then
        '.PumpRequest = False
        .PumpControl.StopMainPump()
        LidNotLocked = True
      End If
      If Not (.IO.AdvancePb AndAlso .LidLocked) Then LidLockedTimer.Seconds = 2
      If .IO.AdvancePb AndAlso LidLockedTimer.Finished Then
        ' If Parent.IsProgramRunning And (PumpLevel Or DR.IsInterlocked Or DR.IsDraining Or RT.IsTransfer) Then PumpRequest = True
        LidNotLocked = False
      End If

      LidNotClosed = .LidControl.IoLowerLid AndAlso .LidControl.Timer.Finished AndAlso Not .IO.KierLidClosed


#If 0 Then ' TODO

      'Lid Locked
  If Parent.IsProgramRunning And PumpRequest And Not (SA.IsSampling Or LD.IsOn Or UL.IsUnloading) And Not LidLocked Then
    PumpRequest = False
    Alarms_LidNotLocked = True
  End If
  If Not (IO_RemoteRun And LidLocked) Then LidLockTimer.TimeRemaining = 2
  If LidLockTimer.Finished Then
    If Parent.IsProgramRunning And (PumpLevel Or DR.IsInterlocked Or DR.IsDraining Or RT.IsTransfer) Then PumpRequest = True
    Alarms_LidNotLocked = False
  End If

'lid alarms
  Alarms_LidNotClosed = (LidControl.IsLowerLid) And LidControl.Timer.Finished _
                        And Not IO_LidLoweredSwitch

#End If


      VesselTempProbeError = (.IO.VesselTemp < 320 OrElse .IO.VesselTemp > 3000)


    End With
  End Sub

End Class
