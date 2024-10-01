Partial Class ControlCode

  Public Delay As DelayValue
  <GraphTrace(0, 1, 0, 1, "Blue")> Public IsDelayed As Boolean

  Private Sub RunDelays()
    Delay = GetDelay()

    If Delay <> DelayValue.NormalRunning Then
      IsDelayed = True
    Else
      IsDelayed = False
    End If
  End Sub

  Private Function GetDelay() As DelayValue

    ' Look at sleeping flag and set delay
    If Parent.IsSleeping Then Return DelayValue.Sleeping

    ' Look at if control is paused
    If Parent.IsPaused Then Return DelayValue.Paused

    ' Commands delayed
    If BO.IsOn Then Return DelayValue.Boilout
    If BR.IsOn Then Return DelayValue.Correction

    ' Look at alarms and set machine delay
    If Alarms.ControlModeDebugOn Then Return DelayValue.MachineError
    If Alarms.ControlModeOverrideOn Then Return DelayValue.MachineError
    If Alarms.ControlModeTestOn Then Return DelayValue.MachineError

    If Alarms.ControlPowerFailure Then Return DelayValue.MachineError
    If Alarms.PlcComs1 Then Return DelayValue.MachineError
    If Alarms.PlcComs2 Then Return DelayValue.MachineError

    If Alarms.EmergencyStopPB Then Return DelayValue.MachineError

    If Alarms.DrainTooLongCheckLevel Then Return DelayValue.MachineError

    If Alarms.LidNotLocked Then Return DelayValue.MachineError
    If Alarms.VesselTempProbeError Then Return DelayValue.MachineError

    If Alarms.MainPumpAlarm Then Return DelayValue.MachineError
    If Alarms.MainPumpTripped Then Return DelayValue.MachineError
    If Alarms.PumpLevelLost Then Return DelayValue.MachineError

    If Alarms.AddPumpTripped Then Return DelayValue.MachineError
    If Alarms.TemperatureTooHigh Then Return DelayValue.MachineError

    'operator delay
    Static operatordelayTimer As New Timer
    If Parent.ButtonText = "" Then operatordelayTimer.Seconds = Parameters.StandardTimeOperator * 60
    If Alarms.SysInStandby Then Return DelayValue.Operator
    If Alarms.ControlModeDebugOn Then Return DelayValue.Operator
    If Alarms.ControlModeOverrideOn Then Return DelayValue.Operator
    If Alarms.ControlModeTestOn Then Return DelayValue.Operator
    If operatordelayTimer.Finished Then Return DelayValue.Operator

    ' Temperature Commands
    If CO.IsOverrun Then Return DelayValue.Cooling
    If HE.IsOverrun Then Return DelayValue.Heating

    'Look at Drain Command for delays
    If DR.IsOverrun Then Return DelayValue.Drain

    'Look at Fill command for delays
    If FI.IsOverrun Then Return DelayValue.Fill
    '   If TU.IsOverrun Then Return DelayValue.Fill

    'Look at Unload command for unload
    If UL.IsOverrun Then Return DelayValue.Unload

    'Look at Load command for load delays
    If LD.IsOverrun Then Return DelayValue.Load

    ' ph check delay
    If PH.IsOverrun Then Return DelayValue.CheckPH

    'Look at Sample and Begin sample command for sample delays
    If SA.IsOverrun Then Return DelayValue.Sample

    'expansion add delay
    If KA.IsOverrun Then Return DelayValue.KitchenAdd
    'TODO  If PB.IsOverrun Then Return DelayValue.KitchenAdd
    '  If DS.IsOverrun Then Return DelayValue.KitchenAdd

    'Step overrun stuff
    If RI.IsOverrun Then Return DelayValue.IsDelayed
    If RH.IsOverrun Then Return DelayValue.IsDelayed
    If TM.IsOverrun Then Return DelayValue.IsDelayed

    If WK.IsOverrun Then Return DelayValue.KitchenAdd

    ' Reserve Tank
    If RD.IsOverrun Then Return DelayValue.MachineError
    If RF.IsOverrun OrElse RT.IsOverrun Then
      If RP.IsWaitReady Then Return DelayValue.Operator
      If KA.IsActive AndAlso KA.Destination = EKitchenDestination.Reserve Then Return DelayValue.Tank1Operator
      If KP.KP1.IsOn AndAlso (KP.KP1.DispenseTank = 2) Then ' (KP.KP1.Param_Destination = EKitchenDestination.Reserve) Then
        With KP
          If .KP1.IsWaitReady Then Return DelayValue.Tank1Operator
          ' If Dispenser Online, test for Dispenser conditions here (Refer to AE.ThenPlatform controlcode)
          ' Note: No current Tank1Heating installed: If .KP1.IsHeatOverrun then return delayValue.Tank1Prepare
          If .KP1.IsInManual AndAlso .KP1.IsOverrun Then Return DelayValue.Tank1Operator
        End With
      End If
      If LA.KP1.IsOn AndAlso (LA.KP1.Param_Destination = EKitchenDestination.Reserve) Then
        With LA
          If .KP1.IsWaitReady Then Return DelayValue.DispenseWaitReady
          ' If Dispenser Online, test for Dispenser conditions here (Refer to AE.ThenPlatform controlcode)
          ' Note: No current Tank1Heating installed: If .KP1.IsHeatOverrun then return delayValue.Tank1Prepare
          If .KP1.IsInManual AndAlso .KP1.IsOverrun Then Return DelayValue.Tank1Operator
        End With
      End If

      ' Default delay value
      Return DelayValue.ReserveTransfer
    End If

    ' Addition Tank
    If AT.IsOverrun OrElse AC.IsOverrun Then
      If AP.IsWaitReady Then Return DelayValue.LocalPrepare
      If AF.IsActive Then Return DelayValue.LocalPrepare
      If KA.IsActive AndAlso (KA.Destination = EKitchenDestination.Add) Then Return DelayValue.Operator
      If KP.KP1.IsOn AndAlso (KP.KP1.DispenseTank = 1) Then ' (KP.KP1.Param_Destination = EKitchenDestination.Add) Then
        With KP
          If .KP1.IsWaitReady Then Return DelayValue.DispenseWaitReady
          ' If Dispenser Online, test for Dispenser conditions here (Refer to AE.ThenPlatform controlcode)
          ' Note: No current Tank1Heating installed: If .KP1.IsHeatOverrun then return delayValue.Tank1Prepare
          If .KP1.IsInManual AndAlso .KP1.IsOverrun Then Return DelayValue.Tank1Operator
        End With
      End If
      If LA.KP1.IsOn AndAlso (LA.KP1.Param_Destination = EKitchenDestination.Add) Then
        With LA
          If .KP1.IsWaitReady Then Return DelayValue.DispenseWaitReady
          ' If Dispenser Online, test for Dispenser conditions here (Refer to AE.ThenPlatform controlcode)
          ' Note: No current Tank1Heating installed: If .KP1.IsHeatOverrun then return delayValue.Tank1Prepare
          If .KP1.IsInManual AndAlso .KP1.IsOverrun Then Return DelayValue.Tank1Operator
        End With
      End If

      ' AT is overrun
      Return DelayValue.LocalPrepare
    End If

    ' Else, all is well - Control system is running normally
    Return DelayValue.NormalRunning
  End Function

  Public Enum DelayValue
    NormalRunning
    Boilout
    Correction
    'StepOverrun
    Sample
    Load
    Unload
    CheckPH
    [Operator]
    Fill
    Drain
    Heating
    Cooling
    Paused
    MachineError
    ExpansionAdd
    KitchenAdd
    ReserveTransfer
    DispenseWaitReady
    DispenseWaitResultDyes
    DipsenseWaitResultChems
    DispenseWaitResultCombined
    LocalPrepare
    Tank1Operator
    Tank1DispenseError
    IsDelayed
    Sleeping = 101
  End Enum

End Class
