' A&E Then Platform
' Version 2024-09-17

Imports Utilities.Translations

Public Class PumpControl : Inherits MarshalByRefObject

  Public Enum EState
    Off
    Interlock
    CoolingWater
    Start
    Running
    [Stop]
  End Enum
  Property State As EState
  Property StateString As String
  Property PumpLevel As Boolean
  Property PumpEnable As Boolean

  Property PumpRunning As Boolean
  Property PumpStartCount As Integer

  ' Flow Direction States
  Public Enum EStateFlowDirection
    Off
    InToOutSwitch
    InToOutAccel
    InToOutError
    InToOut
    InToOutDecel
    OutToInSwitch
    OutToInAccel
    OutToInError    
    OutToIn
    OutToInDecel
  End Enum
  Property StateFlow As EStateFlowDirection
  Property StateFlowLast As EStateFlowDirection
  Property StateFlowString As String
  Property PumpInToOutTime As Integer
  Property PumpOutToInTime As Integer
  Property PumpHoldInToOut As Boolean
  Property PumpHoldOutToIn As Boolean

  Property CountInToOut As Integer
  Property CountOutToIn As Integer


  ' Timers
  Property Timer As New Timer
  Property TimerFlow As New Timer
  Property TimerFlowReversal As New Timer
  Property TimerFlowAdjust As New Timer
  Property TimerPumpEnabled As New Timer

  Private ReadOnly timerPumpLevel_ As New Timer
  Private ReadOnly timerPumpLevelLost_ As New Timer
  Private ReadOnly timerPumpRunningLost_ As New Timer

#Region " PARAMETERS "

  <Parameter(10, 60), Category("Pump Control"),
  Description("Time, in seconds, to delay for all pump interlocks to be true before enabling pump."),
  TranslateDescription("es", "Tiempo, en segundos, para retrasar por todos los software para ser verdadero antes de activar la bomba de la bomba.")>
  Public Parameters_PumpEnableTime As Integer

  <Parameter(0, 1000), Category("Pump Control"),
  Description("Minimum Machine Level, in tenths, for pump to operate."),
  TranslateDescription("es", "Mínimo nivel de la máquina, en décimas, para funcionamiento.")>
  Public Parameters_PumpMinimumLevel As Integer

  <Parameter(5, 15), Category("Pump Control"),
  Description("Time, in seconds, Vessel Level must be above Pump Minimum Vessel Level before allowing Pump to run."),
  TranslateDescription("es", "Tiempo, en segundos, nivel de buque debe estar sobre el nivel del recipiente mínimo de la bomba antes de permitir que la bomba funcione.")>
  Public Parameters_PumpMinimumLevelTime As Integer

  <Parameter(0, 1000), Category("Pump Control"),
  Description("Time over which to accelerate the pump to the default speed, in seconds."),
  TranslateDescription("es", "Tiempo más que acelerar la bomba a la velocidad por defecto, en segundos.")>
  Public Parameters_PumpAccelerationTime As Integer

  <Parameter(0, 9999), Category("Pump Control"),
  Description("The time in seconds for the flow reverse valve to change positions."),
  TranslateDescription("es", "El tiempo en segundos para la válvula de reversa para cambiar de posición.")>
  Public Parameters_FlowReverseSwitchingTime As Integer

  <Parameter(0, 1000), Category("Pump Control"),
  Description("Delay adjusting pump speed output after this many seconds."),
  TranslateDescription("es", "Ajuste de salida de velocidad de la bomba después de tantos segundos de retraso.")>
  Public Parameters_FlowAdjustTime As Integer

  <Parameter(0, 1000), Category("Pump Control"),
  Description("During a DP command, the pump speed output adjustment is proportional to this factor multiplied by the DP error."),
  TranslateDescription("es", "Durante un comando de DP, el ajuste de la salida de velocidad de la bomba es proporcional a este factor multiplicado por el error de DP.")>
  Public Parameters_DPAdjustFactor As Integer

  <Parameter(0, 1000), Category("Pump Control"),
  Description("During a DP command, this is the maximum change in pump speed each time it is changed, in tenths percent."),
  TranslateDescription("es", "Durante un comando DP, esto es el cambio máximo en la velocidad de la bomba cada vez que se cambia, en por ciento de los décimos.")>
  Public Parameters_DPAdjustMaximum As Integer

  <Parameter(0, 1000), Category("Pump Control"),
  Description("Maximum desired differential pressure, in tenths."),
  TranslateDescription("es", "Máximo había deseado de presión diferencial, en décimas.")>
  Public Parameters_MaxDifferentialPressure As Integer

  <Parameter(0, 1000), Category("Pump Control"),
  Description("Pump speed, in tenths, when filling the machine.  Disable with a value of '0'."),
  TranslateDescription("es", "Velocidad de la bomba, en décimas, al llenar la máquina.  Desactivar con un valor de '0'.")>
  Public Parameters_PumpSpeedFillActive As Integer

  '  <Parameter(0, 1000), Category("Pump Control"),
  '  Description(""),
  '  TranslateDescription("es", "")>
  '  Public Parameters_FlowReverseTime As Integer

  <Parameter(0, 9999), Category("Pump Control"),
  Description("Time, in seconds, to delay for pump deceleration, before reversing flow direction."),
  TranslateDescription("es", "Tiempo, en segundos, para retrasar por la desaceleración de la bomba, antes de invertir la dirección del flujo.")>
  Public Parameters_PumpSpeedDecelTime As Integer

  <Parameter(0, 1000), Category("Pump Control"),
  Description("Pump speed, in tenths, when reversing the flow."),
  TranslateDescription("es", "Velocidad de la bomba, en décimas, al invertir el flujo.")>
  Public Parameters_PumpSpeedReversal As Integer

  <Parameter(0, 1000), Category("Pump Control"),
  Description("Time, in seconds, to slow the pump before switching flow directions."),
  TranslateDescription("es", "Tiempo, en segundos, para frenar la bomba antes de cambiar las direcciones de flujo.")>
  Public Parameters_PumpSpeedAdjustTime As Integer

  <Parameter(0, 1000), Category("Pump Control"),
  Description("The maximum allowed pump speed, in tenths.  Used to prevent damage to packages.  Set to '0' to disable."),
  TranslateDescription("es", "El máximo permitió la velocidad de la bomba, en décimas.  Utilizado para evitar daños a los paquetes.  Establece en '0' para desactivar.")>
  Public Parameters_PumpSpeedMaximum As Integer

  <Parameter(0, 1000), Category("Pump Control"),
  Description("Minimum pump speed output when running, in tenths."),
  TranslateDescription("es", "Velocidad de la bomba por defecto, en décimas.")>
  Public Parameters_PumpSpeedMinimum As Integer

  <Parameter(0, 1000), Category("Pump Control"),
  Description("Default pump speed, in tenths."),
  TranslateDescription("es", "Velocidad mínima de salida cuando se ejecuta, en décimas.")>
  Public Parameters_PumpSpeedDefault As Integer

  <Parameter(0, 1000), Category("Pump Control"),
  Description("Pump speed, in tenths, when starting."),
  TranslateDescription("es", "Velocidad de la bomba, en décimas, al comenzar.")>
  Public Parameters_PumpSpeedStart As Integer

  <Parameter(0, 1000), Category("Pump Control"),
  Description("Pump speed, in tenths, when rinsing.  Disable with a value of '0'."),
  TranslateDescription("es", "Velocidad de la bomba, en décimas, cuando enjuague.")>
  Public Parameters_PumpSpeedRinse As Integer

  <Parameter(0, 1000), Category("Pump Control"),
  Description("Pump speed, in tenths, when draining."),
  TranslateDescription("es", "Velocidad de la bomba, en décimas, al drenar.")>
  Public Parameters_PumpSpeedDrain As Integer

  '  <Parameter(0, 1000), Category("Pump Control"),
  '  Description(""),
  '  TranslateDescription("es", "")>
  '  Public Parameters_PumpOnOffCountTooHigh As Integer

  <Parameter(0, 60), Category("Pump Control"),
  Description("Time, in seconds, to delay starting pump to allow pump seal cooling."),
  TranslateDescription("es", "Tiempo en segundos, para retrasar el encender la bomba para permitir el sello de la bomba de refrigeración.")>
  Public Parameters_PumpOnCoolingDelay As Integer

  '  <Parameter(0, 100000), Category("Pump Control"),
  '  Description("Set to max pressure allowed at pump output.  Used to limit pump speed, especially during high pressure states.  Used with Machine Pressure and Differential Pressure feedback to prevent unsafe conditions.  Set to '0' to disable."),
  '  TranslateDescription("es", "")>
  '  Public Parameters_PumpPressureMax As Integer

#End Region

  'Keep a local reference to the control code object for convenience
  Private ReadOnly controlCode As ControlCode

  Public Sub New(ByVal controlCode As ControlCode)
    Me.controlCode = controlCode
  End Sub

  'Run the state machine
  Public Sub Run()
    With controlCode

      ' Determine if PumpLevel requirement is met
      If (.MachineLevel < Parameters_PumpMinimumLevel) Then timerPumpLevel_.Seconds = Parameters_PumpMinimumLevelTime + 1

      ' PumpLevel met
      If timerPumpLevel_.Finished OrElse .FI.IsActive OrElse .RT.IsActive OrElse .DR.IsActive Then
        PumpLevel = True
        timerPumpLevelLost_.Seconds = Parameters_PumpMinimumLevelTime

        ' Reset alarms (TODO Confirm)
        If isAlarmLevelLost_ Then isAlarmLevelLost_ = False
      End If

      ' Monitor while PumpLevel met and turn off after short level delay
      If PumpLevel Then
        ' Delay if pump level is lost 
        If timerPumpLevelLost_.Finished AndAlso Not .RT.IsActive Then
          PumpLevel = False
          timerPumpLevel_.Seconds = Parameters_PumpMinimumLevelTime + 1

          'todo check this
          If PumpRequest Then isAlarmLevelLost_ = True
        End If
      End If

      ' Update pump enable timer
      If Not PumpLevel Then TimerPumpEnabled.Seconds = MinMax(Parameters_PumpEnableTime, 5, 60)
      If Not .MachineClosed Then TimerPumpEnabled.Seconds = MinMax(Parameters_PumpEnableTime, 5, 60)
      If Not .Parent.IsProgramRunning Then TimerPumpEnabled.Seconds = MinMax(Parameters_PumpEnableTime, 5, 60)
      If .EStop Then TimerPumpEnabled.Seconds = MinMax(Parameters_PumpEnableTime, 5, 60)

      ' Disable pump during Fill & Drain Steps
      '   If .FI.IsActive OrElse .DR.IsActive Then   PumpEnabledTimer.Seconds = PumpEnableTime

      PumpEnable = TimerPumpEnabled.Finished

      ' Foward/Reverse Hold 
      PumpHoldInToOut = (.FR.IsOn AndAlso (.FR.OutToInTime = 0)) OrElse
                        (.FC.IsOn AndAlso (.FC.OutToInTime = 0)) OrElse
                        .FI.IsActive OrElse .DR.IsActive OrElse .HD.IsOn OrElse
                        .RC.IsActive OrElse .RH.IsActive OrElse .RI.IsActive OrElse
                        .RT.IsActive

      PumpHoldOutToIn = ((.FR.IsOn AndAlso (.FR.InToOutTime = 0)) OrElse (.FC.IsOn AndAlso (.FC.OutToInTime = 0))) _
                        AndAlso Not PumpHoldInToOut

      ' Set main pump stopped timer
      PumpRunning = .IO.PumpRunning
      If PumpRunning Then timerPumpRunningLost_.Seconds = 8
      '   If Not PumpRunning Then TimerFlow.Seconds = Parameters_PumpAccelerationTime

      'Check pump switch 
      CheckPumpSwitch()

      '==============================================================================
      'PUMP STATE CONTROL
      '  Remember the state at the beginning and loop until it doesn't change
      '  This will force the state machine to through all state changes before exiting
      '==============================================================================
      Static startState As EState
      Do
        startState = State
        Select Case State
          Case EState.Off
            DatePumpStart = Nothing
            StateString = Translate("Pump Off")

          Case EState.Interlock
            DatePumpStart = Nothing
            StateString = Translate("Pump Interlocked") & TimerPumpEnabled.ToString(1)
            Timer.Seconds = MinMax(Parameters_PumpAccelerationTime, 5, 300)
            timerFlow.Seconds = MinMax(Parameters_PumpAccelerationTime, 5, 300)

            If TimerPumpEnabled.Finished Then
              State = EState.CoolingWater
              Timer.Seconds = MinMax(Parameters_PumpOnCoolingDelay, 1, 60)
            End If

          Case EState.CoolingWater
            StateString = Translate("Cooling Delay") & Timer.ToString(1)
            If Timer.Finished Then
              ' Start the pump
              State = EState.Start
              Timer.Seconds = MinMax(Parameters_PumpAccelerationTime, 5, 300)

              If PumpHoldOutToIn OrElse WasReverse Then
                StateFlow = EStateFlowDirection.OutToInSwitch
              Else
                StateFlow = EStateFlowDirection.InToOutSwitch
              End If
              ' Update reversal switch alarm time
              TimerFlowReversal.Seconds = MinMax(Parameters_FlowReverseSwitchingTime, 5, 300)
            End If
            If Not PumpEnabled Then
              State = EState.Interlock
              StateFlow = EStateFlowDirection.Off
            End If

          Case EState.Start
            StateString = Translate("Pump Starting")
            If PumpRunning Then
              StateString = "Accelerating " & Timer.ToString
            Else : Timer.Seconds = MinMax(Parameters_PumpAccelerationTime, 1, 300)
            End If
            If Timer.Finished Then
              DatePumpStart = Date.Now.ToLocalTime
              State = EState.Running
            End If
            If Not PumpEnabled Then
              State = EState.Interlock
              StateFlow = EStateFlowDirection.Off
            End If

          Case EState.Running
            StateString = Translate("Pump Running")
            If Not PumpRunning Then
              TimerPumpEnabled.Seconds = 5
              State = EState.Interlock
            End If
            If Not PumpEnabled Then
              State = EState.Interlock
              StateFlow = EStateFlowDirection.Off
            End If

          Case EState.Stop
            StateString = Translate("Pump Stopping") & timerPumpRunningLost_.ToString(1)
            If timerPumpRunningLost_.Finished Then
              If Not PumpRunning Then
                State = EState.Off
                StateFlow = EStateFlowDirection.Off
              End If
            End If

        End Select
      Loop Until (startState = State)


      '==============================================================================
      ' FLOW REVERSAL & SPEED CONTROL
      '==============================================================================
      Static pumpSpeedLastInOut As Short
      Static pumpspeedLastOutIn As Short
      Select Case StateFlow
        Case EStateFlowDirection.Off
          StateFlowString = ""
          pumpSpeedLastInOut = 0
          pumpspeedLastOutIn = 0

        Case EStateFlowDirection.InToOutSwitch
          If TimerFlow.Finished OrElse (Parameters_FlowReverseSwitchingTime = 0) Then
            StateFlow = EStateFlowDirection.InToOutAccel
            TimerFlow.Seconds = MinMax(Parameters_PumpAccelerationTime, 5, 300)
            CountInToOut += 1
          End If
          ' Update Pump Speed Output
          pumpSpeedOutput_ = CType(Parameters_PumpSpeedReversal, Short)
          StateFlowString = ("I-O") & (" ") & Translate("Switching") & TimerFlow.ToString(1)

        Case EStateFlowDirection.InToOutAccel
          ' Acceleration Timer Completed
          If TimerFlow.Finished Then
            ' Next working speed
            If pumpSpeedLastInOut > 0 Then
              PumpSpeedInOut = pumpSpeedLastInOut
              TimerFlowAdjust.Seconds = Parameters_FlowAdjustTime
            Else
              ' Use current active value 
              PumpSpeedInOut = pumpSpeedOutput_
            End If
            ' Update Flow Reversal Time
            StateFlow = EStateFlowDirection.InToOut
            TimerFlow.Minutes = PumpInToOutTime
          End If
          ' Update Pump Speed Output
          pumpSpeedOutput_ = CType(Parameters_PumpSpeedReversal, Short)
          StateFlowString = ("I-O") & (" ") & Translate("Accel") & TimerFlow.ToString(1)

        Case EStateFlowDirection.InToOut
          ' Check to see if forward hold command is active or if flow reversal is set to only Forward
          If PumpHoldInToOut Then TimerFlow.Minutes = PumpInToOutTime
          ' Pump Speed Adjustment
          If TimerFlowAdjust.Finished Then
            PumpSpeedInOut = CalculateSpeedInOut(PumpSpeedInOut)
          End If
          ' Update Pump Speed Output
          pumpSpeedOutput_ = CType(PumpSpeedInOut, Short)
          ' Update last speed values
          pumpSpeedLastInOut = CShort(PumpSpeedInOut)
          ' Flow Timer finished
          If TimerFlow.Finished AndAlso (PumpOutToInTime > 0) Then
            StateFlow = EStateFlowDirection.InToOutDecel
            TimerFlow.Seconds = Parameters_PumpSpeedDecelTime
          End If
          StateFlowString = ("I-O") & (" ") & TimerFlow.ToString & (" ") & (pumpSpeedOutput_ / 10).ToString("#0.0") & ("%")

        Case EStateFlowDirection.InToOutDecel
          ' Flow Direction Timer Completed
          If TimerFlow.Finished Then
            ' Switch Flow Directions
            StateFlow = EStateFlowDirection.OutToInSwitch
            TimerFlow.Seconds = MinMax(Parameters_FlowReverseSwitchingTime, 5, 300)
          End If
          ' Stepped into a hold I/O state, accel back in I/O (RI/RT/etc)
          If PumpHoldInToOut Then
            StateFlow = EStateFlowDirection.InToOutAccel
            TimerFlow.Seconds = MinMax(Parameters_PumpAccelerationTime, 5, 300)
          End If
          ' Update Pump Speed Output - if speed greater than reversal parameter speed
          If pumpSpeedOutput_ > CType(Parameters_PumpSpeedReversal, Short) Then
            pumpSpeedOutput_ = CType(Parameters_PumpSpeedReversal, Short)
          End If
          StateFlowString = ("I-O") & (" ") & Translate("Decel") & TimerFlow.ToString(1)

        Case EStateFlowDirection.OutToInSwitch
          If TimerFlow.Finished OrElse (Parameters_FlowReverseSwitchingTime = 0) Then
            StateFlow = EStateFlowDirection.OutToInAccel
            TimerFlow.Seconds = MinMax(Parameters_PumpAccelerationTime, 5, 300)
            CountOutToIn += 1
          End If
          ' Update Pump Speed Output - if speed greater than reversal parameter speed
          If pumpSpeedOutput_ > CType(Parameters_PumpSpeedReversal, Short) Then
            pumpSpeedOutput_ = CType(Parameters_PumpSpeedReversal, Short)
          End If
          StateFlowString = ("O-I") & (" ") & Translate("Switching") & TimerFlow.ToString(1)

        Case EStateFlowDirection.OutToInAccel
          ' Acceleration Timer Completed
          If TimerFlow.Finished Then
            ' Next working speed
            If pumpspeedLastOutIn > 0 Then
              PumpSpeedOutIn = pumpspeedLastOutIn
            Else
              ' Use current active value 
              PumpSpeedOutIn = pumpSpeedOutput_
            End If
            ' Update Flow Reversal Time
            StateFlow = EStateFlowDirection.OutToIn
            TimerFlow.Minutes = PumpOutToInTime
            '  TimerFlow.Seconds = PumpTimeOutIn
            TimerFlowAdjust.Seconds = Parameters_FlowAdjustTime
          End If
          ' Update Pump Speed Output
          pumpSpeedOutput_ = CType(Parameters_PumpSpeedReversal, Short)
          StateFlowString = ("O-I") & (" ") & Translate("Accel") & TimerFlow.ToString(1)

        Case EStateFlowDirection.OutToIn
          ' Check to see if flow reversal is set to only reverse
          If PumpHoldOutToIn Then TimerFlow.Minutes = PumpOutToInTime
          ' Pump Speed Adjustment
          If TimerFlowAdjust.Finished Then
            PumpSpeedOutIn = CalculateSpeedOutIn(PumpSpeedOutIn)
          End If
          ' Flow Timer finished orelse Need to switch for command (RI/RT/etc)
          If (TimerFlow.Finished AndAlso (PumpInToOutTime > 0)) OrElse PumpHoldInToOut Then
            StateFlow = EStateFlowDirection.OutToInDecel
            TimerFlow.Seconds = Parameters_PumpSpeedDecelTime
          End If
          ' Update Pump Speed Output
          pumpSpeedOutput_ = CType(PumpSpeedOutIn, Short)
          ' Update last speed values
          pumpspeedLastOutIn = CShort(PumpSpeedOutIn)
          StateFlowString = ("O-I") & (" ") & TimerFlow.ToString & (" ") & (pumpSpeedOutput_ / 10).ToString("#0.0") & ("%")

        Case EStateFlowDirection.OutToInDecel
          ' Flow Direction Timer Completed
          If TimerFlow.Finished Then
            ' Next working speed
            If pumpSpeedLastInOut > 0 Then
              PumpSpeedInOut = pumpSpeedLastInOut
            Else
              ' Use current active value 
              PumpSpeedInOut = pumpSpeedOutput_
            End If
            ' Switch InToOut
            StateFlow = EStateFlowDirection.InToOutSwitch
            TimerFlow.Seconds = MinMax(Parameters_FlowReverseSwitchingTime, 5, 300)
            ' Update FlowAdjust Timer
            TimerFlowAdjust.Seconds = Parameters_FlowAdjustTime
          End If
          ' Update Pump Speed Output - if speed greater than reversal parameter speed
          If pumpSpeedOutput_ > CType(Parameters_PumpSpeedReversal, Short) Then
            pumpSpeedOutput_ = CType(Parameters_PumpSpeedReversal, Short)
          End If
          StateFlowString = ("O-I") & (" ") & Translate("Decel") & TimerFlow.ToString(1)

      End Select

      ' Update last direction
      StateFlowLast = StateFlow

      ' Pump no longer running
      If Not (IsStarting OrElse IsRunning) AndAlso (StateFlow <> EStateFlowDirection.Off) Then
        StateFlow = EStateFlowDirection.Off
      End If

    End With
  End Sub

  Private Function CalculateSpeedInOut(ByVal pumpSpeedCurrent As Integer) As Integer
    Dim returnValue As Integer = MinMax(Parameters_PumpSpeedStart, 0, Parameters_PumpSpeedMaximum)
    With controlCode
      ' Working Variables
      Dim dpError As Integer, dpAdjust As Integer
      Dim flError As Integer, flAdjust As Integer

      ' DP command is active
      If .DP.IsOn AndAlso (.DP.InsideOutPress > 0) Then
        ' Don't let DP exceed max DP, if enabled
        If (.PackageDifferentialPressure > Parameters_MaxDifferentialPressure) And (Parameters_MaxDifferentialPressure > 0) Then
          dpError = Parameters_MaxDifferentialPressure - .PackageDifferentialPressure
          dpAdjust = MulDiv(dpError, Parameters_DPAdjustFactor, 100)
          If dpAdjust < -Parameters_DPAdjustMaximum Then dpAdjust = -Parameters_DPAdjustMaximum
          If dpAdjust > Parameters_DPAdjustMaximum Then dpAdjust = Parameters_DPAdjustMaximum
          ' Check that new pump speed is not greater than maximum value
          If (pumpSpeedCurrent + dpAdjust) >= Parameters_PumpSpeedMaximum Then
            ' Use maximum value 
            returnValue = Parameters_PumpSpeedMaximum
          Else
            ' Update value
            returnValue = (pumpSpeedCurrent + dpAdjust)
          End If
        Else
          ' Adjust pump speed based on Differential pressure
          dpError = .DP.InsideOutPress - .PackageDifferentialPressure
          dpAdjust = MulDiv(dpError, Parameters_DPAdjustFactor, 100)
          If dpAdjust < -Parameters_DPAdjustMaximum Then dpAdjust = -Parameters_DPAdjustMaximum
          If dpAdjust > Parameters_DPAdjustMaximum Then dpAdjust = Parameters_DPAdjustMaximum
          ' Check that new pump speed is not greater than maximum value
          If (pumpSpeedCurrent + dpAdjust) >= Parameters_PumpSpeedMaximum Then
            ' Use maximum value 
            returnValue = Parameters_PumpSpeedMaximum
          Else
            ' Update value
            returnValue = (pumpSpeedCurrent + dpAdjust)
          End If
        End If
      End If

      ' FL command is active (AND ENABLED)
      If .FL.IsOn AndAlso (.FL.InToOutFlow > 0) AndAlso (.Parameters.FlowRateMachineRange > 0) AndAlso (.Parameters.FlowRateMachineMax > 0) Then
        ' Adjust pump speed based on flowmeter rate feedback - Don't let DP exceed max DP, if enabled
        If (.PackageDifferentialPressure > Parameters_MaxDifferentialPressure) And (Parameters_MaxDifferentialPressure > 0) Then
          dpError = Parameters_MaxDifferentialPressure - .PackageDifferentialPressure
          dpAdjust = MulDiv(dpError, Parameters_DPAdjustFactor, 100)
          If dpAdjust < -Parameters_DPAdjustMaximum Then dpAdjust = -Parameters_DPAdjustMaximum
          If dpAdjust > Parameters_DPAdjustMaximum Then dpAdjust = Parameters_DPAdjustMaximum
          ' Check that new pump speed is not greater than maximum value
          If (pumpSpeedCurrent + dpAdjust) >= Parameters_PumpSpeedMaximum Then
            ' Use maximum value 
            returnValue = Parameters_PumpSpeedMaximum
          Else
            ' Update value
            returnValue = (pumpSpeedCurrent + dpAdjust)
          End If
        Else
          ' Adjust pump speed based on machine flowrate feedback
          flError = .FL.InToOutFlow - .MachineFlowRatePv
          flAdjust = MulDiv(flError, .Parameters.FLAdjustFactor, 100)
          If flAdjust < - .Parameters.FLAdjustMaximum Then flAdjust = - .Parameters.FLAdjustMaximum
          If flAdjust > .Parameters.FLAdjustMaximum Then flAdjust = .Parameters.FLAdjustMaximum
          ' Check that new pump speed is not greater than maximum value
          If (pumpSpeedCurrent + flAdjust) >= Parameters_PumpSpeedMaximum Then
            ' Use maximum value 
            returnValue = Parameters_PumpSpeedMaximum
          Else
            ' Update value
            returnValue = (pumpSpeedCurrent + flAdjust)
          End If
        End If
      End If

      ' FP command is active
      If .FP.IsOn AndAlso (.FP.InToOutPercent > 0) Then
        ' Don't let DP exceed max DP, if enabled
        If (.PackageDifferentialPressure > Parameters_MaxDifferentialPressure) And (Parameters_MaxDifferentialPressure > 0) Then
          dpError = Parameters_MaxDifferentialPressure - .PackageDifferentialPressure
          dpAdjust = MulDiv(dpError, Parameters_DPAdjustFactor, 100)
          If dpAdjust < -Parameters_DPAdjustMaximum Then dpAdjust = -Parameters_DPAdjustMaximum
          If dpAdjust > Parameters_DPAdjustMaximum Then dpAdjust = Parameters_DPAdjustMaximum
          ' Check that new pump speed is not greater than maximum value
          If (pumpSpeedCurrent + dpAdjust) >= Parameters_PumpSpeedMaximum Then
            ' Use maximum value
            returnValue = Parameters_PumpSpeedMaximum
          Else
            ' Update value
            returnValue = (pumpSpeedCurrent + dpAdjust)
          End If
        Else
          ' set pump speed based on flow percent defined
          If .FP.InToOutPercent > Parameters_PumpSpeedMaximum Then
            ' Use maximum value
            returnValue = Parameters_PumpSpeedMaximum
          Else
            ' Update Value
            returnValue = .FP.InToOutPercent
          End If
          ' Check that new pump speed is not greater than maximum value
          'If (pumpSpeedCurrent + dpAdjust) >= Parameters_PumpSpeedMaximum Then
          '  ' Use maximum value and reset timer adjust
          '  returnValue = Parameters_PumpSpeedMaximum
          'Else
          '  ' Update value
          '  returnValue = (pumpSpeedCurrent + dpAdjust)
          'End If
        End If
      End If

      ' Pump Speed Override Commands Active
      If .FI.IsActive AndAlso (Parameters_PumpSpeedFillActive > 0) Then returnValue = MinMax(Parameters_PumpSpeedFillActive, 0, Parameters_PumpSpeedMaximum)
      If .DR.IsActive AndAlso (Parameters_PumpSpeedDrain > 0) Then returnValue = MinMax(Parameters_PumpSpeedDrain, 0, Parameters_PumpSpeedMaximum)
      If .RC.IsActive OrElse .RH.IsActive OrElse .RI.IsActive OrElse .RT.IsActive Then
        If (Parameters_PumpSpeedRinse > 0) Then returnValue = MinMax(Parameters_PumpSpeedRinse, 0, Parameters_PumpSpeedMaximum)
      End If

      ' Pump Speed Minimum
      If (Parameters_PumpSpeedMinimum > 0) AndAlso returnValue < Parameters_PumpSpeedMinimum Then returnValue = Parameters_PumpSpeedMinimum

      ' Reset Flow Adjust Timer
      TimerFlowAdjust.Seconds = Parameters_FlowAdjustTime

      ' Return final value
      Return returnValue
    End With
  End Function

  Private Function CalculateSpeedOutIn(ByVal pumpSpeedCurrent As Integer) As Integer
    Dim returnValue As Integer = MinMax(Parameters_PumpSpeedStart, Parameters_PumpSpeedMinimum, Parameters_PumpSpeedMaximum)
    With controlCode
      ' Working Variables
      Dim dpError As Integer, dpAdjust As Integer
      Dim flError As Integer, flAdjust As Integer

      ' DP command is active
      If .DP.IsOn AndAlso (.DP.OutsideInPress > 0) Then
        ' Don't let DP exceed max DP, if enabled
        If (.PackageDifferentialPressure > Parameters_MaxDifferentialPressure) And (Parameters_MaxDifferentialPressure > 0) Then
          dpError = Parameters_MaxDifferentialPressure - .PackageDifferentialPressure
          dpAdjust = MulDiv(dpError, Parameters_DPAdjustFactor, 100)
          If dpAdjust < -Parameters_DPAdjustMaximum Then dpAdjust = -Parameters_DPAdjustMaximum
          If dpAdjust > Parameters_DPAdjustMaximum Then dpAdjust = Parameters_DPAdjustMaximum
          ' Check that new pump speed is not greater than maximum value
          If (pumpSpeedCurrent + dpAdjust) >= Parameters_PumpSpeedMaximum Then
            ' Use maximum value 
            returnValue = Parameters_PumpSpeedMaximum
          Else
            ' Update value
            returnValue = (pumpSpeedCurrent + dpAdjust)
          End If
        Else
          ' Adjust pump speed based on Differential pressure
          dpError = .DP.OutsideInPress - .PackageDifferentialPressure
          dpAdjust = MulDiv(dpError, Parameters_DPAdjustFactor, 100)
          If dpAdjust < -Parameters_DPAdjustMaximum Then dpAdjust = -Parameters_DPAdjustMaximum
          If dpAdjust > Parameters_DPAdjustMaximum Then dpAdjust = Parameters_DPAdjustMaximum
          ' Check that new pump speed is not greater than maximum value
          If (pumpSpeedCurrent + dpAdjust) >= Parameters_PumpSpeedMaximum Then
            ' Use maximum value 
            returnValue = Parameters_PumpSpeedMaximum
          Else
            ' Update value
            returnValue = (pumpSpeedCurrent + dpAdjust)
          End If
        End If
      End If

      ' FL command is active (AND ENABLED)
      If .FL.IsOn AndAlso (.FL.OutToInFlow > 0) AndAlso (.Parameters.FlowRateMachineRange > 0) AndAlso (.Parameters.FlowRateMachineMax > 0) Then
        ' Adjust pump speed based on flowmeter rate feedback - Don't let DP exceed max DP, if enabled
        If (.PackageDifferentialPressure > Parameters_MaxDifferentialPressure) And (Parameters_MaxDifferentialPressure > 0) Then
          dpError = Parameters_MaxDifferentialPressure - .PackageDifferentialPressure
          dpAdjust = MulDiv(dpError, Parameters_DPAdjustFactor, 100)
          If dpAdjust < -Parameters_DPAdjustMaximum Then dpAdjust = -Parameters_DPAdjustMaximum
          If dpAdjust > Parameters_DPAdjustMaximum Then dpAdjust = Parameters_DPAdjustMaximum
          ' Check that new pump speed is not greater than maximum value
          If (pumpSpeedCurrent + dpAdjust) >= Parameters_PumpSpeedMaximum Then
            ' Use maximum value 
            returnValue = Parameters_PumpSpeedMaximum
          Else
            ' Update value
            returnValue = (pumpSpeedCurrent + dpAdjust)
          End If
        Else
          ' Adjust pump speed based on machine flowrate feedback
          flError = .FL.OutToInFlow - .MachineFlowRatePv
          flAdjust = MulDiv(flError, .Parameters.FLAdjustFactor, 100)
          If flAdjust < - .Parameters.FLAdjustMaximum Then flAdjust = - .Parameters.FLAdjustMaximum
          If flAdjust > .Parameters.FLAdjustMaximum Then flAdjust = .Parameters.FLAdjustMaximum
          ' Check that new pump speed is not greater than maximum value
          If (pumpSpeedCurrent + flAdjust) >= Parameters_PumpSpeedMaximum Then
            ' Use maximum value 
            returnValue = Parameters_PumpSpeedMaximum
          Else
            ' Update value
            returnValue = (pumpSpeedCurrent + flAdjust)
          End If
        End If
      End If

      ' FP command is active
      If .FP.IsOn AndAlso (.FP.OutToInPercent > 0) Then
        ' Don't let DP exceed max DP, if enabled
        If (.PackageDifferentialPressure > Parameters_MaxDifferentialPressure) And (Parameters_MaxDifferentialPressure > 0) Then
          dpError = Parameters_MaxDifferentialPressure - .PackageDifferentialPressure
          dpAdjust = MulDiv(dpError, Parameters_DPAdjustFactor, 100)
          If dpAdjust < -Parameters_DPAdjustMaximum Then dpAdjust = -Parameters_DPAdjustMaximum
          If dpAdjust > Parameters_DPAdjustMaximum Then dpAdjust = Parameters_DPAdjustMaximum
          ' Check that new pump speed is not greater than maximum value
          If (pumpSpeedCurrent + dpAdjust) >= Parameters_PumpSpeedMaximum Then
            ' Use maximum value
            returnValue = Parameters_PumpSpeedMaximum
          Else
            ' Update value
            returnValue = (pumpSpeedCurrent + dpAdjust)
          End If
        Else
          ' set pump speed based on flow percent defined
          If .FP.OutToInPercent > Parameters_PumpSpeedMaximum Then
            ' Use maximum value
            returnValue = Parameters_PumpSpeedMaximum
          Else
            ' Update Value
            returnValue = .FP.OutToInPercent
          End If
          ' Check that new pump speed is not greater than maximum value
          'If (pumpSpeedCurrent + dpAdjust) >= Parameters_PumpSpeedMaximum Then
          '  ' Use maximum value and reset timer adjust
          '  returnValue = Parameters_PumpSpeedMaximum
          'Else
          '  ' Update value
          '  returnValue = (pumpSpeedCurrent + dpAdjust)
          'End If
        End If
      End If

      ' Pump Speed Override Commands Active
      '   None for Out-To-In flow

      ' Pump Speed Minimum
      If (Parameters_PumpSpeedMinimum > 0) AndAlso returnValue < Parameters_PumpSpeedMinimum Then returnValue = Parameters_PumpSpeedMinimum

      ' Reset Flow Adjust Timer
      TimerFlowAdjust.Seconds = Parameters_FlowAdjustTime

      ' Return final value
      Return returnValue

    End With
  End Function

  ' Start Main Pump
  Public Sub StartAuto()
    With controlCode
      DatePumpStartAuto = Date.Now.ToLocalTime
      PumpRequest = True
      PumpStartCount += 1
      '  If Not .IO.SwPumpInAuto Then .Parent.Signal = "Pump Switch Off "
      If Not (IsRunning OrElse IsStarting) Then
        State = EState.Interlock
      End If
    End With
  End Sub

  Public Sub StartManual()
    With controlCode
      'Disable manual start if no program is running and also during drain
      If (Not .Parent.IsProgramRunning) OrElse .DR.IsOn Then Exit Sub
      If Not PumpEnable Then Exit Sub

      PumpStartCount += 1

      DatePumpStartManual = Date.Now.ToLocalTime
      If TimerPumpEnabled.Finished Then
        PumpRequest = True
        State = EState.Interlock
      End If
    End With
  End Sub

  Public Sub StopMainPump()
    'Pump Running - stop pump after delay
    PumpRequest = False
    PumpStartCount = 0
    If State <> EState.Off Then State = EState.Off
  End Sub

  Private Sub CheckPumpSwitch()                                               ' NOTE: No Pump Switch Installed on these machines
    ' Check Pump Switch Position
    ' Static previousPumpSwitch As Boolean

    ' Catch leading edge and issue manual start command
    '  If PumpSwitch AndAlso (Not previousPumpSwitch) Then StartManual()

    ' Stop Main Pump if pump switch is off - regardless of last position
    '   If (Not PumpSwitch) Then StopMainPump()

    ' Remember the value for next time
    '   previousPumpSwitch = PumpSwitch
  End Sub

  'Stop circulation
  Public Sub Cancel()
    PumpRequest = False
    If (State <> EState.Off) Then State = EState.Stop
  End Sub

  Public Sub SwitchInToOut()
    StateFlow = EStateFlowDirection.OutToInDecel
    TimerFlow.Seconds = Parameters_PumpSpeedDecelTime
  End Sub

  Public Sub SwitchOutToIn()
    StateFlow = EStateFlowDirection.InToOutDecel
    TimerFlow.Seconds = Parameters_PumpSpeedDecelTime
  End Sub

  Public Sub UpdateFlowReversalTimes(inToOutTime As Integer, outToInTime As Integer)
    Me.PumpInToOutTime = inToOutTime
    Me.PumpOutToInTime = outToInTime

    If IsInToOut Then
      TimerFlow.Minutes = inToOutTime
    ElseIf IsOutToIn Then
      TimerFlow.Minutes = outToInTime
    End If
  End Sub

#Region " PROPERTIES "

  Private pumpOffDelay_ As Integer
  Public Property PumpOffDelay() As Integer
    Get
      Return pumpOffDelay_
    End Get
    Set(ByVal value As Integer)
      pumpOffDelay_ = MinMax(value, 0, 30)
    End Set
  End Property

  Public Property PumpEnabled() As Boolean
    Get
      Return TimerPumpEnabled.Finished
    End Get
    Private Set(ByVal value As Boolean)
      If Not value Then TimerPumpEnabled.Seconds = MinMax(Parameters_PumpEnableTime, 5, 60)
    End Set
  End Property

  Private pumpSpeedInOut_ As Integer
  Public Property PumpSpeedInOut As Integer
    Get
      Return pumpSpeedInOut_
    End Get
    Private Set(value As Integer)
      'Limit pump speed if parameter is set
      If (Parameters_PumpSpeedMaximum > 0) Then
        If value > Parameters_PumpSpeedMaximum Then value = CType(Parameters_PumpSpeedMaximum, Short)
      End If

      ' Limit value to 0/1000
      pumpSpeedInOut_ = MinMax(value, 0, 1000)
    End Set
  End Property

  Private pumpSpeedOutIn_ As Integer
  Public Property PumpSpeedOutIn As Integer
    Get
      Return pumpSpeedOutIn_
    End Get
    Private Set(value As Integer)
      'Limit pump speed if parameter is set
      If (Parameters_PumpSpeedMaximum > 0) Then
        If value > Parameters_PumpSpeedMaximum Then value = CType(Parameters_PumpSpeedMaximum, Short)
      End If

      ' Limit value to 0/1000
      pumpSpeedOutIn_ = MinMax(value, 0, 1000)
    End Set
  End Property

  Private datePumpStart_ As Date
  Public Property DatePumpStart() As Date
    Get
      Return datePumpStart_
    End Get
    Set(ByVal value As Date)
      datePumpStart_ = value
    End Set
  End Property

  Private datePumpStartAuto_ As Date
  Public Property DatePumpStartAuto() As Date
    Get
      Return datePumpStartAuto_
    End Get
    Set(ByVal value As Date)
      datePumpStartAuto_ = value
    End Set
  End Property

  Private datePumpStartManual_ As Date
  Public Property DatePumpStartManual() As Date
    Get
      Return datePumpStartManual_
    End Get
    Set(ByVal value As Date)
      datePumpStartManual_ = value
    End Set
  End Property

  Private pumpRequest_ As Boolean
  Public Property PumpRequest As Boolean
    Get
      Return pumpRequest_
    End Get
    Private Set(ByVal value As Boolean)
      pumpRequest_ = value
    End Set
  End Property

  Friend ReadOnly Property PumpRunTime() As TimeSpan
    Get
      If IsRunning Then Return Date.Now.ToLocalTime.Subtract(DatePumpStart)
      Return Nothing
    End Get
  End Property

  Public ReadOnly Property PumpRunSeconds() As Integer
    Get
      If IsRunning Then Return Convert.ToInt32(Math.Min(PumpRunTime.TotalSeconds, Int32.MaxValue))
      Return 0
    End Get
  End Property


  ' Pump Alarm Status
  Private isAlarmLevelLost_ As Boolean
  Public ReadOnly Property IsAlarmLevelLost As Boolean
    Get
      Return isAlarmLevelLost_
    End Get
  End Property

  Private isAlarmPumpStartCountTooHigh_ As Boolean
  Public ReadOnly Property IsAlarmPumpStartCountTooHigh As Boolean
    Get
      Return isAlarmPumpStartCountTooHigh_
    End Get
  End Property

  ' Pump Status Properties
  Public ReadOnly Property IsActive As Boolean
    Get
      Return (State >= EState.CoolingWater) AndAlso (State <= EState.Running)
    End Get
  End Property

  Public ReadOnly Property IsRunning() As Boolean
    Get
      Return (State = EState.Running)
    End Get
  End Property

  Public ReadOnly Property IsStopped() As Boolean
    Get
      Return (State = EState.Off)
    End Get
  End Property

  Public ReadOnly Property IsStarting() As Boolean
    Get
      Return (State = EState.CoolingWater) OrElse (State = EState.Start)
    End Get
  End Property

  Public ReadOnly Property IsInToOut As Boolean
    Get
      Return (StateFlow = EStateFlowDirection.InToOutSwitch) OrElse (StateFlow = EStateFlowDirection.InToOutAccel) OrElse (StateFlow = EStateFlowDirection.InToOut)
    End Get
  End Property

  Public ReadOnly Property IsOutToIn As Boolean
    Get
      Return (StateFlow = EStateFlowDirection.OutToInSwitch) OrElse (StateFlow = EStateFlowDirection.OutToInAccel) OrElse (StateFlow = EStateFlowDirection.OutToIn)
    End Get
  End Property

  Public ReadOnly Property WasForward As Boolean
    Get
      Return (StateFlowLast = EStateFlowDirection.InToOutSwitch) OrElse (StateFlowLast = EStateFlowDirection.InToOutAccel) OrElse (StateFlowLast = EStateFlowDirection.InToOut)
    End Get
  End Property

  Public ReadOnly Property WasReverse As Boolean
    Get
      Return (StateFlowLast = EStateFlowDirection.OutToInSwitch) OrElse (StateFlowLast = EStateFlowDirection.OutToInAccel) OrElse (StateFlowLast = EStateFlowDirection.OutToIn)
    End Get
  End Property

#End Region

#Region " I/O PROPERTIES "

  Public ReadOnly Property IoFlowReverse As Boolean
    Get
      Return IsOutToIn
    End Get
  End Property

  Public ReadOnly Property IoMainPump() As Boolean
    Get
      Return (State = EState.Start) OrElse (State = EState.Running)
    End Get
  End Property

  Public ReadOnly Property IoMainPumpCooling() As Boolean
    Get
      Return (State = EState.CoolingWater) OrElse (State = EState.Start) OrElse (State = EState.Running) OrElse (State = EState.Stop)
    End Get
  End Property

  Private pumpSpeedOutput_ As Short
  Public ReadOnly Property IoMainPumpSpeed As Short
    Get

      'Limit pump speed if parameter is set
      If (Parameters_PumpSpeedMaximum > 0) Then
        If pumpSpeedOutput_ > Parameters_PumpSpeedMaximum Then pumpSpeedOutput_ = CType(Parameters_PumpSpeedMaximum, Short)
      End If

      ' Set to '0' if main pump not started
      If Not IoMainPump Then pumpSpeedOutput_ = 0

      ' Limit Pump Speed Output
      If pumpSpeedOutput_ < 0 Then pumpSpeedOutput_ = 0
      If pumpSpeedOutput_ > 1000 Then pumpSpeedOutput_ = 1000

      Return pumpSpeedOutput_
    End Get
  End Property

#End Region

End Class

