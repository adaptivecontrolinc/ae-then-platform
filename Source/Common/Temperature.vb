'American & Efird - Then Platform
' Version 2016-08-25
' - Incorporate Heat & Cool On Delay Timers into this class module, instead of leaving them in ControlCode.vb
'   Disable with Heat/Cool parameters set to '0'

' TODO - Incorporate Heat/Cool by Contacts into this class module
'         no need for two classes like A&E Then Platform vb6

' Some possibly needed:
'   Private IdleTimer As New acTimer (Timer as new acTimer)
'   Private StateTimer As New acTimer




Imports Utilities.Translations

Public Class Temperature : Inherits MarshalByRefObject

  Public Enum EMode
    None
    HeatOnly
    CoolOnly
    HeatAndCool
  End Enum

  Public Enum EState
    Off
    Idle
    Pause
    Restart

    HeatStart
    HeatVent1
    Heat
    HeatVent2

    CoolStart
    CoolVent1
    Cool
    CoolVent2

    CrashCoolStart
    CrashCoolVent
    CrashCool
    CrashCoolDone
    CrashCoolRestart
  End Enum

  Public State As EState
  Public RestartState As EState
  Public Status As String

  Public Timer As New Timer
  Public RestartSeconds As Integer

  Public StartTemp As Integer
  Public FinalTemp As Integer
  Public Gradient As Integer

  Public Enabled As Boolean
  Public EnableDelay As Integer
  Public EnabledDelayTimer As New Timer




  'Keep a local reference to the control code object for convenience
  Private ReadOnly controlCode As ControlCode

  Public Sub New(ByVal controlCode As ControlCode)
    Me.controlCode = controlCode
    Timer.Seconds = 5
    EnableDelay = 5
    EnabledDelayTimer.Seconds = EnableDelay
    State = EState.Off
    Mode = EMode.None
  End Sub

  Public Sub Start(ByVal startTemp As Integer, ByVal finalTemp As Integer, ByVal gradient As Integer)
    Me.StartTemp = startTemp
    Me.FinalTemp = finalTemp
    Me.Gradient = gradient
    Me.Mode = EMode.HeatAndCool
    StartBase()
  End Sub

  Public Sub Start(ByVal startTemp As Integer, ByVal finalTemp As Integer, ByVal gradient As Integer, ByVal mode As EMode)
    Me.StartTemp = startTemp
    Me.FinalTemp = finalTemp
    Me.Gradient = gradient
    Me.Mode = mode
    StartBase()
  End Sub

  Public Sub CrashCoolStart()
    If Not IsCrashCoolOn Then State = EState.CrashCoolStart
  End Sub

  Public Sub CrashCoolStop()
    If IsCrashCoolOn Then State = EState.CrashCoolRestart
  End Sub

  Private Sub StartBase()
    With controlCode

      'If crash cooling exit - PID will start when crash cool done
      If IsCrashCoolOn Then Exit Sub

      Select Case Mode
        Case EMode.HeatOnly
          StartHeat()
        Case EMode.CoolOnly
          StartCool()
        Case Else
          'Decide whether to start heating or cooling -> favour heating
          If StartTemp < (FinalTemp + .Parameters.HeatPropBand) Then
            StartHeat()
          Else
            StartCool()
          End If
      End Select
    End With
  End Sub

  Private Sub StartHeat()
    With controlCode
      'Use heating parameters
      LoadHeatParameters()
      'Start the pid
      Pid.Start(StartTemp, FinalTemp, Gradient)
      'Start heating 
      If State = EState.Heat Then
        'If we're already heating then just reset the mode timer
        HeatModeChangeTimer.Seconds = .Parameters.HeatModeChangeDelay + 1
      Else
        SetupRestart(EState.HeatStart, 2)
        State = EState.HeatStart
        Timer.Seconds = 2
      End If
    End With
  End Sub

  Private Sub StartCool()
    With controlCode
      'Use cooling parameters
      LoadCoolParameters()
      'Start the pid
      Pid.Start(StartTemp, FinalTemp, Gradient)
      'Start cooling
      If State = EState.Cool Then
        'If we're already cooling then just reset the mode timer
        HeatModeChangeTimer.Seconds = .Parameters.CoolModeChangeDelay + 1
      Else
        SetupRestart(EState.CoolStart, 2)
        State = EState.CoolStart
        Timer.Seconds = 2
      End If
    End With
  End Sub

  Public Sub Run(ByVal temperature As Integer)
    RunBase(temperature)
  End Sub

  Public Sub Run(ByVal temperature As Integer, ByVal enabled As Boolean)
    Me.Enabled = enabled
    If Not Me.Enabled Then EnabledDelayTimer.Seconds = 5
    RunBase(temperature)
  End Sub

  Public Sub RunBase(ByVal temperature As Integer)
    With controlCode
      Try
        'Check enabled - set state machine to pause if we are active and we don't have enabled set
        If (Not Enabled) AndAlso (State > EState.Restart) Then State = EState.Pause

        'Reset Steam Delay Timer if Not Heating
        If (State <> EState.Heat) Then SteamDelayTimer.Seconds = SteamDelayTime

        'Run temperature alarm checks
        RunTemperatureAlarms(temperature)

        'Loop until all state changes have completed
        Static startState As EState
        Do
          startState = State

          Select Case State
            Case EState.Off
              Timer.Seconds = 5
              Status = Translate("Idle")

            Case EState.Idle
              Status = Translate("Idle")

            Case EState.Pause
              'Set mode change timers while were in the pause state
              HeatModeChangeTimer.Seconds = .Parameters.HeatModeChangeDelay + 1
              CoolModeChangeTimer.Seconds = .Parameters.CoolModeChangeDelay + 1
              'Pause the pid to stop error sums etc.
              Pid.Pause()
              'If we get the enable signal back then restart temperature control
              If EnabledDelayTimer.Finished Then State = EState.Restart
              Status = Translate("Paused")


            Case EState.Restart
              'Go back to stored restart state and set the state timer appropriately
              Pid.Restart()
              State = RestartState
              Timer.Seconds = RestartSeconds
              Status = Translate("Restarting")


              'Heating
              '---------------------------------

            Case EState.HeatStart
              Status = Translate("Heat Start") & Timer.ToString(1)
              If Timer.Finished Then
                State = EState.HeatVent1
                Timer.Seconds = .Parameters.HeatVentTime
              End If
              SetupRestart(EState.HeatStart, 2)

            Case EState.HeatVent1
              Status = Translate("Vent Before Heat") & Timer.ToString(1)
              If Timer.Finished Then
                LoadHeatParameters()
                Pid.Reset()
                State = EState.Heat
                HeatModeChangeTimer.Seconds = .Parameters.HeatModeChangeDelay + 1
              End If
              SetupRestart(EState.HeatVent1, .Parameters.HeatVentTime)

            Case EState.Heat
              If Pid.IsMaxGradient Then
                Status = Translate("Heating") & (" ") & (Pid.PidSetpoint / 10).ToString("#0") & "F"
              Else
                Status = Translate("Heating") & (" ") & (Pid.PidSetpoint / 10).ToString("#0") & " / " & (Pid.FinalTemp / 10).ToString("#0") & "F"
              End If
              'Run PID Control
              LoadHeatParameters()
              Pid.Run(temperature)
              'Set mode change timer if the output is positive
              If (Pid.PidOutput > 0) Then HeatModeChangeTimer.Seconds = .Parameters.HeatModeChangeDelay + 1
              'Set mode change timer if we're within the HeatStepMargin of the setpoint
              If temperature < (Pid.PidSetpoint + .Parameters.HeatStepMargin) Then HeatModeChangeTimer.Seconds = .Parameters.HeatModeChangeDelay + 1
              'Check to see if we need to change mode
              If HeatModeChangeTimer.Finished Then
                If (.Parameters.HeatModeChange = 0) OrElse ((.Parameters.HeatModeChange = 1) AndAlso Pid.IsHold) Then
                  State = EState.CoolStart
                End If
              End If
              SetupRestart(EState.HeatStart, 2)

            Case EState.HeatVent2
              Status = Translate("Vent After Heat") & Timer.ToString(1)
              If Timer.Finished Then
                State = EState.Off
              End If
              SetupRestart(EState.HeatStart, 2)


              'Cooling
              '---------------------------------

            Case EState.CoolStart
              Status = Translate("Cool start") & Timer.ToString(1)
              'Set previous state variable
              If Timer.Finished Then
                State = EState.CoolVent1
                Timer.Seconds = .Parameters.CoolVentTime
              End If
              SetupRestart(EState.CoolStart, 2)

            Case EState.CoolVent1
              Status = Translate("Vent before cool") & Timer.ToString(1)
              If Timer.Finished Then
                LoadCoolParameters()
                Pid.Reset()
                State = EState.Cool
                CoolModeChangeTimer.Seconds = .Parameters.CoolModeChangeDelay + 1
              End If
              SetupRestart(EState.CoolVent1, .Parameters.CoolVentTime)

            Case EState.Cool
              If Pid.IsMaxGradient Then
                Status = Translate("Cooling") & (" ") & (Pid.PidSetpoint / 10).ToString("#0") & "F"
              Else
                Status = Translate("Cooling") & (" ") & (Pid.PidSetpoint / 10).ToString("#0") & " / " & (Pid.FinalTemp / 10).ToString("#0") & "F"
              End If
              'Run PID Control
              LoadCoolParameters()
              Pid.Run(temperature)
              'Set mode change timer if the output is negative
              If Pid.PidOutput < 0 Then CoolModeChangeTimer.Seconds = .Parameters.CoolModeChangeDelay + 1
              'Set mode change timer if we're within the Step Margin of the setpoint
              If temperature >= (Pid.PidSetpoint - .Parameters.CoolStepMargin) Then CoolModeChangeTimer.Seconds = .Parameters.CoolModeChangeDelay + 1
              'Check to see if we need to change mode
              If (CoolModeChangeTimer.Finished) Then
                If (.Parameters.CoolModeChange = 0) OrElse ((.Parameters.CoolModeChange = 1) AndAlso Pid.IsHold) Then
                  State = EState.HeatStart
                End If
              End If
              SetupRestart(EState.CoolStart, 2)

            Case EState.CoolVent2
              RestartState = EState.CoolStart
              Status = Translate("Vent after cool") & Timer.ToString(1)
              If Timer.Finished Then
                State = EState.Off
              End If
              SetupRestart(EState.CoolStart, 2)

              'Crash Cooling
              '---------------------------------

            Case EState.CrashCoolStart
              Status = Translate("Crash-Cool start") & Timer.ToString(1)
              'Set previous state variable
              If Timer.Finished Then
                pid_.Start(temperature, (.Parameters.CrashCoolTemperature - 20), 0)
                State = EState.CrashCoolVent
                Timer.Seconds = .Parameters.CoolVentTime
              End If
              SetupRestart(EState.CrashCoolStart, 2)

            Case EState.CrashCoolVent
              Status = Translate("Vent before crash-cool") & Timer.ToString(1)
              If Timer.Finished Then
                LoadCoolParameters()
                Pid.Reset()
                CoolModeChangeTimer.Seconds = .Parameters.CoolModeChangeDelay + 1
                State = EState.CrashCool
              End If
              SetupRestart(EState.CrashCoolVent, .Parameters.CoolVentTime)

            Case EState.CrashCool
              Status = Translate("Crash-Cooling")
              'Run PID Control
              LoadCoolParameters()
              Pid.Run(temperature)
              'Set mode change timer if the output is negative
              If Pid.PidOutput < 0 Then CoolModeChangeTimer.Seconds = .Parameters.CoolModeChangeDelay + 1
              'Set mode change timer if we're within the HeatStepMargin of the setpoint
              If temperature >= (Pid.PidSetpoint - .Parameters.CoolStepMargin) Then CoolModeChangeTimer.Seconds = .Parameters.CoolModeChangeDelay + 1
              'If we're at or below crash cool temperature start holding
              If temperature <= .Parameters.CrashCoolTemperature Then State = EState.CrashCoolDone
              SetupRestart(EState.CrashCool, 2)

            Case EState.CrashCoolDone
              Status = Translate("Crash-Cool Holding")
              'Run PID Control
              Pid.Run(temperature)
              'Set mode change timer if the output is negative
              If Pid.PidOutput < 0 Then CoolModeChangeTimer.Seconds = .Parameters.CoolModeChangeDelay + 1
              'Set mode change timer if we're within the HeatStepMargin of the setpoint
              If temperature >= (Pid.PidSetpoint - .Parameters.CoolStepMargin) Then CoolModeChangeTimer.Seconds = .Parameters.CoolModeChangeDelay + 1

            Case EState.CrashCoolRestart
              Status = Translate("Crash-Cool Restarting")
              'If temperature control not active then clear PID (?) and carry on
              If controlCode.HE.IsOn Then
                Start(temperature, controlCode.HE.FinalTemp, controlCode.HE.Gradient)
              ElseIf controlCode.CO.IsOn Then
                Start(temperature, controlCode.CO.FinalTemp, controlCode.CO.Gradient)
              ElseIf controlCode.TP.IsOn Then
                Start(temperature, controlCode.TP.FinalTemp, controlCode.TP.Gradient)
              Else
                pid_.Cancel()
                State = EState.Off
              End If

          End Select
        Loop Until (startState = State)

      Catch ex As Exception
        'We don't normally do this but we need to be more careful with temperature control
      End Try
    End With
  End Sub

  Public Sub Cancel()
    Pid.Cancel()
    Mode = EMode.None
    State = EState.Off
    StartTemp = 0
    FinalTemp = 0
    Gradient = 0
  End Sub

  Private Sub SetupRestart(ByVal state As EState, ByVal seconds As Integer)
    RestartState = state
    RestartSeconds = seconds
  End Sub

  Private Sub LoadHeatParameters()
    With controlCode
      'Heating parameters
      Pid.ProportionalBand = .Parameters.HeatPropBand
      Pid.Integral = .Parameters.HeatIntegral
      Pid.MaximumGradient = .Parameters.HeatMaxGradient
      Pid.BalanceTemperature = .Parameters.BalanceTemperature
      Pid.BalancePercent = .Parameters.BalancePercent
      Pid.StepMargin = .Parameters.HeatStepMargin
      If .Parameters.HeatMaxOutput > 0 Then
        Pid.MaximumOutput = .Parameters.HeatMaxOutput
      Else : Pid.MaximumOutput = 1000
      End If
      Pid.GradientRampDown = MinMax(.Parameters.GradientRampDownTime, 0, 120)
      Pid.GradientRetentionPct = MinMax(.Parameters.GradientRampDownPercent, 0, 1000)

      SteamDelayTime = .Parameters.SteamValveDelay
    End With
  End Sub

  Private Sub LoadCoolParameters()
    With controlCode
      'Cooling parameters
      Pid.ProportionalBand = .Parameters.CoolPropBand
      Pid.Integral = .Parameters.CoolIntegral
      Pid.MaximumGradient = .Parameters.CoolMaxGradient
      Pid.BalanceTemperature = .Parameters.BalanceTemperature
      Pid.BalancePercent = .Parameters.BalancePercent
      Pid.StepMargin = .Parameters.CoolStepMargin
      Pid.MaximumOutput = 1000
      Pid.GradientRampDown = .Parameters.GradientRampDownTime
      Pid.GradientRetentionPct = .Parameters.GradientRampDownPercent
    End With
  End Sub

#Region " STATE PROPERTIES "

  Private pid_ As New Pid
  Public Property Pid() As Pid
    Get
      Return pid_
    End Get
    Set(ByVal value As Pid)
      pid_ = value
    End Set
  End Property

  Private mode_ As EMode
  Public Property Mode() As EMode
    Get
      Return mode_
    End Get
    Set(ByVal value As EMode)
      mode_ = value
    End Set
  End Property

  Private heatModeChangeTimer_ As New Timer
  Public Property HeatModeChangeTimer() As Timer
    Get
      Return heatModeChangeTimer_
    End Get
    Set(ByVal value As Timer)
      heatModeChangeTimer_ = value
    End Set
  End Property

  Private heatOnDelayTimer_ As New Timer
  Public Property HeatOnDelayTimer() As Timer
    Get
      Return heatOnDelayTimer_
    End Get
    Set(ByVal value As Timer)
      heatOnDelayTimer_ = value
    End Set
  End Property

  Private coolModeChangeTimer_ As New Timer
  Public Property CoolModeChangeTimer() As Timer
    Get
      Return coolModeChangeTimer_
    End Get
    Set(ByVal value As Timer)
      coolModeChangeTimer_ = value
    End Set
  End Property

  Private coolOnDelayTimer_ As New Timer
  Public Property CoolOnDelayTimer() As Timer
    Get
      Return coolOnDelayTimer_
    End Get
    Set(ByVal value As Timer)
      coolOnDelayTimer_ = value
    End Set
  End Property

  Private steamDelayTimer_ As New Timer
  Public Property SteamDelayTimer As Timer
    Get
      Return steamDelayTimer_
    End Get
    Set(ByVal value As Timer)
      steamDelayTimer_ = value
    End Set
  End Property

  Private steamDelayTime_ As Integer
  Public Property SteamDelayTime As Integer
    Get
      Return steamDelayTime_
    End Get
    Set(ByVal value As Integer)
      steamDelayTime_ = value
    End Set
  End Property

  Public ReadOnly Property IsEnabled() As Boolean
    Get
      Return EnabledDelayTimer.Finished
    End Get
  End Property

  Public ReadOnly Property IsIdle() As Boolean
    Get
      Return (State = EState.Idle) OrElse (State = EState.Off)
    End Get
  End Property

  Public ReadOnly Property IsPaused() As Boolean
    Get
      Return (State = EState.Pause)
    End Get
  End Property

  Public ReadOnly Property IsHeating() As Boolean
    Get
      Return (State = EState.Heat)
    End Get
  End Property

  Public ReadOnly Property IsCooling() As Boolean
    Get
      Return (State = EState.Cool)
    End Get
  End Property

  Public ReadOnly Property IsCrashCooling As Boolean
    Get
      Return (State = EState.CrashCool)
    End Get
  End Property

  Public ReadOnly Property IsCrashCoolOn As Boolean
    Get
      Return (State = EState.CrashCoolStart) OrElse (State = EState.CrashCoolVent) OrElse
             (State = EState.CrashCool) OrElse (State = EState.CrashCoolDone) OrElse (State = EState.CrashCoolRestart)
    End Get
  End Property

  Public ReadOnly Property IsCrashCoolingDone As Boolean
    Get
      Return (State = EState.CrashCoolDone)
    End Get
  End Property

#End Region

#Region " Public IO Properties "

  Public ReadOnly Property IoHeatSelect() As Boolean
    Get
      If (controlCode.Parameters.HeatOnDelay > 0) AndAlso (controlCode.Parameters.HeatOnOutput > 0) Then
        Return ((State = EState.Heat) AndAlso (HeatOnDelayTimer.Finished)) OrElse
               ((State = EState.HeatVent1) AndAlso (controlCode.Parameters.HeatVentOutput > 0) AndAlso (IoOutput >= controlCode.Parameters.HeatOnOutput))
      Else
        Return (State = EState.Heat) OrElse
              ((State = EState.HeatVent1) AndAlso (controlCode.Parameters.HeatVentOutput > 0))
      End If
    End Get
  End Property

  Public ReadOnly Property IoHeatReturn() As Boolean
    Get
      Return (State = EState.Heat)
    End Get
  End Property

  Public ReadOnly Property IoCoolSelect() As Boolean
    Get
      If (controlCode.Parameters.CoolOnDelay > 0) AndAlso (controlCode.Parameters.CoolOnOutput > 0) Then
        Return ((State = EState.Cool) AndAlso (CoolOnDelayTimer.Finished)) OrElse
               ((State = EState.CoolVent1) AndAlso (controlCode.Parameters.CoolVentOutput > 0) AndAlso (IoOutput >= controlCode.Parameters.CoolOnOutput)) OrElse
                (State = EState.CrashCool)
      Else
        Return (State = EState.Cool) OrElse
              ((State = EState.CoolVent1) AndAlso (controlCode.Parameters.CoolVentOutput > 0)) OrElse
               (State = EState.CrashCool)
      End If
    End Get
  End Property

  Public ReadOnly Property IoCoolReturn() As Boolean
    Get
      Return (State = EState.Cool) OrElse (State = EState.CrashCool)
    End Get
  End Property

  Public ReadOnly Property IoHxDrain() As Boolean
    Get
      Return Not (IoHeatReturn OrElse IoCoolReturn)
    End Get
  End Property

  Public ReadOnly Property IoOutput() As Short
    Get
      'Analog Output
      IoOutput = 0

      'Heating
      If (State = EState.HeatVent1) AndAlso (controlCode.Parameters.HeatVentOutput > 0) Then IoOutput = Convert.ToInt16(controlCode.Parameters.HeatVentOutput)
      If (State = EState.Heat) Then
        'Steam Factor Used by Glen Raven to prevent Demand Spikes on Boiler 
        Dim SteamFactor As Double = (SteamDelayTime - SteamDelayTimer.Seconds) / 120

        If SteamDelayTimer.Finished Then
          IoOutput = Convert.ToInt16(Pid.PidOutput)
        Else
          IoOutput = Convert.ToInt16(Pid.PidOutput * SteamFactor)
        End If

        ' Heating Output lower than parameter for HeatOnOutput - reset timer to turn off HeatSelect (Fix for Leaking Proportional Valve)
        If (IoOutput < controlCode.Parameters.HeatOnOutput) Then HeatOnDelayTimer.Seconds = controlCode.Parameters.HeatOnDelay + 1
      Else
        ' Not Heating - reset HeatSelect On Delay Timer
        HeatOnDelayTimer.Seconds = controlCode.Parameters.HeatOnDelay + 1
      End If

      'Cooling
      If (State = EState.CoolVent1) AndAlso (controlCode.Parameters.CoolVentOutput > 0) Then IoOutput = Convert.ToInt16(controlCode.Parameters.CoolVentOutput)
      If (State = EState.Cool) Then
        IoOutput = Convert.ToInt16(-Pid.PidOutput)

        ' Cooling Output lower than parameter for CoolOnOutput - reset timer to turn off CoolSelect (Fix for Leaking Proportional Valve)
        If (IoOutput < controlCode.Parameters.CoolOnOutput) Then CoolOnDelayTimer.Seconds = controlCode.Parameters.CoolOnDelay + 1
      Else
        ' Not cooling - reset HeatSelect On Delay Timer
        CoolOnDelayTimer.Seconds = controlCode.Parameters.CoolOnDelay + 1
      End If
      If (State = EState.CrashCool) Then IoOutput = Convert.ToInt16(-Pid.PidOutput)

      'Limit output
      If IoOutput < 0 Then IoOutput = 0
      If IoOutput > 1000 Then IoOutput = 1000
    End Get
  End Property

#End Region

#Region " ALARMS "

  Private Sub RunTemperatureAlarms(ByVal temperature As Integer)
    With controlCode
      'For convenience
      Dim setpoint As Integer = Pid.PidSetpoint

      'If we do not have a setpoint or we're on a max gradient and are not holding yet then just reset timers
      If (setpoint = 0) OrElse (Pid.IsMaxGradient AndAlso (Not Pid.IsHold)) Then
        TempHighAlarmTimer.Seconds = .Parameters.TemperatureAlarmDelay + 1
        TempLowAlarmTimer.Seconds = .Parameters.TemperatureAlarmDelay + 1
        Exit Sub
      End If

      'Set temperature high alarm
      If temperature <= (setpoint + .Parameters.TemperatureAlarmBand) Then
        TempHighAlarmTimer.Seconds = .Parameters.TemperatureAlarmDelay + 1
      End If

      'Set temperature low alarm
      If temperature >= (setpoint - .Parameters.TemperatureAlarmBand) Then
        TempLowAlarmTimer.Seconds = .Parameters.TemperatureAlarmDelay + 1
      End If

      ' Check parameters are set for alarms to be enabled
      If .Parameters.TemperatureAlarmBand = 0 Then TempHighAlarmTimer.Seconds = 10
      If .Parameters.TemperatureAlarmDelay = 0 Then TempHighAlarmTimer.Seconds = 10

      'If we're holding then don't check for pid pause / restart these are only used on a ramp
      If Pid.IsHold Then
        'Just in case the pid was paused earlier
        If Pid.IsPaused Then Pid.Restart()
      Else
        RunPidAlarms(temperature)
      End If

      ' Clear Alarms with Advance Pushbutton
      If Not .AdvancePb Then advanceTimer_.Seconds = 2
      If .AdvancePb AndAlso advanceTimer_.Finished Then
        TempHighAlarmTimer.Seconds = .Parameters.TemperatureAlarmDelay + 1
        TempLowAlarmTimer.Seconds = .Parameters.TemperatureAlarmDelay + 1
      End If

    End With
  End Sub

  Private Sub RunPidAlarms(ByVal temperature As Integer)
    With controlCode
      'For convenience
      Dim setpoint As Integer = Pid.PidSetpoint

      'Calculate absolute temperature error
      Dim deltaAbs As Integer = Math.Abs(setpoint - temperature)

      'Pause the pid if the error is too great
      If deltaAbs > .Parameters.TemperaturePidPause Then
        Pid.Pause()
      End If

      'Restart the pid if the error is back within parameter values
      If Pid.IsPaused Then
        If (deltaAbs < .Parameters.TemperaturePidPause) OrElse (deltaAbs < .Parameters.TemperaturePidRestart) Then
          Pid.Restart()
        End If
      End If

      'Reset the pid if the error is too great
      If Pid.IsPaused Then
        If deltaAbs > .Parameters.TemperaturePidReset Then
          Pid.Restart()
        End If
      End If
    End With
  End Sub

  Private advanceTimer_ As New Timer

  ' TODO from vb6
  Public IgnoreErrors As Boolean
  Public TempLow As Boolean
  Public TempHigh As Boolean

  Private tempHighAlarmTimer_ As New Timer
  Public Property TempHighAlarmTimer() As Timer
    Get
      Return tempHighAlarmTimer_
    End Get
    Set(ByVal value As Timer)
      tempHighAlarmTimer_ = value
    End Set
  End Property

  Private tempLowAlarmTimer_ As New Timer
  Public Property TempLowAlarmTimer() As Timer
    Get
      Return tempLowAlarmTimer_
    End Get
    Set(ByVal value As Timer)
      tempLowAlarmTimer_ = value
    End Set
  End Property

#End Region

End Class
