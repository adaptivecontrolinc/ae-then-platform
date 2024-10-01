' Temperature control for machines with a heat exchanger drain
Public Class TemperatureControl : Inherits MarshalByRefObject

  Public Enum EState
    Off
    Heat
    Cool
    CrashCool
  End Enum
  Public State As EState
  Public Enable As Boolean

  Public Enum EModeChange
    Always = 0       ' Can change at any time ramp or hold
    HoldOnly = 1     ' Only change on a hold
    Never = 2        ' Never change state
    HeatOnly = 3     ' Heat only - no mode change
    CoolOnly = 4     ' Cool only - no mode change
  End Enum

  Public Heat As New TemperatureControlLoop With {.HeatLoop = True}
  Public Cool As New TemperatureControlLoop With {.CoolLoop = True}
  Public CrashCool As New TemperatureControlLoop With {.CoolLoop = True}

  Public StartTemp As Integer
  Public FinalTemp As Integer
  Public Gradient As Integer

  Sub Start(startTemp As Integer, finalTemp As Integer, gradient As Integer, stateChange As EModeChange, stateChangeDelay As Integer)
    Me.StartTemp = startTemp
    Me.FinalTemp = finalTemp
    Me.Gradient = gradient

    If finalTemp < startTemp - 50 Then
      StartCool(startTemp, finalTemp, gradient, stateChange, stateChangeDelay)
    Else
      StartHeat(startTemp, finalTemp, gradient, stateChange, stateChangeDelay)
    End If
  End Sub

  Sub StartHeat(startTemp As Integer, finalTemp As Integer, gradient As Integer, stateChange As EModeChange, stateChangeDelay As Integer)
    Cool.Cancel()
    CrashCool.Cancel()

    Me.StartTemp = startTemp
    Me.FinalTemp = finalTemp
    Me.Gradient = gradient

    Heat.Start(startTemp, finalTemp, gradient, stateChange, stateChangeDelay)
    State = EState.Heat
  End Sub

  Sub StartCool(startTemp As Integer, finalTemp As Integer, gradient As Integer, stateChange As EModeChange, stateChangeDelay As Integer)
    Heat.Cancel()
    CrashCool.Cancel()

    Me.StartTemp = startTemp
    Me.FinalTemp = finalTemp
    Me.Gradient = gradient

    Cool.Start(startTemp, finalTemp, gradient, stateChange, stateChangeDelay)
    State = EState.Cool
  End Sub

  Sub StartCrashCool(startTemp As Integer, finalTemp As Integer, gradient As Integer)
    If Heat.IsOn Then Heat.Pause()
    If Cool.IsOn Then Cool.Pause()

    Me.StartTemp = startTemp
    Me.FinalTemp = finalTemp
    Me.Gradient = gradient

    CrashCool.Start(startTemp, finalTemp, gradient, EModeChange.CoolOnly, 0)
    State = EState.CrashCool
  End Sub

  Sub ResetCrashCool(currentTemp As Integer)
    Heat.Cancel()
    Cool.Cancel()

    State = EState.CrashCool
  End Sub

  Sub Run(currentTemp As Integer)
    ' Run each temperature control loop
    Heat.Run(currentTemp)
    Cool.Run(currentTemp)
    CrashCool.Run(currentTemp)

    ' Heat cool mode change
    Select Case State
      Case EState.Heat
        If Heat.IsModeChange Then
          Heat.ModeChangePause() : Cool.Start(currentTemp, Heat.TargetTemp, Heat.Gradient, EModeChange.Always, 120)
        ElseIf Heat.IsModeChangePause AndAlso (currentTemp < (Heat.TargetTemp + Heat.HoldMargin)) Then
          Cool.Cancel() : Heat.ModeChangeRestart()
        End If

      Case EState.Cool
        If Cool.IsModeChange Then
          Cool.ModeChangePause() : Heat.Start(currentTemp, Cool.TargetTemp, Cool.Gradient, EModeChange.Always, 120)
        ElseIf Cool.IsModeChangePause AndAlso (currentTemp > (Cool.TargetTemp - Cool.HoldMargin)) Then
          Heat.Cancel() : Cool.ModeChangeRestart()
        End If

    End Select
  End Sub

  Sub Cancel()
    Heat.Cancel()
    Cool.Cancel()
    CrashCool.Cancel()
    State = EState.Off
  End Sub

  Sub Pause()
    Heat.Pause()
    Cool.Pause()
    CrashCool.Pause()
  End Sub

  Sub Restart()
    Heat.Restart()
    Cool.Restart()
    CrashCool.Restart()
  End Sub

  Sub Reset()
    'TODO ??
  End Sub

  ReadOnly Property IsOn As Boolean
    Get
      Return State <> EState.Off
    End Get
  End Property

  ReadOnly Property IsOff As Boolean
    Get
      Return State = EState.Off
    End Get
  End Property

  ReadOnly Property IsHeat As Boolean
    Get
      Return State = EState.Heat
    End Get
  End Property

  ReadOnly Property IsCool As Boolean
    Get
      Return State = EState.Cool
    End Get
  End Property

  ReadOnly Property IsCrashCool As Boolean
    Get
      Return State = EState.CrashCool
    End Get
  End Property

  ReadOnly Property IsGradientMax As Boolean
    Get
      Return Gradient = 0 OrElse Gradient = 99
    End Get
  End Property

  ReadOnly Property IsRamp() As Boolean
    Get
      Return (IsHeat AndAlso Heat.IsRamp) OrElse (IsCool AndAlso Cool.IsRamp)
    End Get
  End Property

  ReadOnly Property IsHold() As Boolean
    Get
      Return (IsHeat AndAlso Heat.IsHold) OrElse (IsCool AndAlso Cool.IsHold)
    End Get
  End Property


  ReadOnly Property IsPidPaused() As Boolean
    Get
      Return (IsHeat AndAlso Heat.Pid.IsPaused) OrElse (IsCool AndAlso Cool.Pid.IsPaused)
    End Get
  End Property

  ReadOnly Property TargetTemp As Integer
    Get
      If CrashCool.IsOn Then Return CrashCool.TargetTemp
      If Heat.IsOn Then Return Heat.TargetTemp
      If Cool.IsOn Then Return Cool.TargetTemp
    End Get
  End Property

  ReadOnly Property Status As String
    Get
      If CrashCool.IsOn Then Return "CrashCool: " & CrashCool.Status
      If Heat.IsOn Then Return "Heat: " & Heat.Status
      If Cool.IsOn Then Return "Cool: " & Cool.Status
      Return "Off"
    End Get
  End Property

  ReadOnly Property ActiveLoop As TemperatureControlLoop
    Get
      If CrashCool.IsOn Then Return CrashCool
      If Heat.IsOn Then Return Heat
      If Cool.IsOn Then Return Cool
      Return Nothing
    End Get
  End Property

  ReadOnly Property IOOutput As Short
    Get
      If CrashCool.IsOn Then Return -CrashCool.IOOutput
      If Heat.IsOn Then Return Heat.IOOutput
      If Cool.IsOn Then Return -Cool.IOOutput
      Return 0
    End Get
  End Property

  ReadOnly Property IsPaused As Boolean
    Get
      If ActiveLoop IsNot Nothing Then Return ActiveLoop.IsPaused
    End Get
  End Property



  Public Class TemperatureControlLoop

    Public Enum EState
      Off
      Interlock

      Start
      Vent
      Ramp
      Hold

      ModeChangePause

      Pause
      Restart
    End Enum
    Public State As EState
    Public RestartState As EState
    Public Status As String
    Public Timer As New Timer

    Public Pid As New PidControlML
    Public HeatLoop As Boolean
    Public CoolLoop As Boolean

    Public ModeChange As EModeChange
    Public ModeChangeDelay As Integer
    Public ModeChangeDelayTimer As New Timer

    Public StartTemp As Integer
    Public FinalTemp As Integer
    Public Gradient As Integer
    Public TargetTemp As Integer

    Public HoldMargin As Integer
    Public HoldMarginDelay As Integer = 4
    Public HoldMarginDelayTimer As New Timer

    Public Enable As Boolean
    Public EnableDelay As Integer
    Public EnableDelayTimer As New Timer

    Public VentTime As Integer


    Sub Parameters(enableDelay As Integer, ventTime As Integer, holdMargin As Integer, holdMarginDelay As Integer)
      Me.EnableDelay = enableDelay
      Me.VentTime = ventTime
      Me.HoldMargin = holdMargin
      Me.HoldMarginDelay = holdMarginDelay
    End Sub

    Sub Parameters(enableDelay As Integer, ventTime As Integer, holdMargin As Integer, holdMarginDelay As Integer, modechange As Integer, modeChangeDelay As Integer)
      Me.EnableDelay = enableDelay
      Me.VentTime = ventTime
      Me.HoldMargin = holdMargin
      Me.HoldMarginDelay = holdMarginDelay

      Dim newModeChange = CType(modechange, EModeChange)
      Me.ModeChange = newModeChange
      Me.ModeChangeDelay = modeChangeDelay

      ' Reset mode change delay time if it changes
      If modechange <> newModeChange Then ModeChangeDelayTimer.Seconds = modeChangeDelay
    End Sub

    Sub Start(startTemp As Integer, finalTemp As Integer, gradient As Integer, modeChange As EModeChange, modeChangeDelayTime As Integer)
      Me.StartTemp = startTemp
      Me.FinalTemp = finalTemp
      Me.Gradient = gradient
      Me.ModeChange = modeChange
      Me.ModeChangeDelay = modeChangeDelayTime

      Select Case State
        Case EState.Ramp, EState.Hold         ' If we're already heating then just set new pid parameters
          State = EState.Start
        Case Else
          Pid.Cancel()
          State = EState.Interlock
      End Select
    End Sub

    Sub Run(ByVal currentTemp As Integer)
      ' Run the pid
      Pid.HoldMargin = HoldMargin
      Pid.Enabled = (State = EState.Start OrElse State = EState.Ramp OrElse State = EState.Hold)

      TargetTemp = Pid.TargetTemp

      ' Set enable delay timer
      If Not Enable Then EnableDelayTimer.Seconds = MinMax(EnableDelay, 1, 60)

      ' Set hold delay timer (must be in hold margin for delay seconds before switching from ramp to hold)
      If HeatLoop AndAlso currentTemp < (FinalTemp - HoldMargin) Then HoldMarginDelayTimer.Seconds = MinMax(HoldMarginDelay, 2, 60)
      If CoolLoop AndAlso currentTemp > (FinalTemp + HoldMargin) Then HoldMarginDelayTimer.Seconds = MinMax(HoldMarginDelay, 2, 60)

      ' Set mode change timer
      If Not (State = EState.Ramp OrElse State = EState.Hold) Then ModeChangeDelayTimer.Seconds = 120
      If HeatLoop AndAlso currentTemp < TargetTemp + 30 Then ModeChangeDelayTimer.Seconds = ModeChangeDelay
      If CoolLoop AndAlso currentTemp > TargetTemp - 30 Then ModeChangeDelayTimer.Seconds = ModeChangeDelay

      Select Case State

        Case EState.Off
          RestartState = EState.Off
          Status = "Off"

        Case EState.Interlock
          If EnableDelayTimer.Finished Then
            State = EState.Vent
            Timer.Seconds = VentTime
          End If
          Status = "Interlock " & EnableDelayTimer.ToString

        Case EState.Vent
          ' Vent heat exchanger before we start control
          If Timer.Finished Then
            State = EState.Start
          End If
          Status = "Vent " & Timer.ToString

        Case EState.Start
          Pid.Start(currentTemp, FinalTemp, Gradient)
          State = EState.Ramp
          Status = "Start"

        Case EState.Ramp
          Pid.Run(currentTemp)
          If HoldMarginDelayTimer.Finished Then State = EState.Hold
          Status = "Ramp"

        Case EState.Hold
          Pid.Run(currentTemp)
          Status = "Hold"


        Case EState.ModeChangePause
          Status = "Mode Change"

        Case EState.Pause
          Status = "Pause"

      End Select

      ' Force to interlock state if we lose enable 
      If (Not Enable) AndAlso (State = EState.Ramp OrElse State = EState.Hold) Then
        Pid.Pause()
        State = EState.Interlock
      End If

    End Sub

    Sub Cancel()
      State = EState.Off
      RestartState = EState.Off
      TargetTemp = 0
      Pid.Cancel()
    End Sub

    Sub Pause()
      If IsOn AndAlso State <> EState.Pause AndAlso State <> EState.Restart Then
        RestartState = State
        State = EState.Pause
        Pid.Pause()
      End If
    End Sub

    Sub Restart()
      If State = EState.Pause Then
        State = RestartState
        RestartState = EState.Off
        If State = EState.Vent Then Timer.Seconds = VentTime
        Pid.Restart()
      End If
    End Sub

    Sub ModeChangePause()
      If State <> EState.ModeChangePause Then
        RestartState = State
        State = EState.ModeChangePause
      End If
    End Sub

    Sub ModeChangeRestart()
      If State = EState.ModeChangePause Then
        State = RestartState
        RestartState = EState.Off

        Timer.Seconds = VentTime
        State = EState.Vent
      End If
    End Sub

    ReadOnly Property IsOn As Boolean
      Get
        Return State <> EState.Off
      End Get
    End Property

    ReadOnly Property IsPaused As Boolean
      Get
        Return State = EState.Pause
      End Get
    End Property

    ReadOnly Property IsRamp() As Boolean
      Get
        Return State = EState.Ramp
      End Get
    End Property

    ReadOnly Property IsHold() As Boolean
      Get
        Return State = EState.Hold AndAlso Pid.Ramp.IsDone
      End Get
    End Property

    ReadOnly Property IsModeChange As Boolean
      Get
        If State = EState.Hold AndAlso ModeChange = EModeChange.HoldOnly Then Return ModeChangeDelayTimer.Finished
        If (State = EState.Ramp OrElse State = EState.Hold) AndAlso ModeChange = EModeChange.Always Then Return ModeChangeDelayTimer.Finished
        Return False
      End Get
    End Property

    ReadOnly Property IsModeChangePause As Boolean
      Get
        Return State = EState.ModeChangePause
      End Get
    End Property

    ReadOnly Property IOVent As Boolean
      Get
        Return State = EState.Vent
      End Get
    End Property

    ReadOnly Property IOIsolate As Boolean
      Get
        Return State = EState.Start OrElse State = EState.Ramp OrElse State = EState.Hold
      End Get
    End Property

    ReadOnly Property IOOutput As Short
      Get
        If IOIsolate Then Return CShort(Pid.Output)
        Return 0
      End Get
    End Property

    ReadOnly Property IOHeatOutput As Short
      Get
        If IOIsolate Then Return CShort(MinMax(Pid.Output, 0, 1000))
        Return 0
      End Get
    End Property

    ReadOnly Property IOCoolOutput As Short
      Get
        If IOIsolate Then Return CShort(MinMax(-Pid.Output, 0, 1000))
        Return 0
      End Get
    End Property

  End Class

End Class