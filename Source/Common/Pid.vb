Public Class Pid : Inherits MarshalByRefObject
  Public Enum EState
    Off
    RampUp
    RampDown
    Hold
    Pause
  End Enum

  Private oneSecondTimer_ As New Timer

  Private gradientRampTimer_ As New Timer

  Public Sub Start(ByVal startTemp As Integer, ByVal finalTemp As Integer, ByVal gradient As Integer)
    'Set temperatures and gradients
    Me.StartTemp = startTemp
    Me.FinalTemp = finalTemp
    Me.Gradient = gradient

    'Limit the final temp 
    If Me.FinalTemp > MaximumTemperature Then Me.FinalTemp = MaximumTemperature
    If Me.FinalTemp < MinimumTemperature Then Me.FinalTemp = MinimumTemperature

    'Setup ramp
    Me.RampTime = 0
    If Not IsMaxGradient Then Me.RampTime = Math.Abs(Convert.ToInt32(((finalTemp - startTemp) / gradient) * 60))
    Me.RampTimer.Seconds = RampTime

    'Reset control
    Reset()

    'Start the pid
    State = EState.Hold
    If startTemp < finalTemp Then State = EState.RampUp
    If startTemp > finalTemp Then State = EState.RampDown

    gradientRampTimer_.Seconds = GradientRetentionPct

  End Sub

  Public Sub Run(ByVal temperature As Integer)
    'Run the ramp generator
    RunRampGenerator()

    'Calculate pidError
    PidError = PidSetpoint - temperature

    'Only run the calculation if we're ramping or holding
    If State = EState.Off Then
      PidSetpoint = 0
      PidOutput = 0
      Exit Sub
    End If

    'Check to see if we need to switch from ramp to hold
    If IsMaxGradient Then
      If State = EState.RampUp AndAlso temperature >= (FinalTemp - StepMargin) Then
        State = EState.Hold
      End If
      If State = EState.RampDown AndAlso temperature <= (FinalTemp + StepMargin) Then
        State = EState.Hold
      End If
    Else
      If PidSetpoint = FinalTemp Then State = EState.Hold
      If (temperature >= (FinalTemp - StepMargin)) AndAlso (temperature <= (FinalTemp + StepMargin)) Then
        State = EState.Hold
      End If
    End If

    'Calculate error sum 
    'If PID output is maxed out stop the error sum
    '  this should prevent the error sum getting huge and taking forever to wind down (integeral saturation)
    Dim stopPidErrorSum As Boolean
    If (PidOutput >= MaximumOutput) AndAlso (PidError > 0) Then stopPidErrorSum = True
    If (PidOutput <= -MaximumOutput) AndAlso (PidError < 0) Then stopPidErrorSum = True

    'Sum the error every second if allowed
    If oneSecondTimer_.Finished Then
      If Not stopPidErrorSum Then PidErrorSum += PidError
      oneSecondTimer_.Seconds = 1
    End If

    'Calculate proportional term
    TermProportional = (PidError * 1000) \ ProportionalBand
    If TermProportional > 1000 Then TermProportional = 1000
    If TermProportional < -1000 Then TermProportional = -1000

    'Calculate integeral term
    TermIntegral = (PidErrorSum * Integral) \ 6000
    If TermIntegral > 1000 Then TermIntegral = 1000
    If TermIntegral < -1000 Then TermIntegral = -1000

    'Calculate balance term
    TermBalance = ((temperature - BalanceTemperature) * BalancePercent) \ 1000
    If TermBalance > 1000 Then TermBalance = 1000
    If TermBalance < -1000 Then TermBalance = -1000


    'Calculate gradient term
    TermGradient = 0
    If MaximumGradient > 0 Then
      If IsRampUp Then TermGradient = (1000 * Gradient) \ MaximumGradient
      If IsRampDown Then TermGradient = -(1000 * Gradient) \ MaximumGradient
      ' Reset Gradient Ramp Timer
      If IsRampUp OrElse IsRampDown Then gradientRampTimer_.Seconds = GradientRampDown
    Else
      TermGradient = 0
      gradientRampTimer_.Seconds = GradientRampDown
    End If
    If TermGradient > 1000 Then TermGradient = 1000
    If TermGradient < -1000 Then TermGradient = -1000

    ' Use Gradient timer to prevent over/under shooting of temperature when first stop ramping
    If IsHold Then
      If GradientRampDown > 0 Then
        TermGradient = CInt((TermGradient / GradientRampDown) * gradientRampTimer_.Seconds * (GradientRetentionPercent / 1000))
      Else : TermGradient = 0
      End If
    End If

    ' Sum all values
    PidOutput = TermProportional + TermIntegral + TermBalance + TermGradient

    'Limit based on maximum Parameter
    If PidOutput > MaximumOutput Then PidOutput = MaximumOutput
    If PidOutput < -MaximumOutput Then PidOutput = -MaximumOutput

    'Limit based on 100%
    If PidOutput > 1000 Then PidOutput = 1000
    If PidOutput < -1000 Then PidOutput = -1000
  End Sub

  Private Sub RunRampGenerator()
    'If this is a maximum gradient setpoint = final temp
    If IsMaxGradient Then
      PidSetpoint = FinalTemp
      Exit Sub
    End If

    'If the ramp timer has finished setpoint = final temperature
    If RampTimer.Finished Then
      PidSetpoint = FinalTemp
      Exit Sub
    End If

    'Calculate our postion on the ramp
    Dim rampPosition As Double = (FinalTemp - StartTemp) * ((RampTime - RampTimer.Seconds) / RampTime)

    'Calculate the setpoint temperature and make sure it does not exceed temperature limits
    Dim temperature As Integer = (StartTemp + Convert.ToInt32(rampPosition))
    If temperature > MaximumTemperature Then temperature = MaximumTemperature
    If temperature < MinimumTemperature Then temperature = MinimumTemperature

    'Set the value
    PidSetpoint = temperature
  End Sub

  Public Sub Cancel()
    Reset()
    PidSetpoint = 0
    StartTemp = 0
    finalTemp_ = 0
    termGradient_ = 0
    termBalance_ = 0
    termProportional_ = 0
    termIntegral_ = 0

    State = EState.Off
  End Sub

  Public Sub Pause()
    If State <> EState.Pause Then
      State = EState.Pause
      RampTimer.Pause()
    End If
  End Sub

  Public Sub Restart()
    If State = EState.Pause Then
      RampTimer.Restart()
      State = EState.Hold
      If StartTemp < FinalTemp Then State = EState.RampUp
      If StartTemp > FinalTemp Then State = EState.RampDown
    End If
  End Sub

  Public Sub Reset(ByVal temperature As Integer)
    Me.StartTemp = temperature
    Reset()
  End Sub

  Public Sub Reset()
    PidError = 0
    PidErrorSum = 0
    PidOutput = 0
  End Sub

#Region " PARAMETERS "


  Private integral_ As Integer
  Public Property Integral() As Integer
    Get
      Return integral_
    End Get
    Set(ByVal value As Integer)
      integral_ = value
      If integral_ < 0 Then integral_ = 0
      If integral_ > 1000 Then integral_ = 1000
    End Set
  End Property

  Private proportionalBand_ As Integer
  Public Property ProportionalBand() As Integer
    Get
      'Prevent divide by 0
      If proportionalBand_ = 0 Then proportionalBand_ = 1
      Return proportionalBand_
    End Get
    Set(ByVal value As Integer)
      proportionalBand_ = value
      If proportionalBand_ = 0 Then proportionalBand_ = 1
      If proportionalBand_ > 1000 Then proportionalBand_ = 1000
    End Set
  End Property

  Private balanceTemperature_ As Integer
  Public Property BalanceTemperature() As Integer
    Get
      Return balanceTemperature_
    End Get
    Set(ByVal value As Integer)
      balanceTemperature_ = value
    End Set
  End Property

  Private balancePercent_ As Integer
  Public Property BalancePercent() As Integer
    Get
      Return balancePercent_
    End Get
    Set(ByVal value As Integer)
      balancePercent_ = MinMax(value, 0, 1000)
    End Set
  End Property

  Private maximumGradient_ As Integer
  Public Property MaximumGradient() As Integer
    Get
      Return maximumGradient_
    End Get
    Set(ByVal value As Integer)
      maximumGradient_ = MinMax(value, 0, 1000)
    End Set
  End Property

  Private gradientRampDown_ As Integer
  Public Property GradientRampDown As Integer
    Get
      Return gradientRampDown_
    End Get
    Set(value As Integer)
      gradientRampDown_ = value
    End Set
  End Property

  Private gradientRetentionPct_ As Integer
  Public Property GradientRetentionPercent As Integer
    Get
      Return gradientRetentionPct_
    End Get
    Set(value As Integer)
      gradientRetentionPct_ = value
    End Set
  End Property

  Public ReadOnly Property GradientRampTime As Integer
    Get
      Return gradientRampTimer_.Seconds
    End Get
  End Property

  Private stepMargin_ As Integer
  Public Property StepMargin() As Integer
    Get
      Return stepMargin_
    End Get
    Set(ByVal value As Integer)
      stepMargin_ = value
      If stepMargin_ < 10 Then stepMargin_ = 10
      If stepMargin_ > 100 Then stepMargin_ = 100
    End Set
  End Property

  Private minimumTemperature_ As Integer = 400
  Public Property MinimumTemperature() As Integer
    Get
      If minimumTemperature_ < 400 Then Return 400
      Return minimumTemperature_
    End Get
    Set(ByVal value As Integer)
      minimumTemperature_ = value
    End Set
  End Property

  Private maximumTemperature_ As Integer = 2800 '1400
  Public Property MaximumTemperature() As Integer
    Get
      If maximumTemperature_ > 2800 Then Return 2800
      Return maximumTemperature_
    End Get
    Set(ByVal value As Integer)
      maximumTemperature_ = value
    End Set
  End Property

  Private maximumOutput_ As Integer = 1000
  Public Property MaximumOutput() As Integer
    Get
      If maximumOutput_ > 1000 Then Return 1000
      Return maximumOutput_
    End Get
    Set(ByVal value As Integer)
      maximumOutput_ = value
      If maximumOutput_ < 400 Then maximumOutput_ = 400
      If maximumOutput_ > 1000 Then maximumOutput_ = 1000
    End Set
  End Property

#End Region

#Region " PROPERTIES "

  Private state_ As EState
  Public Property State() As EState
    Get
      Return state_
    End Get
    Private Set(ByVal value As EState)
      state_ = value
    End Set
  End Property

  Private timer_ As New Timer
  Public Property Timer() As Timer
    Get
      Return timer_
    End Get
    Private Set(ByVal value As Timer)
      timer_ = value
    End Set
  End Property

  Private startTemp_ As Integer
  Public Property StartTemp() As Integer
    Get
      Return startTemp_
    End Get
    Set(ByVal value As Integer)
      startTemp_ = value
    End Set
  End Property

  Private finalTemp_ As Integer
  Public Property FinalTemp() As Integer
    Get
      Return finalTemp_
    End Get
    Set(ByVal value As Integer)
      finalTemp_ = value
    End Set
  End Property

  Private gradient_ As Integer
  Public Property Gradient() As Integer
    Get
      Return gradient_
    End Get
    Set(ByVal value As Integer)
      gradient_ = value
    End Set
  End Property

  Private rampTime_ As Integer
  Public Property RampTime() As Integer
    Get
      Return rampTime_
    End Get
    Set(ByVal value As Integer)
      rampTime_ = value
    End Set
  End Property

  Private rampTimer_ As New Timer
  Public Property RampTimer() As Timer
    Get
      Return rampTimer_
    End Get
    Set(ByVal value As Timer)
      rampTimer_ = value
    End Set
  End Property

  Public ReadOnly Property IsMaxGradient() As Boolean
    Get
      Return ((Gradient = 0) OrElse (Gradient = 99))
    End Get
  End Property

  Public ReadOnly Property IsRampUp() As Boolean
    Get
      Return (State = EState.RampUp)
    End Get
  End Property

  Public ReadOnly Property IsRampDown() As Boolean
    Get
      Return (State = EState.RampDown)
    End Get
  End Property

  Public ReadOnly Property IsHold() As Boolean
    Get
      Return (State = EState.Hold)
    End Get
  End Property

  Public ReadOnly Property IsPaused() As Boolean
    Get
      Return (State = EState.Pause)
    End Get
  End Property

#End Region

#Region " PID TERMS "

  Private pidSetpoint_ As Integer
  Public Property PidSetpoint() As Integer
    Get
      Return pidSetpoint_
    End Get
    Private Set(ByVal value As Integer)
      pidSetpoint_ = value
    End Set
  End Property

  Private pidError_ As Integer
  Public Property PidError() As Integer
    Get
      Return pidError_
    End Get
    Private Set(ByVal value As Integer)
      pidError_ = value
    End Set
  End Property

  Private pidErrorSum_ As Integer
  Public Property PidErrorSum() As Integer
    Get
      Return pidErrorSum_
    End Get
    Private Set(ByVal value As Integer)
      pidErrorSum_ = value
    End Set
  End Property

  Private pidOutput_ As Integer
  Public Property PidOutput() As Integer
    Get
      Return pidOutput_
    End Get
    Private Set(ByVal value As Integer)
      pidOutput_ = value
    End Set
  End Property

  Private termProportional_ As Integer
  Public Property TermProportional() As Integer
    Get
      Return termProportional_
    End Get
    Private Set(ByVal value As Integer)
      termProportional_ = value
    End Set
  End Property

  Private termIntegral_ As Integer
  Public Property TermIntegral() As Integer
    Get
      Return termIntegral_
    End Get
    Private Set(ByVal value As Integer)
      termIntegral_ = value
    End Set
  End Property

  Private termGradient_ As Integer
  Public Property TermGradient() As Integer
    Get
      Return termGradient_
    End Get
    Private Set(ByVal value As Integer)
      termGradient_ = value
    End Set
  End Property

  Public Property GradientRetentionPct As Integer

  Private termBalance_ As Integer
  Public Property TermBalance() As Integer
    Get
      Return termBalance_
    End Get
    Private Set(ByVal value As Integer)
      termBalance_ = value
    End Set
  End Property

#End Region

End Class