Public Class PidControlML
  Inherits MarshalByRefObject

  Public Enum EPidState ' Are we heating, cooling or holding temp
    Off
    Interlock
    Ramp
    RampToHold
    Hold
    RampPause
    HoldPause
  End Enum
  Public State As EPidState

  Public Enabled As Boolean

  Public Ramp As New RampGenerator
  Public HoldDelayTimer As New Timer

  ' Current values
  Public CurrentTemp As Integer
  Public TargetTemp As Integer
  Public Delta As Integer
  'Public Output As Integer

  ' Starting values
  Public StartTemp As Integer
  Public FinalTemp As Integer
  Public Gradient As Integer
  Public HoldMargin As Integer

  ' Pid control parameters
  Public PropBand As Integer
  Public Integral As Integer
  Public Balance As Integer = 32
  Public GradientMax As Integer
  Public GradientTaperMargin As Integer = 0
  Public GradientTaperTimer As New Timer

  Public GradientHoldTime As Integer
  Public IntegralBump As Integer

  ' Pid terms - these are summed to calculate the pid output
  Public PropTerm As Integer
  Public IntegralTerm As Integer
  Public BalanceTerm As Integer
  Public GradientTerm As Integer

  ' TODO Make these private once we're debugged thoroughly

  ' Integral array - limit number of values we store for integral action 
  Public IntegralTimer As New Timer
  Public IntegralSum As Integer              ' Sum of integral values
  Public IntegralCount As Integer

  Public Sub Parameters(propBand As Integer, integral As Integer, gradientMax As Integer)
    Me.PropBand = MinMax(propBand, 10, 1000)
    Me.Integral = MinMax(integral, 10, 1000)
    Me.GradientMax = MinMax(gradientMax, 10, 1000)
    Me.HoldMargin = HoldMargin
  End Sub

  Public Sub Start(ByVal startTemp As Integer, ByVal finalTemp As Integer, ByVal gradient As Integer)
    ' Reset integral if this is a new heating / cooling cycle
    If IsHeat AndAlso finalTemp < Me.FinalTemp Then ResetIntegral()
    If IsCool AndAlso finalTemp > Me.FinalTemp Then ResetIntegral()

    'Set parameters
    Me.StartTemp = startTemp
    Me.FinalTemp = finalTemp
    Me.Gradient = gradient

    ' Start ramp generator
    If Math.Abs(finalTemp - startTemp) <= (HoldMargin * 2) Then
      Ramp.Start(finalTemp, finalTemp, gradient)
    Else
      Ramp.Start(startTemp, finalTemp, gradient)
    End If

    ' Set gradient negative if we are cooling
    If finalTemp < startTemp Then Me.Gradient = -gradient

    IntegralTimer.Seconds = 45  ' delay start of integral action
    State = EPidState.Ramp
  End Sub

  Public Sub Run(ByVal temperature As Integer)
    ' Get setpoint from ramp generator
    CurrentTemp = temperature
    TargetTemp = Ramp.Setpoint

    ' Calculate temperature difference 
    Delta = TargetTemp - temperature

    ' Switch to hold if ramp is finished and temperature is above setpoint
    If State = EPidState.Ramp AndAlso Ramp.IsDone Then
      If IsHeat AndAlso temperature >= (FinalTemp - 2) Then State = EPidState.RampToHold
      If IsCool AndAlso temperature <= (FinalTemp + 2) Then State = EPidState.RampToHold
    End If

    ' Run simple state machine
    Select Case State
      Case EPidState.Off
        Output = 0
      Case EPidState.Interlock
        Output = 0
        If Enabled Then State = EPidState.Ramp
      Case EPidState.Ramp
        PropTerm = MinMax(GetPropTerm(Delta), -1200, 1200)                  ' let this be bigger than 1000 so we still get to 1000 during cooling with the balance term
        IntegralTerm = MinMax(GetIntegralTerm(Delta), -500, 500)            ' limit to 50%
        BalanceTerm = MinMax(GetBalanceTerm(temperature), -250, 250)        ' limit to 25%
        GradientTerm = MinMax(GetGradientTerm(temperature), -500, 500)      ' limit to 50%  'TODO maybe limit to 75% ??
        Output = MinMax(PropTerm + IntegralTerm + GradientTerm + BalanceTerm, -1000, 1000)
        GradientTaperTimer.Seconds = GradientHoldTime
        If Not Enabled Then Pause()
      Case EPidState.RampToHold
        IntegralGradientBump()
        State = EPidState.Hold
      Case EPidState.Hold
        PropTerm = MinMax(GetPropTerm(Delta), -1200, 1200)                  ' let this be bigger than 1000 so we still get to 1000 during cooling with the balance term
        IntegralTerm = MinMax(GetIntegralTerm(Delta), -500, 500)            ' limit to 50%
        BalanceTerm = MinMax(GetBalanceTerm(temperature), -250, 250)        ' limit to 25%
        GradientTerm = GetGradientTermHold(GradientTaperTimer.TimeRemaining, GradientHoldTime)
        Output = MinMax(PropTerm + IntegralTerm + GradientTerm + BalanceTerm, -1000, 1000)
        If Not Enabled Then Pause()
      Case EPidState.RampPause
        Output = 0
        If Enabled Then Restart()
      Case EPidState.HoldPause
        Output = 0
        If Enabled Then Restart()
      Case Else
        Output = 0
    End Select
  End Sub

  Private Function GetPropTerm(delta As Integer) As Integer
    If PropBand > 0 Then Return CInt((delta / PropBand) * 1000)
    Return 0
  End Function

  Private Function GetIntegralTerm(delta As Integer) As Integer
    ' Keep integral at 0 until we get close to setpoint if this is a max ramp rate
    If State = EPidState.Ramp AndAlso Ramp.RampMax Then
      If Not IntegralStartRampMax(delta) Then
        ResetIntegral()
        Return 0
      End If
    End If

    If IntegralTimer.Finished Then
      IntegralTimer.Seconds = 1
      If delta > 0 AndAlso Output < 1000 Then ' Do not sum if the output is maxxed
        IntegralSum += delta
        IntegralCount += 1
        If IntegralSum > (1000 * 60) Then IntegralSum = (1000 * 60)
      End If
      If delta < 0 AndAlso Output > -1000 Then
        IntegralSum += delta
        IntegralCount += 1
        If IntegralSum < -(1000 * 60) Then IntegralSum = -(1000 * 60)
      End If
    End If

    ' Return integral (basically a percentage of the running error)
    Return CInt((IntegralSum / 60) * (Integral / 1000))
  End Function

  ' Use hold margin and prop band to calculate how close we need to be to the setpoint before starting integral 
  Private Function IntegralStartRampMax(delta As Integer) As Boolean
    Dim startMargin = MinMax(Math.Max(HoldMargin * 2, PropBand \ 2), 10, 40)
    Return (Math.Abs(delta) <= startMargin)
  End Function

  ' Bump integral when going from ramp to hold so we don't get a sudden drop
  Private Sub IntegralGradientBump()
    ' No adjustment for the gradient 
    If GradientMax <= 0 Then Exit Sub

    ' Don't do this on a max ramp
    If Ramp.RampMax Then Exit Sub

    ' Bump the integral sum based on the ramp gradient term (75 %)
    Dim integralBump = CInt((1000 * 60) * (Gradient / GradientMax) * (1000 / Integral) * Me.IntegralBump / 1000)
    IntegralSum += integralBump
  End Sub

  Public Sub ResetIntegral()
    IntegralSum = 0
    IntegralCount = 0
    IntegralTerm = 0
  End Sub

  Private Function GetGradientTerm(currentTemp As Integer) As Integer
    ' No adjustment for the gradient 
    If GradientMax <= 0 Then Return 0

    ' Calculate standard gradient term
    Dim standardTerm = (Gradient / GradientMax) * 1000

    ' Just use standard term if taper margin not set
    If GradientTaperMargin = 0 Then Return CInt(standardTerm)

    'Taper off the gradient term as the setpoint approaches the final temp (limit the taper)
    Dim gradientRemaining = Math.Abs(FinalTemp - TargetTemp)
    If gradientRemaining < GradientTaperMargin Then
      Dim gradientTaper = MinMax(gradientRemaining / MinMax(GradientTaperMargin, 0, 100), 0.5, 1)
      Return CInt(standardTerm * gradientTaper)
    Else
      Return CInt(standardTerm)
    End If
  End Function

  'Ramp gradient term down during start of hold 
  Private Function GetGradientTermHold(secondsRemaining As Integer, secondsTotal As Integer) As Integer
    ' No adjustment for the gradient 
    If GradientMax <= 0 Then Return 0

    ' Don't do this on a max ramp
    If Ramp.RampMax Then Return 0

    ' Don't do this if parameter = 0
    If secondsTotal <= 0 Then Return 0

    ' We're done
    If secondsRemaining <= 0 Then Return 0

    ' Taper gradient term down so we don't get a sudden drop when we go to a hold
    Dim standardTerm = (Gradient / GradientMax) * 1000
    Return CInt(standardTerm * (secondsRemaining / secondsTotal))
  End Function


  Private Function GetBalanceTerm(currentTemp As Integer) As Integer
    Dim balanceDelta = TargetTemp - 1200
    Return CInt(balanceDelta * Balance / 1000)
  End Function


  Public Sub Pause()
    If State = EPidState.Ramp Then
      State = EPidState.RampPause
      Ramp.Pause()
    ElseIf State = EPidState.Hold Then
      State = EPidState.HoldPause
    End If
  End Sub

  Public Sub Restart()
    If State = EPidState.RampPause Then
      State = EPidState.Ramp
      Ramp.Restart()
    ElseIf State = EPidState.HoldPause Then
      State = EPidState.Hold
    End If
  End Sub

  Public Sub Reset(ByVal currentTemp As Integer)
    ResetIntegral()
    Ramp.Restart(currentTemp)
  End Sub

  Public Sub Cancel()
    Output = 0
    TargetTemp = 0
    ResetIntegral()
    State = EPidState.Off
  End Sub

  ReadOnly Property IsPaused As Boolean
    Get
      Return State = EPidState.RampPause OrElse State = EPidState.HoldPause
    End Get
  End Property

  ReadOnly Property IsRamp As Boolean
    Get
      Return Ramp.IsRamp
    End Get
  End Property

  ReadOnly Property IsHold As Boolean
    Get
      Return Ramp.IsDone
    End Get
  End Property

  ReadOnly Property IsHeat As Boolean
    Get
      Return State <> EPidState.Off AndAlso FinalTemp >= StartTemp
    End Get
  End Property

  ReadOnly Property IsCool As Boolean
    Get
      Return State <> EPidState.Off AndAlso FinalTemp < StartTemp
    End Get
  End Property

  Private output_ As Integer
  Property Output As Integer
    Get
      If State = EPidState.Ramp OrElse State = EPidState.RampToHold OrElse State = EPidState.Hold Then Return output_
      Return 0
    End Get
    Private Set(value As Integer)
      output_ = value
    End Set
  End Property
End Class
