Public Class RampGenerator

  Public Enum ERampState ' Are we heating, cooling or holding temp
    Off
    Up
    Down
    Done
  End Enum
  Property State As ERampState
  Property Timer As New Timer

  Property StartValue As Integer
  Property FinalValue As Integer
  Property RampRate As Integer

  Property RampDelta As Integer
  Property RampSeconds As Integer
  Property RampMax As Boolean

  Sub Start(startValue As Integer, finalValue As Integer, rampRate As Integer)
    Me.StartValue = startValue
    Me.FinalValue = finalValue
    Me.RampRate = rampRate

    ' If this is a max rate setpoint = final value
    If rampRate = 0 OrElse rampRate = 99 OrElse (finalValue = startValue) Then
      RampDelta = 0
      RampSeconds = 0
      RampMax = True

      State = ERampState.Done
      Timer.Cancel()
    Else
      RampDelta = Math.Abs(finalValue - startValue)
      RampSeconds = (RampDelta \ rampRate) * 60
      RampMax = False

      State = ERampState.Done
      If startValue < finalValue Then State = ERampState.Up
      If startValue > finalValue Then State = ERampState.Down
      Timer.Seconds = RampSeconds
    End If
  End Sub

  Function Setpoint() As Integer
    ' Switch to hold once the ramp is done
    If Timer.Finished Then State = ERampState.Done

    Select Case State
      Case ERampState.Off : Return 0
      Case ERampState.Up : Return StartValue + GetRampValue()
      Case ERampState.Down : Return StartValue - GetRampValue()
      Case ERampState.Done : Return FinalValue
    End Select
  End Function

  Private Function GetRampValue() As Integer
    ' Calculate current ramp position (from start)
    Dim rampPercent, rampPosition As Double
    If RampSeconds > 0 Then
      rampPercent = ((RampSeconds - Timer.Seconds) / RampSeconds)
      rampPosition = (RampDelta * rampPercent)
    Else
      rampPercent = 1
      rampPosition = RampDelta
    End If

    Return CInt(rampPosition)
  End Function

  Sub Cancel()
    State = ERampState.Off
    Timer.Cancel()
  End Sub

  Sub Pause()
    Timer.Pause()
  End Sub

  Sub Restart()
    Timer.Restart()
  End Sub

  Sub Restart(currentTemp As Integer)
    'TODO....
  End Sub

  Sub Reset(currentValue As Integer)
    Start(currentValue, FinalValue, RampRate)
  End Sub


  ReadOnly Property IsOn As Boolean
    Get
      Return State <> ERampState.Off
    End Get
  End Property

  ReadOnly Property IsDone As Boolean
    Get
      Return State = ERampState.Done
    End Get
  End Property

  ReadOnly Property IsRamp() As Boolean
    Get
      Return IsRampUp OrElse IsRampDown
    End Get
  End Property

  ReadOnly Property IsRampUp As Boolean
    Get
      Return State = ERampState.Up
    End Get
  End Property

  ReadOnly Property IsRampDown As Boolean
    Get
      Return State = ERampState.Down
    End Get
  End Property

  ReadOnly Property IsPaused As Boolean
    Get
      Return Timer.Paused
    End Get
  End Property

End Class