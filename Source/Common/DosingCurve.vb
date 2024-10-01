
Public Class DosingCurve

  Public Seconds As Integer
  Public Curve As Integer
  Public StartLevel As Integer
  Public EndLevel As Integer

  Public Timer As New Timer

  Public Sub Start(ByVal seconds As Integer, ByVal curve As Integer, ByVal startLevel As Integer, ByVal EndLevel As Integer)
    Me.Seconds = seconds
    Me.Curve = curve
    Me.StartLevel = startLevel
    Me.EndLevel = EndLevel

    Me.Timer.Seconds = seconds
  End Sub

  Public Sub Pause()
    Me.Timer.Pause()
  End Sub

  Public Sub Restart()
    Me.Timer.Restart()
  End Sub

  Public ReadOnly Property Setpoint() As Integer
    Get
      Select Case Curve
        Case 0
          Return Linear()
        Case 1, 3, 5, 7, 9
          Return OddCurve()
        Case 2, 4, 6, 8
          Return EvenCurve()
        Case Else
          Return 0
      End Select
    End Get
  End Property

  Private Function Linear() As Integer
    Try

      'If Seconds is 0 then return 0 
      If Seconds = 0 Then Return 0

      'Calculate the level range 
      Dim levelRange As Integer = StartLevel - EndLevel

      'Calculate elapsed time
      Dim elapsedSeconds As Integer = Seconds - Timer.Seconds

      'Calculate how much time has elapsed as a fraction of the total dose time
      Dim elapsedFraction = (Seconds - Timer.Seconds) / Seconds

      'Calculate the level we should have transfered so far
      Dim transferAmount As Integer = CInt(levelRange * elapsedFraction)

      'Calculate setpoint
      Dim setpoint As Integer = (StartLevel - transferAmount) + EndLevel

      'Return the calculated setpoint
      Return setpoint

    Catch ex As Exception
      'Ignore errors
    End Try
    Return 0
  End Function

  Private Function EvenCurve() As Integer
    Try

      'If Seconds is 0 then return 0 
      If Seconds = 0 Then Return 0

      'Calculate the level range 
      Dim levelRange As Integer = StartLevel - EndLevel

      'Calculate remaining time
      Dim remainingSeconds As Integer = Seconds - Timer.Seconds

      'Calculate remaining time as a fraction of the total dose time
      Dim remainingFraction = 1 - ((Seconds - Timer.Seconds) / Seconds)

      'Calculate maximum curve
      Dim maxCurve As Double = 1 - Math.Sqrt(1 - (remainingFraction ^ 3))

      'Calculate 
      Dim scaledCurve As Double = (((8 - Curve) * remainingSeconds) + (Curve * maxCurve)) / 8

      'Calculate the level we should have transfered so far
      Dim transferAmount As Integer = CInt(1 - (levelRange * scaledCurve))

      'Calculate setpoint
      Dim setpoint As Integer = (StartLevel - transferAmount) + EndLevel

      'Return the calculated setpoint
      Return setpoint
    Catch ex As Exception
      'Ignore errors
    End Try
    Return 0
  End Function

  Private Function OddCurve() As Integer
    Try

      'If Seconds is 0 then return 0 
      If Seconds = 0 Then Return 0

      'Calculate the level range 
      Dim levelRange As Integer = StartLevel - EndLevel

      'Calculate elapsed time
      Dim elapsedSeconds As Integer = Seconds - Timer.Seconds

      'Calculate how much time has elapsed as a fraction of the total dose time
      Dim elapsedFraction = (Seconds - Timer.Seconds) / Seconds

      'Calculate maximum curve
      Dim maxCurve As Double = 1 - Math.Sqrt(1 - (elapsedFraction ^ 3))

      'Calculate 
      Dim scaledCurve As Double = (((9 - Curve) * elapsedSeconds) + ((Curve + 1) * maxCurve)) / 10

      'Calculate the level we should have transfered so far
      Dim transferAmount As Integer = CInt(levelRange * scaledCurve)

      'Calculate setpoint
      Dim setpoint As Integer = (StartLevel - transferAmount) + EndLevel

      'Return the calculated setpoint
      Return setpoint
    Catch ex As Exception
      'Ignore errors
    End Try
    Return 0
  End Function

End Class
