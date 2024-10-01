#If 0 Then
    Public Module TickCountModule
    Friend TickCount As UInt32
    End Module
#End If

Public Class OnDelay
  Private ReadOnly timer_ As New Timer
  Private lastInput_ As Boolean

  Default Public ReadOnly Property Run(ByVal input As Boolean, ByVal delay As Integer) As Boolean
    Get
      Return Run(input, New TimeSpan(delay * TimeSpan.TicksPerSecond))
    End Get
  End Property
  Default Public ReadOnly Property Run(ByVal input As Boolean, ByVal delay As TimeSpan) As Boolean
    Get
      If input Then
        If Not lastInput_ Then timer_.TimeRemainingMs = CType(delay.TotalMilliseconds, Integer)
        Run = timer_.Finished
      Else
        timer_.Cancel()
      End If
      lastInput_ = input
    End Get
  End Property
End Class

' Timer - down timer
Public Class Timer : Inherits MarshalByRefObject
  Private interval_, startTickCount_, pauseTickCount_ As UInt32, paused_, finished_ As Boolean

  Friend Property TimeRemainingMs() As Integer
    Get
      If finished_ Then Return 0
      Dim t As UInt32
      If paused_ Then
        t = pauseTickCount_
      Else
        t = TickCount
      End If
      ' Check to see if we have wrapped
      If t >= startTickCount_ Then
        Return Math.Max(CType(interval_ - (t - startTickCount_), Integer), 0)   ' Don't return a negative
      Else
        Return Math.Max(CType(interval_ - (t + (UInt32.MaxValue - startTickCount_)), Integer), 0)  ' Don't return a negative
      End If
    End Get
    Set(ByVal value As Integer)
      interval_ = CType(value, UInt32)
      startTickCount_ = TickCount
      paused_ = False : finished_ = False
    End Set
  End Property

  Public Property TimeRemaining() As Integer
    Get
      Return (TimeRemainingMs + 500) \ 1000
    End Get
    Set(ByVal value As Integer)
      TimeRemainingMs = value * 1000
    End Set
  End Property

  ' TimeRemainingMs and MilliSeconds are the same
  Friend Property Milliseconds As Integer
    Get
      Return TimeRemainingMs
    End Get
    Set(ByVal value As Integer)
      TimeRemainingMs = value
    End Set
  End Property

  ' TimeRemaining and Seconds are the same
  Friend Property Seconds As Integer
    Get
      Return TimeRemaining
    End Get
    Set(ByVal value As Integer)
      TimeRemaining = value
    End Set
  End Property

  Friend Property Minutes() As Integer
    Get
      Return (TimeRemainingMs + 30000) \ 60000
    End Get
    Set(ByVal value As Integer)
      TimeRemainingMs = value * 60000
    End Set
  End Property

  ReadOnly Property Finished() As Boolean
    Get
      ' This could get jammed if tick count wraps
      'If Not finished_ AndAlso Not paused_ AndAlso TickCount - startTickCount_ >= interval_ Then finished_ = True

      If Not finished_ AndAlso Not paused_ AndAlso TimeRemainingMs <= 0 Then finished_ = True
      Return finished_
    End Get
  End Property

  Private Sub CheckFinished()
    If Not paused_ AndAlso TickCount - startTickCount_ >= interval_ Then finished_ = True
  End Sub

  Public Sub Pause()
    'If we're already finshed or paused then ignore
    If Finished OrElse paused_ Then Exit Sub
    paused_ = True : pauseTickCount_ = TickCount
  End Sub

  Public Sub Restart()
    'If we're not paused then ignore
    If finished_ OrElse Not paused_ Then Exit Sub
    Dim lostTime As UInt32 = TickCount - pauseTickCount_
    startTickCount_ += lostTime
    paused_ = False
  End Sub

  Public Sub Cancel()
    TimeRemainingMs = 0
  End Sub

  Public ReadOnly Property Paused() As Boolean
    Get
      Return paused_
    End Get
  End Property

  Public Overloads Function ToString(padSpaces As Integer) As String
    If padSpaces > 0 Then
      Dim timerString As String = Me.ToString
      Return timerString.PadLeft(timerString.Length + padSpaces)
    End If
    Return Me.ToString
  End Function

  Public Overrides Function ToString() As String
    With TimeSpan.FromTicks(TimeRemainingMs * TimeSpan.TicksPerMillisecond)
      Select Case .TotalSeconds
        Case Is >= 86400
          Return .Days.ToString("00") & ":" & .Hours.ToString("00") & "h"
        Case 3600 To 86399
          Return .Hours.ToString("00") & ":" & .Minutes.ToString("00") & "m"
        Case 1 To 3599
          Return .Minutes.ToString("00") & ":" & .Seconds.ToString("00") & "s"
        Case Is <= 0
          Return "00:00s"
        Case Else
          Return "00:00s"
      End Select
    End With
  End Function

  Friend Function ToStringMs(padSpaces As Integer) As String
    If padSpaces > 0 Then
      Dim timerString As String = Me.ToStringMs
      Return timerString.PadLeft(timerString.Length + padSpaces)
    End If
    Return Me.ToString
  End Function

  Friend Function ToStringMs() As String
    Try
      Dim ms As Integer = Milliseconds

      'If there is more than 10 seconds left then just return the normal formatting
      If ms > 9999 Then Return Me.ToString

      'Return the time left formatted for milliseonds
      If ms > 0 Then
        Return ms.ToString & "ms"
      Else
        Return "00:00s"
      End If

    Catch ex As Exception
      'Ignore errors
    End Try
    Return Nothing
  End Function

End Class

' -------------------------------------------------------------------------
' Timer - up timer
Public Class TimerUp : Inherits MarshalByRefObject
  Private startTickCount_, pauseInterval_ As UInt32, started_, paused_ As Boolean

  Public Sub Start()
    started_ = True
    startTickCount_ = TickCount
    paused_ = False
  End Sub

  Public Sub [Stop]()
    started_ = False
  End Sub

  ' TimeRemaining and Seconds are identical
  Public ReadOnly Property TimeElapsed() As Integer
    Get
      Return TimeElapsedMs \ 1000 ' return in seconds
    End Get
  End Property

  Friend Property Seconds As Integer
    Get
      Return TimeElapsed
    End Get
    Set(value As Integer)
      ' TODO - Update StartTick
      startTickCount_ = CUInt(TickCount - value)
    End Set
  End Property

  Friend ReadOnly Property TimeElapsedMs() As Integer
    Get
      If Not started_ Then Return 0
      If paused_ Then Return CType(pauseInterval_, Integer)
      Return CType(TickCount - startTickCount_, Integer)
    End Get
  End Property



#If 0 Then
  'Total time paused - allows us to calculate time elapsed that doesn't include pause time
  Private timePaused_ As New TimeSpan
  Friend Property TimePaused() As TimeSpan
    Get
      Return timePaused_
    End Get
    Set(ByVal value As TimeSpan)
      timePaused_ = value
    End Set
  End Property


#End If

  Public Sub Pause()
    'If we're already paused then ignore
    If Not paused_ Then paused_ = True : pauseInterval_ = TickCount - startTickCount_
  End Sub

  Public Sub Restart()
    'If we're not paused then ignore
    If Not paused_ Then Exit Sub

    startTickCount_ = TickCount - pauseInterval_
    paused_ = False
  End Sub

  Public Overrides Function ToString() As String
    With TimeSpan.FromTicks(TimeElapsedMs * TimeSpan.TicksPerMillisecond)
      Select Case .TotalSeconds
        Case Is >= 86400
          Return .Days.ToString("00") & ":" & .Hours.ToString("00") & "h"
        Case 3600 To 86399
          Return .Hours.ToString("00") & ":" & .Minutes.ToString("00") & "m"
        Case 1 To 3599
          Return .Minutes.ToString("00") & ":" & .Seconds.ToString("00") & "s"
        Case Is <= 0
          Return "00:00s"
        Case Else
          Return "00:00s"
      End Select
    End With
  End Function

  Public ReadOnly Property IsRunning As Boolean
    Get
      Return started_
    End Get
  End Property

  Public ReadOnly Property IsPaused As Boolean
    Get
      Return paused_
    End Get
  End Property

End Class
