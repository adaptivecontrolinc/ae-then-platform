' Up Timer
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

  Friend ReadOnly Property Seconds As Integer
    Get
      Return TimeElapsed
    End Get
  End Property

  Friend ReadOnly Property TimeElapsedMs() As Integer
    Get
      If Not started_ Then Return 0
      If paused_ Then Return CType(pauseInterval_, Integer)
      Return CType(TickCount - startTickCount_, Integer)
    End Get
  End Property

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