Public Class Ticker : Inherits MarshalByRefObject

  Public Sub New(ByVal seconds As Integer)
    Me.Seconds = seconds
    Me.Milliseconds = 0

    Timer.Milliseconds = TotalMilliseconds
  End Sub

  Public Sub New(ByVal seconds As Integer, ByVal milliseconds As Integer)
    Me.Seconds = seconds
    Me.Milliseconds = milliseconds

    Timer.Milliseconds = TotalMilliseconds
  End Sub

  Public Function Tick() As Boolean
    'Returns true once every tick period
    If Timer.Finished Then
      Timer.Milliseconds = TotalMilliseconds
      Return True
    End If
    Return False
  End Function

  Public Function Tick(ByVal seconds As Integer) As Boolean
    'Set new value and returb tick - the new value will be used the next time the timer resets
    Me.Seconds = seconds
    Return Tick
  End Function

  Public Function Tick(ByVal seconds As Integer, ByVal milliseconds As Integer) As Boolean
    'Set new value and returb tick - the new value will be used the next time the timer resets
    Me.Seconds = seconds
    Me.Milliseconds = milliseconds
    Return Tick
  End Function

  Private timer_ As New Timer
  Public Property Timer() As Timer
    Get
      Return timer_
    End Get
    Private Set(ByVal value As Timer)
      timer_ = value
    End Set
  End Property

  Private seconds_ As Integer
  Public Property Seconds() As Integer
    Get
      Return seconds_
    End Get
    Set(ByVal value As Integer)
      seconds_ = value
    End Set
  End Property

  Private milliseconds_ As Integer
  Public Property Milliseconds() As Integer
    Get
      Return milliseconds_
    End Get
    Set(ByVal value As Integer)
      milliseconds_ = value
    End Set
  End Property

  Private ReadOnly Property TotalMilliseconds() As Integer
    Get
      Return (Seconds * 1000) + Milliseconds
    End Get
  End Property

End Class