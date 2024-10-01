Public Class LevelFloat : Inherits MarshalByRefObject
  'Returns the Level value, true for level "Made", depending on Input status and selection of 
  ' Normally-Closed or Normally-Open Float Switch definition
  'Includes an optional delay for signal bounce, in milliseconds

  Public Timer As New Timer

  Public ReadOnly Property TimerString() As String
    Get
      Return Timer.ToStringMs
    End Get
  End Property

  'If Input is "On" when level is below float:
  Private normallyClosed_ As Boolean
  Public Property NormallyClosed() As Boolean
    Get
      Return normallyClosed_
    End Get
    Set(ByVal value As Boolean)
      normallyClosed_ = value
    End Set
  End Property

  'Millisecond delay to compensate for signal bounce:
  Private signalDelayTimeMs_ As Integer
  Public Property SignalDelayTimeMs() As Integer
    Get
      Return SignalDelayTimeMs_
    End Get
    Set(ByVal value As Integer)
      signalDelayTimeMs_ = value
    End Set
  End Property

  'Reference to Input Status passed
  Private levelInput_ As Boolean
  Public ReadOnly Property LevelInput() As Boolean
    Get
      Return levelInput_
    End Get
  End Property

  Private levelAboveFloat_ As Boolean
  Public ReadOnly Property IsLevelAboveFloat() As Boolean
    Get
      Return levelAboveFloat_
    End Get
  End Property

  Public Sub CalibrateLevel(ByVal Input As Boolean, ByVal normallyClosed As Boolean, ByVal delayTime As Integer)
    Try

      'Set the input status for history reference
      Me.levelInput_ = Input
      Me.normallyClosed_ = normallyClosed
      Me.signalDelayTimeMs_ = delayTime

      If normallyClosed_ Then
        'Normally-Closed Float Switch > Input On when level below float, Off when level above float
        If Input Then Timer.Milliseconds = signalDelayTimeMs_
        If (signalDelayTimeMs_ > 0) Then
          'Using Signal Delay Option - noise compensation
          If Timer.Finished Then
            levelAboveFloat_ = True
          Else : levelAboveFloat_ = False
          End If
        Else
          'Not Using Signal Delay Option
          If Not Input Then
            levelAboveFloat_ = True
          Else : levelAboveFloat_ = False
          End If
        End If

      Else
        'Normally-Open Float Switch > Input Off when level below float, On when level above float
        If Not Input Then Timer.Milliseconds = signalDelayTimeMs_
        If (signalDelayTimeMs_ > 0) Then
          'Using Signal Delay Option - noise compensation
          If Timer.Finished Then
            levelAboveFloat_ = True
          Else : levelAboveFloat_ = False
          End If
        Else
          'Not Using Signal Delay Option
          If Input Then
            levelAboveFloat_ = True
          Else : levelAboveFloat_ = False
          End If
        End If

      End If

    Catch ex As Exception
      'Ignore errors
    End Try
  End Sub

End Class
