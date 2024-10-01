Public Class Flowmeter

  Private counterMax As Integer = 65536   ' = 2^16
  Private counter_ As Integer = 0

  Public Property Counter() As Integer
    Get
      Return counter_
    End Get
    Set(ByVal value As Integer)
      counter_ = value
      'Increment wrap counter 
      If Counter < CounterPrevious Then CounterWraps += 1
      CounterPrevious = value
    End Set
  End Property

  Private counterPrevious_ As Integer = 0
  Public Property CounterPrevious() As Integer
    Get
      Return counterPrevious_
    End Get
    Private Set(ByVal value As Integer)
      counterPrevious_ = value
    End Set
  End Property

  Private counterWraps_ As Integer = 0
  Public Property CounterWraps() As Integer
    Get
      Return counterWraps_
    End Get
    Private Set(ByVal value As Integer)
      counterWraps_ = value
    End Set
  End Property

  Private pausedWraps_ As Integer = 0
  Public Property PausedWraps() As Integer
    Get
      Return pausedWraps_
    End Get
    Private Set(ByVal value As Integer)
      pausedWraps_ = value
    End Set
  End Property

  Private startCount_ As Integer
  Public Property StartCount() As Integer
    Get
      Return startCount_
    End Get
    Private Set(ByVal value As Integer)
      startCount_ = value
    End Set
  End Property

  Public ReadOnly Property TotalCount() As Integer
    Get
      Return ((Counter - StartCount) + (CounterWraps * counterMax)) - PausedCount
    End Get
  End Property

  Private countsPerGallon_ As Integer
  Public Property CountsPerGallon As Integer
    Get
      Return countsPerGallon_
    End Get
    Set(ByVal value As Integer)
      countsPerGallon_ = value
    End Set
  End Property

  Private flowLossTime_ As Integer
  Public Property FlowLossTime() As Integer
    Get
      Return flowLossTime_
    End Get
    Set(ByVal value As Integer)
      flowLossTime_ = MinMax(value, 0, 15)
    End Set
  End Property

  Private flowMinRate_ As Integer
  Public Property FlowMinRate() As Integer
    Get
      Return flowMinRate_
    End Get
    Set(ByVal value As Integer)
      flowMinRate_ = MinMax(value, 0, 1000)
    End Set
  End Property

  Private paused_ As Boolean
  Public Property Paused() As Boolean
    Get
      Return paused_
    End Get
    Set(ByVal value As Boolean)
      paused_ = value
    End Set
  End Property

  Private pausedCounter_ As Integer
  Public Property PausedCounter() As Integer
    Get
      Return pausedCounter_
    End Get
    Set(ByVal value As Integer)
      pausedCounter_ = value
    End Set
  End Property

  Public ReadOnly Property PausedCount() As Integer
    Get
      If Not Paused Then Return Nothing
      Return (Counter - PausedCounter) + (PausedWraps * counterMax)
    End Get
  End Property

  Private pausedCountTotal_ As Integer
  Public Property PausedCountTotal() As Integer
    Get
      Return pausedCountTotal_
    End Get
    Set(ByVal value As Integer)
      pausedCountTotal_ = value
    End Set
  End Property

  Public Const LitersToGallonsUS As Double = 0.264172
  Public ReadOnly Property Gallons() As Double
    Get
      If countsPerGallon_ <> 0 Then
        Return TotalCount / CountsPerGallon
      Else
        Return 0
      End If
      If CountsPerGallon = 0 Then Return 0
    End Get
  End Property

  Public ReadOnly Property iGallons() As Integer
    Get
      Return Convert.ToInt32(Gallons * 10)
    End Get
  End Property

  Private previousVolume_ As Double
  Private totalVolume_ As Double
  Public ReadOnly Property TotalVolume() As Double
    Get
      Return totalVolume_
    End Get
  End Property
  Public ReadOnly Property iTotalVolume() As Integer
    Get
      Return Convert.ToInt32(totalVolume_ * 10)
    End Get
  End Property

  Public ReadOnly Property AlarmFlowLoss() As Boolean
    Get
      Return (flowLossTime_ > 0) AndAlso (flowMinRate_ > 0) AndAlso flowTimer_.Finished
    End Get
  End Property

  Private flowTimer_ As New Timer
  Public Property FlowTimer As Timer
    Get
      Return flowTimer_
    End Get
    Private Set(ByVal value As Timer)
      flowTimer_ = value
    End Set
  End Property

  Private timer_ As New Timer
  Public Property Timer As Timer
    Get
      Return timer_
    End Get
    Private Set(ByVal value As Timer)
      timer_ = value
    End Set
  End Property

  Private totalCountPrevious_ As Integer

  Private pulseRate_ As Double
  Public ReadOnly Property FlowRate() As Double
    Get
      If CountsPerGallon = 0 Then Return 0
      Return Math.Min(Math.Round((pulseRate_ * 60 / countsPerGallon_), 2), Double.MaxValue)
    End Get
  End Property

  Public Sub Run(ByVal currentlyFilling As Boolean)

    If timer_.Finished Then
      pulseRate_ = TotalCount - totalCountPrevious_
      totalCountPrevious_ = TotalCount
      Timer.Seconds = 1
    End If

    If currentlyFilling Then
      If (pulseRate_ > flowMinRate_) Then FlowTimer.Seconds = flowLossTime_
    Else : FlowTimer.Seconds = flowLossTime_
    End If

    'Keep a running total of volume used whenever the fill valve is open during a cycle (reset at end of cycle with sub "ResetTotals")
    Static wasFilling As Boolean
    If currentlyFilling Then
      totalVolume_ = previousVolume_ + Gallons
    Else
      If wasFilling AndAlso Not currentlyFilling Then
        previousVolume_ = totalVolume_
      End If
    End If
    wasFilling = currentlyFilling

  End Sub

  Public Sub Reset()
    totalVolume_ += Gallons
    StartCount = Counter
    CounterWraps = 0
    PausedCounter = Counter
    PausedWraps = 0
  End Sub

  Public Sub ResetTotals()
    previousVolume_ = 0
    totalVolume_ = 0
    Reset()
  End Sub

  Public Sub Pause()
    If Paused Then
      Exit Sub
    Else
      Paused = True
      PausedWraps = 0
      PausedCounter = Counter
    End If
  End Sub

  Public Sub Restart()
    If Not Paused Then
      Exit Sub
    Else
      Paused = False
      PausedWraps = 0
      PausedCountTotal += PausedCount
    End If
  End Sub

End Class
