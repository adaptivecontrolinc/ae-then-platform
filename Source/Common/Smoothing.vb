' Analog input smoothing with rejection.
'   Ignore input values that are outside the specified range or change more than the specified amount on consecutive reads.
Public Class Smoothing

  Private smoothingDefault As Integer = 24
  Private values(smoothingDefault) As Integer

  Property Input As Integer

  Property Smoothing As Integer        ' smoothing rate
  Property ValueMax As Integer         ' reject values lower than this
  Property ValueMin As Integer         ' reject values higher than this
  Property ValueDeltaMax As Integer    ' reject values that change more than this on consecutive reads

  Property Value As Integer            ' smoothed value
  Property ValueSum As Integer         ' sum of valid values
  Property ValueCount As Integer       ' count of valid values, should = smoothing if no values are rejected

  Property ErrorNoValues As Boolean    ' set this if there are no valid values

  Property Raw As Integer              ' hide this when the code has been debugged
  Property RawSum As Integer
  Property RawCount As Integer

  Property ScanTimer As New TimerUp
  Property ScanMs As Integer

  Sub New(valueMin As Integer, valueMax As Integer, valueDeltaMax As Integer)
    Me.ValueMin = valueMin
    Me.ValueMax = valueMax
    Me.ValueDeltaMax = valueDeltaMax
  End Sub

  Function Smooth(input As Integer, smoothing As Integer) As Integer
    Me.Input = input
    If smoothing <= 0 Then smoothing = smoothingDefault
    Me.Smoothing = smoothing

    Static currentIndex As Integer        ' 
    Static arrayFull As Boolean           ' set to true when we have filled the values array

    ' Initialize the array and variables on a resize or startup
    If values Is Nothing OrElse values.Length <> smoothing Then
      values = New Integer(smoothing - 1) {}
      currentIndex = 0
      arrayFull = False

      Value = 0
      ValueSum = 0
      ValueCount = 0
      ErrorNoValues = False
    End If

    ' Set array full flag
    If currentIndex = values.GetUpperBound(0) Then arrayFull = True

    ' Wrap the current index - set back to 0 when we go past the last element 
    '  Start timer when we wrap so we can time how long it takes to fill the array
    If currentIndex > values.GetUpperBound(0) Then
      ScanMs = ScanTimer.TimeElapsedMs
      ScanTimer.Start()
      currentIndex = 0
    End If

    ' Store the input
    values(currentIndex) = input
    currentIndex += 1

    ' Return the raw input if we have not filled the array, do not smooth until the array is fully populated with input values
    If Not arrayFull Then Return input

    ' Sum all values to calculate raw average
    Dim sumRaw As Integer
    For i As Integer = 0 To values.GetUpperBound(0)
      sumRaw += values(i)
    Next

    ' Temporary - just here for verification
    RawSum = sumRaw
    RawCount = values.Length
    Raw = RawSum \ RawCount


    ' Sum all valid values
    Dim sumValid As Integer, countValid As Integer
    For i As Integer = 0 To values.GetUpperBound(0)
      If ValueValid(i) Then
        sumValid += values(i)
        countValid += 1
      End If
    Next

    ' Temporary - just here for verification
    ValueSum = sumValid
    ValueCount = countValid

    ' Calculate the average if we have any valid values
    If countValid > 0 Then
      Value = sumValid \ countValid
      ErrorNoValues = False
      Return Value
    Else
      Value = 0
      ErrorNoValues = True
      Return input
    End If

  End Function

  Private Function ValueValid(index As Integer) As Boolean
    ' Check the value against the min max limits
    If values(index) < ValueMin OrElse values(index) > ValueMax Then Return False

    '' Get the index of the previous value
    'Dim prevIndex = index - 1
    'If prevIndex < 0 Then prevIndex = values.GetUpperBound(0)

    '' Check the change in value if the change is too great reject the value
    'Dim delta = Math.Abs(values(index) - values(prevIndex))
    'If delta > ValueDeltaMax Then Return False

    Return True
  End Function

End Class
