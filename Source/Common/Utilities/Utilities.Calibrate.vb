'Version 2022-02-03

Namespace Utilities.Calibrate

  Public Module Calibrate

    Function CalibrateAnalogInput(ByVal value As Integer, ByVal minimumValue As Integer, ByVal maximumValue As Integer) As Integer
      'Convert signal input to a percentage
      Try
        'Check to see that maximumValue has been set (zero could cause workingRange to be 0 = divide by zero error)
        If maximumValue = 0 Then Return 0

        'Check to see if we're outside limits
        If value <= minimumValue Then Return 0
        If value >= maximumValue Then Return 1000

        'Calculate level
        Dim workingRange As Integer = maximumValue - minimumValue
        Dim workingValue As Integer = value - minimumValue

        If workingRange <> 0 Then Return Convert.ToInt32((workingValue / workingRange) * 1000)

      Catch ex As Exception
        'Ignore error
      End Try
      Return 0
    End Function

    Function CalibrateAnalogInput(ByVal value As Integer, ByVal minimumValue As Integer, ByVal maximumValue As Integer, ByVal maximumRange As Integer) As Integer
      'Convert signal input to a range
      Try
        'Check to see that maximumValue has been set (zero could cause workingRange to be 0 = divide by zero error)
        If maximumValue = 0 Then Return 0
        If maximumRange = 0 Then Return 0 ' Force operator to set maximum range

        'Check to see if we're outside limits
        If value <= minimumValue Then Return 0
        If value >= maximumValue Then Return maximumRange

        'Calculate level
        Dim workingRange As Integer = maximumValue - minimumValue
        Dim workingValue As Integer = value - minimumValue

        If workingRange <> 0 Then Return Convert.ToInt32((workingValue / workingRange) * maximumRange)

      Catch ex As Exception
        'Ignore error
      End Try
      Return 0
    End Function

    Function CalibrateTemperature(ByVal value As Integer, ByVal minimumValue As Integer, ByVal maximumValue As Integer) As Integer
      'Convert analog input to a temperature
      Try
        'Check to see if we're outside limits
        If value <= minimumValue Then Return 0
        If value >= maximumValue Then Return maximumValue

        'Calculate temperature
        Dim workingRange As Integer = maximumValue - minimumValue
        Dim workingValue As Integer = value - minimumValue
        Return Convert.ToInt32((workingValue / 1000) * maximumValue)

      Catch ex As Exception
        'Ignore error
      End Try
      Return 0
    End Function

    Function CalibrateTemperature(ByVal value As Integer, ByVal minimumValue As Integer, ByVal maximumValue As Integer, ByVal maximumRange As Integer) As Integer
      'Convert analog input to a temperature
      Try
        'Check to see if we have parameters
        If maximumValue = 0 Then Return 0
        If maximumRange = 0 Then Return 0

        'Check to see if we're outside limits
        If value <= minimumValue Then Return 0
        If value >= maximumValue Then Return maximumRange

        'Calculate working value (percent including offset)
        Dim workingRange As Integer = maximumValue - minimumValue
        Dim workingOffset As Integer = value - minimumValue

        Return MinMax(Convert.ToInt32(workingOffset * maximumRange \ workingRange), 0, maximumRange)

      Catch ex As Exception
        'Ignore error
      End Try
      Return 0
    End Function

    Function LineFeed() As String
      Return Convert.ToChar(13) & Convert.ToChar(10)
    End Function
  End Module
End Namespace
