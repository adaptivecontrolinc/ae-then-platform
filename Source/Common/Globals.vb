Module Globals

  Private ReadOnly newLineBytes As Byte() = {&HD, &HA}
  Public Property NewLine As String = System.Text.ASCIIEncoding.ASCII.GetString(newLineBytes)

  Public ReadOnly Property DefaultTimeStamp As String
    Get
      Return Date.Now.ToString("HHmmss")
    End Get
  End Property

  Public ReadOnly Property DefaultDateAndTimeStamp As String
    Get
      Return Date.Now.ToString("yyMMddHHmmss")
    End Get
  End Property

  Public Function TimerString(ByVal secs As Integer) As String
    Dim hours As Integer = secs \ 3600, minutes As Integer = (secs Mod 3600) \ 60, _
        seconds As Integer = (secs Mod 60)

    ' Hours and minutes
    If hours > 0 Then Return hours.ToString("00", Globalization.CultureInfo.InvariantCulture) & ":" _
                             & minutes.ToString("00", Globalization.CultureInfo.InvariantCulture) & "m"
    ' Minutes and seconds
    Return minutes.ToString("00", Globalization.CultureInfo.InvariantCulture) & ":" _
             & seconds.ToString("00", Globalization.CultureInfo.InvariantCulture) & "s"
  End Function

  Public Function ShowWarning(warning As String) As DialogResult
    Return MessageBox.Show(warning, "Adaptive", MessageBoxButtons.OK, MessageBoxIcon.Warning)
  End Function

  Public Function ShowInformation(information As String) As DialogResult
    Return MessageBox.Show(information, "Adaptive", MessageBoxButtons.OK, MessageBoxIcon.Information)
  End Function

  Public Function ShowQuestion(question As String) As DialogResult
    Return MessageBox.Show(question, "Adaptive", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
  End Function

#If 0 Then
  'TODO

  'Added 2009/06/15 to use with Parameters_SevenDaySchedule to disregard Load Delays upon sunday night startup
  ' requested by allen-michael due to delay compensation for initial week startup delays
  ' a&e schedules a machine unblocked at startup, and the operator may take as long as 120minutes to actually load and complete
  Public Function DOW(ByVal GregDate As Date) As String
    ' Return values:
    ' 0 = Sunday
    ' 1 = Monday
    ' 2 = Tuesday
    ' 3 = Wednesday
    ' 4 = Thursday
    ' 5 = Friday
    ' 6 = Saturday
    Dim y As Integer
    Dim m As Integer
    Dim d As Integer

    ' monthdays:
    ' This is a "template" for a year. Each number
    ' stands for a day of the week. The general idea
    ' is that, in a standard year, if Jan 1 is on a
    ' Friday, then Feb 1 will be a Monday, Mar 1
    ' will be a Monday, April 1 will be a Thursday,
    ' May 1 will be Saturday, etc..
    Dim mcode As String
    Dim monthdays() As String
    monthdays = Split("5 1 1 4 6 2 4 0 3 5 1 3")

    ' Grab our date info
    y = val(Format(GregDate, "yyyy"))
    m = val(Format(GregDate, "mm"))
    d = val(Format(GregDate, "dd"))

    ' Snatch the corresponding month code
    mcode = val(monthdays(m - 1))

    ' Multiplying by 1.25 takes care of leap years,
    ' but not completely. Jan and Feb of a leap year
    ' will end up a day extra.
    ' The 'mod 7' gives us our day.
    DOW = ((Int(y * 1.25) + mcode + d) Mod 7)

    ' This takes care of leap year Jan and Feb days.
    If y Mod 4 = 0 And m < 3 Then DOW = (DOW + 6) Mod 7

  End Function




#End If
#If 0 Then
''' <summary>Return dispense list as a delimited string so the dispense list view in the mimic will work in PE / history.</summary>
  ''' <remarks>[2015-07-29] Sage Request all units to be in Kg</remarks>
  ''' <returns>String: code|name|Kilograms|DispenseKilograms|tolerancegrams|dispenseresult|dispenseerror|dispensestatus</returns>
  Public Function GetDispenseString(Job As DispenseJob) As String
    Dim returnString As String = Nothing
    If Job IsNot Nothing Then
      For Each item As DispenseItem In Job.Dispenses
        Dim dispenseResult As Integer = item.DispenseResult
        Dim dispenseStatus As Integer = item.DispenseStatus


        Dim newLine As String = item.Product.Code & "|" & item.Product.Name & "|" & item.Kilograms.ToString("#0.00") & "|" & _
                                item.DispenseKilograms.ToString("#0.00") & "|" & item.ToleranceKilogramsUnder.ToString("#0.00") & "|" & _
                                dispenseResult.ToString("#0") & "|" & item.DispenseError.ToString("#0.00") & "|" & _
                                dispenseStatus.ToString("#0")
        If returnString Is Nothing Then
          returnString = newLine
        Else
          returnString &= ";" & newLine
        End If
      Next
      'Add fill volume at the end
      returnString &= ";" & "Water|" & Job.FillLitersTarget.ToString("#0") & "|" & Job.FillLitersActual.ToString("#0")
    End If
    If returnString IsNot Nothing Then Return returnString
    Return ""
  End Function

#End If

  Public Function GetDispenseInformation(ByVal dispenseinformation As String, ByVal calloff As Integer) As String
    Try
      'These separators are used by programs and prefixed steps
      Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
      Dim separator1 As String = Convert.ToChar(255)
      Dim separator2 As String = ","

      'strings
      Dim steps() As String
      Dim StepDetailSplit() As String
      Dim stepnumbersplit() As String
      Dim NumberOfPrepsFound As Integer = 0
      Dim stepnumber As String = ""
      Dim scaleInfo As String = ""
      'Split the program string into an array of program steps - one step per array element
      steps = dispenseinformation.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
      'Make sure we've have something to check
      If steps.GetUpperBound(0) <= 0 Then Return ""

      For i = 1 To steps.GetUpperBound(0)
        StepDetailSplit = steps(i).Split(CChar(separator2))
        If StepDetailSplit(1) Like "Fnc=Prep*" Then
          NumberOfPrepsFound += 1
          If NumberOfPrepsFound = calloff Then
            stepnumbersplit = StepDetailSplit(0).Split(CChar("="))
            stepnumber = stepnumbersplit(1)
            scaleInfo = StepDetailSplit(3).Substring(StepDetailSplit(3).Length - 5, 5)
            Return stepnumber & "," & scaleInfo
          End If
        End If
      Next

      Return ""
    Catch ex As Exception
      Return ""
    End Try

  End Function

#If 0 Then

  Public Function MinMax(value As Short, min As Short, max As Short) As Short
    If value < min Then Return min
    If value > max Then Return max
    Return value
  End Function

  Public Function MinMax(value As Integer, min As Integer, max As Integer) As Integer
    If value < min Then Return min
    If value > max Then Return max
    Return value
  End Function

  Public Function MinMax(value As Double, min As Double, max As Double) As Double
    If value < min Then Return min
    If value > max Then Return max
    Return value
  End Function


#End If
  Public Function DoublePrecision(value As Double, decimalPlaces As Integer) As Double
    Dim formatString As String
    If decimalPlaces <= 0 Then
      formatString = "#0"
    Else
      formatString = "#0.00000000".Substring(0, decimalPlaces + 3)
    End If
    Return Double.Parse(value.ToString(formatString))
  End Function

  Public Function DisplayAmount(amount As Double, units As String) As String
    Return Amount.ToString & " " & units
  End Function

  Public Function DisplayAmount(amount As Double, units As String, decimalPlaces As Integer) As String
    Dim formatString As String
    If decimalPlaces <= 0 Then
      formatString = "#0"
    Else
      formatString = "#0.00000000".Substring(0, decimalPlaces + 3) ' no more than 8 decimal places
    End If
    Return Amount.ToString(formatString) & " " & units
  End Function

  Public Function SameJob(dyelot1 As String, reDye1 As Integer, dyelot2 As String, reDye2 As Integer) As Boolean
    Return (dyelot1 = dyelot2) AndAlso (reDye1 = reDye2)
  End Function

#If 0 Then

  Public Function MulDiv(ByVal value As Integer, ByVal multiply As Integer, ByVal divide As Integer) As Integer
    If divide = 0 Then Return 0 ' no divide by zero error please
    Return CType((CType(value, Long) * multiply) \ divide, Integer)
  End Function
  Public Function MulDiv(ByVal value As Short, ByVal multiply As Integer, ByVal divide As Integer) As Short
    If divide = 0 Then Return 0 ' no divide by zero error please
    Return CType((CType(value, Long) * multiply) \ divide, Short)
  End Function

  ''' <summary>Returns a rescaled value.</summary>
  ''' <param name="value"></param>
  ''' <param name="inMin"></param>
  ''' <param name="inMax"></param>
  ''' <param name="outMin"></param>
  ''' <param name="outMax"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ReScale(ByVal value As Integer, ByVal inMin As Integer, ByVal inMax As Integer, ByVal outMin As Integer, ByVal outMax As Integer) As Integer
    If inMin = inMax Then Return 0 ' avoid division by zero
    If value < inMin Then value = inMin
    If value > inMax Then value = inMax
    Return MulDiv(value - inMin, outMax - outMin, inMax - inMin) + outMin
  End Function

  Public Function ReScale(ByVal value As Short, ByVal inMin As Integer, ByVal inMax As Integer, ByVal outMin As Integer, ByVal outMax As Integer) As Short
    Return CType(ReScale(CType(value, Integer), inMin, inMax, outMin, outMax), Short)
  End Function

#End If
  Public Function Reverse(ByVal value As Short) As Short
    Return CType(-1 * (value - 1000), Short)
  End Function

  ''' <summary>Gets the percent (actually in tenths of a percent, so a number between 0 and 1000),
  ''' that the given value represents between the given minimum and maximum bounds.</summary>
  Public Function GetPercent(ByVal value As Integer, ByVal minValue As Integer, ByVal maxValue As Integer) As Integer
    Dim range = maxValue - minValue
    If range = 0 Then Return -1 ' avoid a division by zero error, and return a special value
    Dim ret = ((value - minValue) * 1000) \ range
    Return Math.Min(Math.Max(ret, 0), 1000)
  End Function

  Public Function WordWrap(ByVal strText As String, ByVal iWidth As Integer) As String

    If strText.Length <= iWidth Then Return strText ' dont need to do anything 

    Dim sResult As String = strText
    Dim sChar As String             ' temp holder for current string char 
    Dim iEn As Integer
    Dim iLineNO As Integer = iWidth
    Do While sResult.Length >= iLineNO
      For iEn = iLineNO To 1 Step -1         ' work backwards from the max len to 1 looking for a space 
        sChar = sResult.Chars(iEn)
        If sChar = " " Then             ' found a space 
          sResult = sResult.Remove(iEn, 1)     ' Remove the space 
          sResult = sResult.Insert(iEn, Environment.NewLine)     ' insert a line feed here, 
          iLineNO += iWidth             ' increment 
          Exit For
        End If
      Next
    Loop
    Return sResult
  End Function

  Public Function WordWrap(ByVal strText As String, ByVal iWidth As Integer, ByVal strSpecifier As String) As String

    If strText.Length <= iWidth Then Return strText ' dont need to do anything 

    Dim sResult As String = strText
    Dim sChar As String             ' temp holder for current string char 
    Dim iEn As Integer
    Dim iLineNO As Integer = iWidth
    Do While sResult.Length >= iLineNO
      For iEn = iLineNO To 1 Step -1         ' work backwards from the max len to 1 looking for a space or specifier (,) 
        sChar = sResult.Chars(iEn)
        If (sChar = " ") OrElse (sChar = strSpecifier) Then      ' found a space or specifier
          sResult = sResult.Remove(iEn, 1)                       ' Remove the space or specifier
          sResult = sResult.Insert(iEn, Environment.NewLine)     ' insert a line feed here, 
          iLineNO += iWidth                                      ' increment 
          Exit For

        End If
      Next
    Loop
    Return sResult
  End Function

  Public Function FindGreater(ByVal value1 As Integer, ByVal value2 As Integer) As Integer
    If value1 > value2 Then
      Return value1
    Else
      Return value2
    End If
  End Function

  Public Function FindLesser(ByVal value1 As Integer, ByVal value2 As Integer) As Integer
    If value1 < value2 Then
      Return value1
    Else
      Return value2
    End If
  End Function

  Public Function Summer(ByVal value1 As Integer, ByVal value2 As Integer) As Integer
    Return value1 + value2
  End Function

  Public Function BcdToDecimal(valueBCD As Integer) As Integer
    Dim ret As Integer
    Dim mult = 1

    Do While valueBCD <> 0
      ret = ret + (valueBCD Mod 16) * mult
      mult = mult * 10
      valueBCD = valueBCD \ 16
    Loop
    Return ret
  End Function


End Module