Module Support



  ''' <summary>Return the default value if the parameter value = 0.</summary>
  ''' <param name="parameterValue">The parameter value to check.</param>
  ''' <param name="defaultValue">The default value to return if the parameter value = 0.</param>
  Public Function DefaultParameter(parameterValue As Integer, defaultValue As Integer) As Integer
    If parameterValue <= 0 Then Return defaultValue
    Return parameterValue
  End Function


  ''' <summary>Return default value if value is empty or null.</summary>
  ''' <param name="value">The value to check.</param>
  ''' <param name="defaultValue">The value to return if value is empty or null.</param>
  Public Function DefaultSetting(value As String, defaultValue As String) As String
    If String.IsNullOrEmpty(value) Then Return defaultValue
    Return value
  End Function

  ''' <summary>Return default value if value cannot be converted to an integer.</summary>
  ''' <param name="value">The value to check.</param>
  ''' <param name="defaultValue">The value to return if value is empty or null.</param>
  Public Function DefaultSetting(value As String, defaultValue As Integer) As Integer
    Dim tryInteger As Integer
    If Integer.TryParse(value, tryInteger) Then Return tryInteger
    Return defaultValue
  End Function


  Public Function MinMax(ByVal value As Integer, ByVal min As Integer, ByVal max As Integer) As Integer
    If value < min Then Return min
    If value > max Then Return max
    Return value
  End Function

  Public Function MinMax(ByVal value As Double, ByVal min As Double, ByVal max As Double) As Double
    If value < min Then Return min
    If value > max Then Return max
    Return value
  End Function

  Public Function MinMaxAbs(ByVal value As Integer, ByVal min As Integer, ByVal max As Integer) As Integer
    Dim testValue = MinMax(Math.Abs(value), min, max)
    If value < 0 Then Return -testValue
    Return testValue
  End Function

  Public Function MinMaxAbs(ByVal value As Double, ByVal min As Double, ByVal max As Double) As Double
    Dim testValue = MinMax(Math.Abs(value), min, max)
    If value < 0 Then Return -testValue
    Return testValue
  End Function

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

End Module

Public Module TickCountModule
  Friend TickCount As UInt32
End Module

' -------------------------------------------------------------------------
''' <summary>A class that raises an event if there has not been mouse or keyboard activity for this application for a while.</summary>
Public Class InactiveTimeout : Implements System.Windows.Forms.IMessageFilter, IDisposable
  Public Event Timeout As EventHandler
  Private reLoad_, remaining_ As Integer
  Private timer_ As New Threading.Timer(AddressOf OnTimer, Nothing, 1000, 1000)

  Public Sub New()
    System.Windows.Forms.Application.AddMessageFilter(Me)
  End Sub

  Public Sub Dispose() Implements IDisposable.Dispose
    If timer_ IsNot Nothing Then
      timer_.Dispose() : timer_ = Nothing
      System.Windows.Forms.Application.RemoveMessageFilter(Me)
    End If
  End Sub

  Public Sub SetTimeout(ByVal seconds As Integer)
    reLoad_ = seconds : remaining_ = seconds
  End Sub

  Private Sub OnTimer(ByVal state As Object)
    Static inside_ As Boolean : If inside_ Then Exit Sub
    inside_ = True
    If remaining_ > 0 Then
      remaining_ -= 1
      If remaining_ = 0 Then RaiseEvent Timeout(Me, EventArgs.Empty)
    End If
    inside_ = False
  End Sub

  Private Function PreFilterMessage(ByRef m As System.Windows.Forms.Message) As Boolean Implements System.Windows.Forms.IMessageFilter.PreFilterMessage
    Select Case m.Msg
      Case &H201 To &H20D, &HA0 To &HAD, &H100 To &H109  ' all mouse and keyboard messages (except WM_MOUSEMOVE)
        If remaining_ > 0 Then remaining_ = reLoad_
    End Select
    Return False
  End Function
End Class

