Public Class SmoothingOld
  ' Use a circular buffer of values
  Private values_() As Integer, count_, firstOfs_ As Integer, sum_ As Long

  Public Function Smooth(ByVal value As Integer, ByVal smoothing As Integer) As Integer
    If smoothing < 2 Then Return value ' no need to smooth
    If values_ Is Nothing OrElse values_.Length <> smoothing Then
      values_ = New Integer(smoothing - 1) {}
      count_ = 0 : firstOfs_ = 0
    End If
    ' Keep a correct sum at all times for performance
    sum_ += value
    If count_ < smoothing Then
      values_(count_) = value : count_ += 1
    Else
      sum_ -= values_(firstOfs_) : values_(firstOfs_) = value
      firstOfs_ += 1 : If firstOfs_ = count_ Then firstOfs_ = 0
    End If
    Return CType(sum_ \ count_, Integer)
  End Function

  Public Function Smooth(ByVal value As Short, ByVal smoothing As Integer) As Short
    Return CType(Smooth(CType(value, Integer), smoothing), Short)
  End Function
End Class
