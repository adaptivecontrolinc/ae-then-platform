Public Class Flasher : Inherits MarshalByRefObject

  Private startTickCount As UInt32

  Property TimeOn As Integer       ' In milliseconds
  Property TimeOff As Integer
  Property Period As Integer

  Sub New()
    NewBase(1000, 1000)
  End Sub
  Sub New(timeOn As Integer)
    NewBase(timeOn, timeOn)
  End Sub
  Sub New(timeOn As Integer, timeOff As Integer)
    NewBase(timeOn, timeOff)
  End Sub
  Private Sub NewBase(timeOn As Integer, timeOff As Integer)
    Me.TimeOn = timeOn
    Me.TimeOff = timeOff
    Me.Period = timeOn + timeOff
    startTickCount = TickCount
  End Sub

  ReadOnly Property [On] As Boolean
    Get
      Dim elapsed = TickCount - startTickCount
      Dim x As UInt32 = elapsed Mod CType(Period, UInt32)
      Return (x < TimeOn)
    End Get
  End Property

End Class
