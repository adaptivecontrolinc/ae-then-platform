Public Class MimicDevicePump : Inherits MimicDevice

  Public Sub New()

    ' Set images for the pump
    ImageHorizontalOff = My.Resources.PumpOff
    ImageHorizontalOn = My.Resources.PumpOn
    ImageHorizontalFault = My.Resources.PumpOn

    ImageVerticalOff = My.Resources.PumpOff
    ImageVerticalOn = My.Resources.PumpOn
    ImageVerticalFault = My.Resources.PumpOn

    ' Set default values 
    Orientation = EOrientation.Horizontal
    Name = "Pump"
    Size = ImageHorizontalOff.Size

    ' Add a context menu for valve overrides
    ContextMenu = New ContextMenu
    ContextMenu.MenuItems.Add("Start Pump", Sub() MenuStartPump())
    ContextMenu.MenuItems.Add("Stop Pump", Sub() MenuStopPump())
    ContextMenu.MenuItems.Add("-")
  End Sub

  Private Sub MenuStartPump()
    If NormallyOn Then
      Value = False
    Else
      Value = True
    End If
  End Sub

  Private Sub MenuStopPump()
    If NormallyOn Then
      Value = True
    Else
      Value = False
    End If
  End Sub

End Class
