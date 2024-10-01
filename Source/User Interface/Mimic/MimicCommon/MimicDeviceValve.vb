Public Class MimicDeviceValve : Inherits MimicDevice

  Public Sub New()

    ' These images were drawn horizontally
    ImageHorizontalOff = My.Resources.ValveOff
    ImageHorizontalOn = My.Resources.ValveOn
    ImageHorizontalFault = My.Resources.ValveError

    ' Rotate the images for vertical
    ImageVerticalOff = My.Resources.ValveOff : ImageVerticalOff.RotateFlip(RotateFlipType.Rotate90FlipNone)
    ImageVerticalOn = My.Resources.ValveOn : ImageVerticalOn.RotateFlip(RotateFlipType.Rotate90FlipNone)
    ImageVerticalFault = My.Resources.ValveError : ImageVerticalFault.RotateFlip(RotateFlipType.Rotate90FlipNone)

    'Set default values 
    Me.Orientation = EOrientation.Horizontal
    Me.Name = "Pump"
    Me.Size = ImageHorizontalOff.Size

    'Add a context menu for valve overrides
    Me.ContextMenu = New ContextMenu
    Me.ContextMenu.MenuItems.Add("Open Valve", Sub() MenuOpen())
    Me.ContextMenu.MenuItems.Add("Close Valve", Sub() MenuClose())
    Me.ContextMenu.MenuItems.Add("-")
    'Me.ContextMenu.MenuItems.Add("Pulse Valve", Sub() MenuPulse())
  End Sub

  Private Sub MenuOpen()
    If NormallyOn Then
      Value = False
    Else
      Value = True
    End If
  End Sub

  Private Sub MenuClose()
    If NormallyOn Then
      Value = True
    Else
      Value = False
    End If
  End Sub

  Private Sub MenuPulse()
    MessageBox.Show("MenuPulse")
  End Sub

End Class
