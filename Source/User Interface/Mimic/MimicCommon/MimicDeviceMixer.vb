Public Class MimicDeviceMixer : Inherits MimicDevice

  Public Sub New()

    ' Set images for a mixer
    ImageHorizontalOff = My.Resources.MixerMotorOff
    ImageHorizontalOn = My.Resources.MixerMotorOn
    ImageHorizontalFault = My.Resources.MixerMotorOn

    ImageVerticalOff = My.Resources.MixerMotorOff
    ImageVerticalOn = My.Resources.MixerMotorOn
    ImageVerticalFault = My.Resources.MixerMotorOn

    ' Set default values 
    Me.Orientation = EOrientation.Horizontal
    Me.Name = "Mixer"
    Me.Size = ImageHorizontalOff.Size

    ' Add a context menu for valve overrides
    ContextMenu = New ContextMenu
    ContextMenu.MenuItems.Add("Start Mixer", Sub() MenuStartMixer())
    ContextMenu.MenuItems.Add("Stop Mixer", Sub() MenuStopMixer())
    ContextMenu.MenuItems.Add("-")
  End Sub

  Private Sub MenuStartMixer()
    If NormallyOn Then
      Value = False
    Else
      Value = True
    End If
  End Sub

  Private Sub MenuStopMixer()
    If NormallyOn Then
      Value = True
    Else
      Value = False
    End If
  End Sub


End Class
