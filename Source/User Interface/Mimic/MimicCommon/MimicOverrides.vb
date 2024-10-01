Public Class MimicOverrides

  Private Property controlCode As ControlCode
  Private Property controls As Windows.Forms.Control.ControlCollection

  Public Sub Update(controlCode As ControlCode, controls As Windows.Forms.Control.ControlCollection)
    Me.controlCode = controlCode
    Me.controls = controls
    UpdateForces()
  End Sub

  Private Sub UpdateForces()
    ' Make sure the control code reference is set
    If controlCode Is Nothing Then
      ResetOverrides()
      Exit Sub
    End If

    ' Get the forces string 
    Dim forcesString = GetForces()

    ' If no forces are set or we are not in override mode then clear any override flags
    If forcesString Is Nothing Then
      ResetOverrides()
      Exit Sub
    End If

    ' Set override property on each control 
    Dim device As MimicDevice
    For Each item As Control In controls
      If TypeOf item Is MimicDevice Then
        device = DirectCast(item, MimicDevice)
        device.Overridden = forcesString.Contains(device.Name.Substring(3))
      End If
    Next
  End Sub

  Private Sub ResetOverrides()
    ' Make sure the refernce is set
    If controls Is Nothing Then Exit Sub
    ' Set override property on each control to false 
    Dim device As MimicDevice
    For Each item As Control In controls
      If TypeOf item Is MimicDevice Then
        device = DirectCast(item, MimicDevice)
        device.Overridden = False
      End If
    Next
  End Sub

  Private Function GetForces() As String
    Try
      ' Get the forces string here so we can trap any errors - only check in override mode
      If controlCode.Parent.Mode = Mode.Override OrElse controlCode.Parent.Mode = Mode.Test Then
        Return controlCode.Parent.Forces
      Else
        Return Nothing
      End If
    Catch ex As Exception
      Utilities.Log.LogError(ex)
    End Try
    Return Nothing
  End Function

End Class
