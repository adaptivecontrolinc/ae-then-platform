Public Class FormGallons

  Public Property Title As String
    Get
      Return Me.Text
    End Get
    Set(value As String)
      Me.Text = value
    End Set
  End Property

  Public Property Label As String
    Get
      Return Me.labelGallons.Text
    End Get
    Set(value As String)
      Me.labelGallons.Text = value
    End Set
  End Property

  Public Property Gallons As Double
    Get
      Return Utilities.Sql.NullToZeroDouble(textBoxGallons.Text)
    End Get
    Set(value As Double)
      textBoxGallons.Text = value.ToString
    End Set
  End Property

  Private Sub buttonOK_Click(sender As System.Object, e As System.EventArgs) Handles buttonOK.Click
    Me.DialogResult = Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub

  Private Sub buttonCancel_Click(sender As System.Object, e As System.EventArgs) Handles buttonCancel.Click
    Me.DialogResult = Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub
End Class