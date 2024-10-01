Imports Settings
Imports Utilities.Translations

Public Class FormQwerty

  Public Sub New()

    ' This call is required by the designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    Me.ActiveControl = textBoxValue
  End Sub

  Public Property Value As String
    Get
      Return textBoxValue.Text
    End Get
    Set(ByVal value As String)
      textBoxValue.Text = value
    End Set
  End Property

  Private Sub keyboardMain_KeyClick(ByVal text As String) Handles keyboardMain.KeyClick
    SendKeys.Send(text)
  End Sub

  Private Sub buttonOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonOK.Click
    Me.DialogResult = Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub

  Private Sub buttonCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonCancel.Click
    Me.DialogResult = Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub

End Class