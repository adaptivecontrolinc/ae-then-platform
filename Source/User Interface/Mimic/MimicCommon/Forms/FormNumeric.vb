Imports Settings
Imports Utilities.Translations

Public Class FormNumeric

  Public Sub New()

    ' This call is required by the designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    Me.ActiveControl = textBoxValue
  End Sub

  Public Property Amount As String
    Get
      Return textBoxValue.Text
    End Get
    Set(ByVal value As String)
      textBoxValue.Text = value
    End Set
  End Property

  Public Property AmountInteger As Integer
    Get
      Dim tryInteger As Integer
      If Integer.TryParse(textBoxValue.Text, tryInteger) Then Return tryInteger
      Return 0
    End Get
    Set(ByVal value As Integer)
      textBoxValue.Text = value.ToString
    End Set
  End Property

  Public Property AmountDouble As Double
    Get
      Dim tryDouble As Double
      If Double.TryParse(textBoxValue.Text, tryDouble) Then Return tryDouble
      Return 0
    End Get
    Set(ByVal value As Double)
      textBoxValue.Text = value.ToString
    End Set
  End Property

  Private Sub keysNumericMain_KeyClick(ByVal text As String) Handles keysNumericMain.KeyClick
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