Imports Settings
Imports Utilities.Translations

Public Class FormNumberPad

  Public Sub New()

    ' This call is required by the designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    Me.ActiveControl = controlNumberPad
  End Sub

  Public Property Title As String
    Get
      Return Me.Text
    End Get
    Set(value As String)
      Me.Text = value
    End Set
  End Property

  Public Property Amount As String
    Get
      Return controlNumberPad.Amount
    End Get
    Set(ByVal value As String)
      controlNumberPad.Amount = value
    End Set
  End Property

  Public Property AmountInteger As Integer
    Get
      Return controlNumberPad.AmountInteger
    End Get
    Set(ByVal value As Integer)
      controlNumberPad.AmountInteger = value
    End Set
  End Property

  Public Property AmountDouble As Double
    Get
      Return controlNumberPad.AmountDouble
    End Get
    Set(ByVal value As Double)
      controlNumberPad.AmountDouble = value
    End Set
  End Property

  Private Sub keysNumericMain_KeyClick(ByVal text As String)
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