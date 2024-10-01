Public Class UserControlKeypad

  Public Shadows Event KeyPress(ByVal Key As String)

  Public ReadOnly Property Value() As String
    Get
      Value = TextBoxValue.Text
    End Get
  End Property

  Public Sub Clear()
    TextBoxValue.Text = ""
  End Sub

  Private Sub Button0_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button0.Click
    TextBoxValue.Text = TextBoxValue.Text & "0"
    RaiseEvent KeyPress("0")
  End Sub
  Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
    TextBoxValue.Text = TextBoxValue.Text & "1"
    RaiseEvent KeyPress("1")
  End Sub
  Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
    TextBoxValue.Text = TextBoxValue.Text & "2"
    RaiseEvent KeyPress("2")
  End Sub
  Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
    TextBoxValue.Text = TextBoxValue.Text & "3"
    RaiseEvent KeyPress("3")
  End Sub
  Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
    TextBoxValue.Text = TextBoxValue.Text & "4"
    RaiseEvent KeyPress("4")
  End Sub
  Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
    TextBoxValue.Text = TextBoxValue.Text & "5"
    RaiseEvent KeyPress("5")
  End Sub
  Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
    TextBoxValue.Text = TextBoxValue.Text & "6"
    RaiseEvent KeyPress("6")
  End Sub
  Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
    TextBoxValue.Text = TextBoxValue.Text & "7"
    RaiseEvent KeyPress("7")
  End Sub
  Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
    TextBoxValue.Text = TextBoxValue.Text & "8"
    RaiseEvent KeyPress("8")
  End Sub
  Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
    TextBoxValue.Text = TextBoxValue.Text & "9"
    RaiseEvent KeyPress("9")
  End Sub
  Private Sub ButtonClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonClear.Click
    Clear()
    RaiseEvent KeyPress("Clear")
  End Sub
  Private Sub ButtonDecimal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonDecimal.Click
    TextBoxValue.Text = TextBoxValue.Text & "."
    RaiseEvent KeyPress(".")
  End Sub

End Class