Public Class FormPassword
  Public ControlCode As ControlCode

  Public Shadows Event KeyPress(ByVal Key As String)

  Public ReadOnly Property Value() As String
    Get
      Value = TextBoxValue.Text
    End Get
  End Property

  Private password_ As Boolean
  Public Property Password() As Boolean
    Get
      Return password_
    End Get
    Set(ByVal value As Boolean)
      password_ = value
    End Set
  End Property

  Public Sub Clear()
    TextBoxValue.Text = ""
  End Sub
  Public Sub ExitForm()
    Visible = False
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
  Private Sub ButtonCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCancel.Click
    If String.IsNullOrEmpty(TextBoxValue.Text) Then
      ExitForm()
    Else
      Clear()
    End If
    RaiseEvent KeyPress("Cancel")
  End Sub

  Private Sub ButtonAccept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonAccept.Click
    If CheckPassword() Then
      Password = True
      Me.Close()
    End If
  End Sub

  Private Function CheckPassword() As Boolean
    Try
      If String.IsNullOrEmpty(Value) Then Return Message("Please enter password.")
      If Value <> "7295" Then Return Message("Incorrect Password, Retry or cancel")
      Return True
    Catch ex As Exception
      'some code
    End Try
    Return False
  End Function

  Private Function Message(ByVal text As String) As Boolean
    MessageBox.Show(text.PadRight(64), "Adaptive Control", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
    Return False
  End Function

End Class