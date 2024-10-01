Public Class TextBoxX : Inherits Windows.Forms.TextBox

  Private fixedHeight As Integer = 24

  Protected Overrides Sub OnSizeChanged(e As System.EventArgs)
    MyBase.OnSizeChanged(e)
  End Sub

  Protected Overrides Sub OnCreateControl()
    MyBase.OnCreateControl()
  End Sub

  Private Sub SetHeight()
    If Me.Height <> fixedHeight Then
      If Not Me.Multiline Then
        Me.Multiline = True
        Me.MinimumSize = New Size(0, fixedHeight)
        Me.Size = New Size(Me.Width, fixedHeight)
        Me.Multiline = False
      End If
    End If
  End Sub

#If 0 Then

  public class MyTextBox : TextBox
{
    const int RequestedHight = 30;

    protected override void OnSizeChanged(EventArgs e)
    {
        base.OnSizeChanged(e);
        AssureRequestedHight();
    }

    protected override void OnCreateControl()
    {
        base.OnCreateControl();
        AssureRequestedHight();
    }

    private void AssureRequestedHight()
    {
        if (this.Size.Height != RequestedHight && !this.Multiline) {
            this.Multiline = true;
            this.MinimumSize = new Size(0, RequestedHight);
            this.Size = new Size(this.Size.Width, RequestedHight);
            this.Multiline = false;
        }
    }
}


#End If


End Class
