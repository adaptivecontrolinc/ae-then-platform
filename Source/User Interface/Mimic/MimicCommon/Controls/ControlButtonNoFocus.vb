Public Class ControlButtonNoFocus : Inherits Windows.Forms.Control

  Public Event KeyClick(ByVal sender As Object, ByVal text As String)

  Private pressed_ As Boolean
  Public Property Pressed() As Boolean
    Get
      Return pressed_
    End Get
    Set(ByVal value As Boolean)
      pressed_ = value
    End Set
  End Property

  Protected Overrides Sub OnClick(ByVal e As System.EventArgs)
    RaiseEvent KeyClick(Me, Me.Text)
  End Sub

  Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
    Me.Pressed = True
    Me.Invalidate()
  End Sub

  Protected Overrides Sub OnMouseUp(ByVal e As System.Windows.Forms.MouseEventArgs)
    Me.Pressed = False
    Me.Invalidate()
  End Sub

  Protected Overrides Sub OnTextChanged(ByVal e As System.EventArgs)
    Me.Invalidate()
  End Sub

  Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
    If Me.Pressed Then
      DrawButtonPressed(e.Graphics)
    Else
      DrawButton(e.Graphics)
    End If
  End Sub

  Private Sub DrawButton(ByVal g As System.Drawing.Graphics)
    ' Draw a rounded border around the outside of the control to give a button effect
    Dim rOuter As New Rectangle(0, 0, Me.ClientSize.Width - 1, Me.ClientSize.Height - 1)
    Dim rInner As New Rectangle(1, 1, Me.ClientSize.Width - 3, Me.ClientSize.Height - 3)
    'Utilities.Draw.DrawRoundedRectangle(g, rInner, New Pen(Color.Gray))
    Utilities.Draw.DrawRoundedRectangle(g, rOuter, New Pen(Color.Silver))

    ' Draw the text if there is any.
    If Me.Text.Length > 0 Then
      Dim size As SizeF = g.MeasureString(Me.Text, Me.Font)
      ' Center the text inside the client area of the control
      g.DrawString(Me.Text, Me.Font, New SolidBrush(Me.ForeColor), (Me.ClientSize.Width - size.Width) / 2, (Me.ClientSize.Height - size.Height) / 2)
    End If
  End Sub

  Private Sub DrawButtonPressed(ByVal g As System.Drawing.Graphics)

    ' Draw a rounded border around the outside of the control to give a button effect
    Dim rOuter As New Rectangle(0, 0, Me.ClientSize.Width - 1, Me.ClientSize.Height - 1)
    Dim rInner As New Rectangle(1, 1, Me.ClientSize.Width - 3, Me.ClientSize.Height - 3)
    Utilities.Draw.DrawFilledRoundedRectangle(g, rOuter, New Pen(Color.Silver))

    ' Draw the text if there is any.
    If Me.Text.Length > 0 Then
      Dim size As SizeF = g.MeasureString(Me.Text, Me.Font)
      ' Center the text inside the client area of the control
      g.DrawString(Me.Text, Me.Font, New SolidBrush(Color.White), (Me.ClientSize.Width - size.Width) / 2, (Me.ClientSize.Height - size.Height) / 2)
    End If
  End Sub
End Class
