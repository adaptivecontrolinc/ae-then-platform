Public Class ControlButtonClose : Inherits Control
  Public Event Close As EventHandler

  Private Shared ReadOnly bmClose_ As Bitmap
  Shared Sub New()
    bmClose_ = New Bitmap(My.Resources.OverviewContainerClose)
    bmClose_.MakeTransparent(Color.Magenta)
  End Sub

  Private lastMouseOverClose_ As Boolean

  Protected Overrides ReadOnly Property DefaultSize() As Size
    Get
      Return New Size(bmClose_.Width \ 2, bmClose_.Height)
    End Get
  End Property

  Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
    If e.ClipRectangle.IsEmpty Then Exit Sub
    Dim xSource As Integer : If MouseOverClose Then xSource = bmClose_.Width \ 2
    e.Graphics.DrawImage(bmClose_, 0, 0, New Rectangle(xSource, 0, bmClose_.Width \ 2, bmClose_.Height), GraphicsUnit.Pixel)
  End Sub

  Protected Overrides Sub OnMouseLeave(ByVal e As EventArgs)
    If lastMouseOverClose_ Then lastMouseOverClose_ = False : Invalidate()
  End Sub

  Private ReadOnly Property MouseOverClose() As Boolean
    Get
      Return ClientRectangle.Contains(PointToClient(MousePosition))
    End Get
  End Property

  Protected Overrides Sub OnMouseMove(ByVal e As MouseEventArgs)
    Dim newMouseOverClose = MouseOverClose
    If lastMouseOverClose_ = newMouseOverClose Then Exit Sub
    lastMouseOverClose_ = newMouseOverClose
    Invalidate()
  End Sub

  Protected Overrides Sub OnClick(ByVal e As EventArgs)
    OnClose(EventArgs.Empty)
  End Sub
  Protected Overridable Sub OnClose(ByVal e As EventArgs)
    RaiseEvent Close(Me, e)
  End Sub
End Class