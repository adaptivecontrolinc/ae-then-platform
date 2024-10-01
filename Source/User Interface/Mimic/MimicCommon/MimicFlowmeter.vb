Public Class MimicFlowmeter

  Private imageSize As New Size(13, 25)
  Private imageOn As Drawing.Image = My.Resources.FlowmeterOn
  Private imageOff As Drawing.Image = My.Resources.FlowmeterOff

  Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    'MyBase.DoubleBuffered = True
    InitializeControl()
  End Sub

  Private Sub InitializeControl()
    Redraw()
  End Sub

  Private _On As Boolean
  Public Property [On] As Boolean
    Get
      Return _On
    End Get
    Set(value As Boolean)
      If _On <> value Then
        _On = value
        Redraw()
      End If
    End Set
  End Property

  Private ReadOnly Property Image As Drawing.Image
    Get
      'Return the appropriate image
      If Me.On Then Return imageOn
      Return imageOff
    End Get
  End Property

  Protected Overrides ReadOnly Property CreateParams() As System.Windows.Forms.CreateParams
    Get
      Dim cp As CreateParams = MyBase.CreateParams
      cp.ExStyle = &H20 ' WS_EX_TRANSPARENT
      Return cp
    End Get
  End Property

  Protected Overrides Sub OnPaintBackground(ByVal e As System.Windows.Forms.PaintEventArgs)
    'Stop background paint
    'MyBase.OnPaintBackground(e)
  End Sub

  Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
    MyBase.OnPaint(e)
    If Image IsNot Nothing Then
      Using g As Graphics = e.Graphics
        g.DrawImage(Image, 0, 0, Me.Width, Me.Height)
      End Using
    End If
  End Sub

  Protected Overrides Sub OnResize(ByVal e As EventArgs)
    Redraw()
  End Sub

  Private Sub Redraw()
    ' Set control size to image size if we have ome
    Me.Size = imageSize

    ' Invalidate the control to force a repaint.
    Me.Invalidate()
  End Sub

End Class
