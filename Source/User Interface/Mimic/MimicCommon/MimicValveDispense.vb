<DefaultBindingPropertyAttribute("Value")> _
Public Class MimicValveDispense : Inherits System.Windows.Forms.UserControl

  Private imageHorizontalOff As Drawing.Image = My.Resources.ValveHorizontalOff
  Private imageHorizontalOn As Drawing.Image = My.Resources.ValveHorizontalOn
  Private imageHorizontalLowFlow As Drawing.Image = My.Resources.ValveHorizontalPartial

  Private imageVerticalOff As Drawing.Image = My.Resources.ValveVerticalOff
  Private imageVerticalOn As Drawing.Image = My.Resources.ValveVerticalOn
  Private imageVerticalLowFlow As Drawing.Image = My.Resources.ValveVerticalPartial

  ' These two properties let us hook into the BatchControl mimic overrides
  '  
  Private _Value As Boolean
  Public Property Value() As Boolean
    Get
      Return _Value
    End Get
    Set(value As Boolean)
      If _Value <> value Then
        _Value = value
        Redraw()
      End If
    End Set
  End Property

  ' Set to true during low flow so we can use a different image
  Public Property LowFlow As Boolean

  Private _UIEnabled As Boolean
  <DefaultValue(True)> _
  Public Property UIEnabled() As Boolean
    Get
      Return _UIEnabled
    End Get
    Set(ByVal value As Boolean)
      _UIEnabled = value
    End Set
  End Property

  Public Enum EOrientation
    Horizontal
    Vertical
  End Enum

  Private _Orientation As EOrientation
  Public Property Orientation As EOrientation
    Get
      Return _Orientation
    End Get
    Set(value As EOrientation)
      'See if the orientation has changed
      If _Orientation <> value Then
        _Orientation = value
        Redraw()
      End If
    End Set
  End Property

  Public ReadOnly Property Image() As Drawing.Image
    Get
      If Orientation = EOrientation.Horizontal Then
        Return HorizontalImage
      Else
        Return VerticalImage
      End If
    End Get
  End Property

  Private ReadOnly Property HorizontalImage As Drawing.Image
    Get
      If Value Then
        If LowFlow Then
          Return My.Resources.ValveHorizontalPartial
        Else
          Return My.Resources.ValveHorizontalOn
        End If
      Else
        Return My.Resources.ValveHorizontalOff
      End If
    End Get
  End Property

  Private ReadOnly Property VerticalImage As Drawing.Image
    Get
      If Value Then
        If LowFlow Then
          Return My.Resources.ValveVerticalPartial
        Else
          Return My.Resources.ValveVerticalOn
        End If
      Else
        Return My.Resources.ValveVerticalOff
      End If
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
    Using g As Graphics = e.Graphics
      g.DrawImage(Image, 0, 0, Image.Width, Image.Height)
    End Using
  End Sub

  Protected Overrides Sub OnResize(ByVal e As EventArgs)
    Redraw()
  End Sub

  Private Sub Redraw()
    ' Set size and force a repaint
    Me.Size = Image.Size
    Me.Invalidate()
  End Sub

  Public Sub New()
    Me.Name = "Valve"
    Me.Orientation = EOrientation.Horizontal
    Me.Size = New System.Drawing.Size(25, 13)

    Redraw()
  End Sub

End Class
