<DefaultBindingPropertyAttribute("Value")> _
Public MustInherit Class MimicDevice : Inherits Windows.Forms.Control

  <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
  Public Property ImageHorizontalOff As Drawing.Image
  <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
  Public Property ImageHorizontalOn As Drawing.Image
  <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
  Public Property ImageHorizontalPartial As Drawing.Image
  <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
  Public Property ImageHorizontalFault As Drawing.Image

  <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
  Public Property ImageVerticalOff As Drawing.Image
  <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
  Public Property ImageVerticalOn As Drawing.Image
  <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
  Public Property ImageVerticalPartial As Drawing.Image
  <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
  Public Property ImageVerticalFault As Drawing.Image

  ' These two properties let us hook into the automatic BatchControl mimic handling
  '  
  Private _Value As Boolean
  <Category("Data"), DefaultValue(False), Bindable(True)> _
  Public Property Value() As Boolean
    Get
      Return _Value
    End Get
    Set(value As Boolean)
      If _Value <> value Then
        _Value = value
        Invalidate()
      End If
    End Set
  End Property

  Private _UIEnabled As Boolean
  <Category("Behavior"), DefaultValue(True)> _
  Public Property UIEnabled() As Boolean
    Get
      Return _UIEnabled
    End Get
    Set(ByVal value As Boolean)
      _UIEnabled = value
    End Set
  End Property

  ' Set this property to flag the dout has been overridden
  Private _Overridden As Boolean
  <Category("Data")> _
  Public Property Overridden As Boolean
    Get
      Return _Overridden
    End Get
    Set(value As Boolean)
      If _Overridden <> value Then
        _Overridden = value
        Invalidate()
      End If
    End Set
  End Property

  ' Set this property to reverse logic - for a normally open valve for instance
  Private _NormallyOn As Boolean
  <Category("Data")> _
  Public Property NormallyOn As Boolean
    Get
      Return _NormallyOn
    End Get
    Set(value As Boolean)
      If _NormallyOn <> value Then
        _NormallyOn = value
        Invalidate()
      End If
    End Set
  End Property

  <Category("Data")> _
  Public ReadOnly Property DeviceOn As Boolean
    Get
      If NormallyOn Then Return Not Value
      Return Value
    End Get
  End Property

  <Category("Data")> _
  Public ReadOnly Property DeviceOff As Boolean
    Get
      If NormallyOn Then Return Value
      Return Not Value
    End Get
  End Property

  Public Enum EOrientation
    Horizontal
    Vertical
  End Enum

  Private _Orientation As EOrientation
  <Category("Behavior")> _
  Public Property Orientation As EOrientation
    Get
      Return _Orientation
    End Get
    Set(value As EOrientation)
      'See if the orientation has changed
      If _Orientation <> value Then
        _Orientation = value
        Invalidate()
      End If
    End Set
  End Property

  <Category("Data")> _
  Public ReadOnly Property Fault As Boolean
    Get
      'Check valve open feedback
      If OnFeedbackEnabled AndAlso (DeviceOn And Not OnFeedback) Then Return True

      'Check valve closed feedback
      If OffFeedbackEnabled AndAlso (DeviceOff And Not OffFeedback) Then Return True

      Return False
    End Get
  End Property

  <Category("Behavior")> _
  Public Property OnFeedbackEnabled As Boolean
  <Category("Behavior")> _
  Public Property OffFeedbackEnabled As Boolean

  Private _OnFeedback As Boolean
  <Category("Data")> _
  Public Property OnFeedback As Boolean
    Get
      Return _OnFeedback
    End Get
    Set(value As Boolean)
      'If this value is set automatically set feedback enabled flag
      OnFeedbackEnabled = True
      If _OnFeedback <> value Then
        _OnFeedback = value
        Invalidate()
      End If
    End Set
  End Property

  Private _OffFeedback As Boolean
  <Category("Data")> _
  Public Property OffFeedback As Boolean
    Get
      Return _OffFeedback
    End Get
    Set(value As Boolean)
      'If this value is set automatically set feedback enabled flag
      OffFeedbackEnabled = True
      If _OffFeedback <> value Then
        _OffFeedback = value
        Invalidate()
      End If
    End Set
  End Property

  Private ReadOnly Property GetImage() As Drawing.Image
    Get
      If Orientation = EOrientation.Horizontal Then
        Return GetHorizontalImage
      Else
        Return GetVerticalImage
      End If
    End Get
  End Property

  Private ReadOnly Property GetHorizontalImage As Drawing.Image
    Get
      If Fault Then Return ImageHorizontalFault
      If DeviceOn Then Return ImageHorizontalOn
      Return ImageHorizontalOff
    End Get
  End Property

  Private ReadOnly Property GetVerticalImage As Drawing.Image
    Get
      If Fault Then Return ImageVerticalFault
      If DeviceOn Then Return ImageVerticalOn
      Return ImageVerticalOff
    End Get
  End Property

  Protected Overrides ReadOnly Property CreateParams() As System.Windows.Forms.CreateParams
    Get
      Dim cp As CreateParams = MyBase.CreateParams
      cp.ExStyle = &H20 ' WS_EX_TRANSPARENT
      Return cp
    End Get
  End Property

  Protected Overrides Sub OnMouseDown(ByVal e As MouseEventArgs)
    ' Flip the value using the mouse left click button
    If e.Button = Windows.Forms.MouseButtons.Left Then
      If UIEnabled Then Value = Not Value
    End If
  End Sub

  Protected Overrides Sub OnPaintBackground(ByVal e As System.Windows.Forms.PaintEventArgs)
    'Stop background paint
  End Sub

  Protected Overrides Sub OnResize(ByVal e As EventArgs)
    ' Fix size to image size
    Me.Size = GetImage.Size
  End Sub

  Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
    MyBase.OnPaint(e)
    Using g = e.Graphics
      'g.Clear(Me.BackColor)
      g.DrawImage(GetImage, 0, 0, GetImage.Width, GetImage.Height)
      If Overridden Then g.DrawRectangle(Pens.Blue, 0, 0, Me.Width - 1, Me.Height - 1)
    End Using
  End Sub

End Class
