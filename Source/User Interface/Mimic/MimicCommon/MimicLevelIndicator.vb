' Basically a smooth vertical progress bar

Imports System.ComponentModel
Imports System.Drawing.Drawing2D

Public Class MimicLevelIndicator

  Public Sub New()
    ' This call is required by the designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    MyBase.DoubleBuffered = True
  End Sub

  Enum ColorStyle
    Solid
    Gradient
  End Enum

  Enum SigmaMode
    None
    SigmaBell
  End Enum

  Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
    'Dim rectf As RectangleF = RectangleF.op_Implicit(Me.ClientRectangle)

    Dim rectControl As RectangleF = Me.ClientRectangle
    Dim rectProgress As RectangleF = rectControl
    Dim brushGradient As LinearGradientBrush
    Dim brushSolid As SolidBrush

    ' Draw the background if gradient is selected
    If BackColorStyle = ColorStyle.Gradient Then
      brushGradient = New LinearGradientBrush(rectControl, BackColorMain, BackColorGradient, LinearGradientMode.Vertical)
      e.Graphics.FillRectangle(brushGradient, rectControl)
    End If

    ' Calculate area for drawing the progress bar
    rectProgress.Height = rectProgress.Height * Percent

    ' Set the position of the progress bar - it goes from the bottom up
    rectProgress.Y = rectControl.Y + Me.Height - rectProgress.Height

    ' Draw the progress bar
    If rectProgress.Width > 0 AndAlso rectProgress.Height > 0 Then

      If BarColorStyle = ColorStyle.Solid Then
        brushSolid = New SolidBrush(BarColorMain)
        'Draw the solid progress bar
        e.Graphics.FillRectangle(brushSolid, rectProgress)
      Else
        brushGradient = New LinearGradientBrush(Me.ClientRectangle, BarColorMain, BarColorGradient, LinearGradientMode.Vertical)
        'Draw the gradient progress bar
        e.Graphics.FillRectangle(brushGradient, rectProgress)
      End If
    End If

    If Not brushGradient Is Nothing Then brushGradient.Dispose()
    If Not brushSolid Is Nothing Then brushSolid.Dispose()

    ' Draw a border around the control.
    If Me.BorderStyle <> Windows.Forms.BorderStyle.None Then DrawBorder(e.Graphics)
  End Sub

  Protected Overrides Sub OnResize(ByVal e As EventArgs)
    ' Invalidate the control to get a repaint.
    Me.Invalidate()
  End Sub

  Private Sub DrawBorder(ByVal g As Graphics)
    Dim PenLT, PenRB As Pen
    Dim PenWidth As Integer = CInt(Pens.White.Width)

    If Me.BorderStyle = Windows.Forms.BorderStyle.Fixed3D Then
      PenLT = New Pen(SystemColors.ControlDark, 1.0F)
      PenRB = New Pen(SystemColors.ControlLightLight, 1.0F)
    Else
      PenLT = Pens.Black : PenRB = Pens.Black
    End If

    With Me.ClientRectangle
      g.DrawLine(PenLT, New Point(.Left, .Top), New Point(.Width - PenWidth, .Top))
      g.DrawLine(PenLT, New Point(.Left, .Top), New Point(.Left, .Height - PenWidth))
      g.DrawLine(PenRB, New Point(.Left, .Height - PenWidth), New Point(.Width - PenWidth, .Height - PenWidth))
      g.DrawLine(PenRB, New Point(.Width - PenWidth, .Top), New Point(.Width - PenWidth, .Height - PenWidth))
    End With
  End Sub

  Private ReadOnly Property Percent As Decimal
    Get
      If Value > 0 AndAlso Maximum > 0 Then
        If Value > Minimum AndAlso Maximum > Minimum Then
          Return Convert.ToDecimal(Value - Minimum) / Convert.ToDecimal(Maximum - Minimum)
        End If
      End If
      Return 0
    End Get
  End Property

#Region " Public Properties "

#Region " Behavior Properties "

  Private _Minimum As Integer
  <Category("Behavior"), RefreshProperties(RefreshProperties.All), DefaultValue(0), Description("The lower bound of the range this ProgressBar is working with.")> _
  Public Property Minimum() As Integer
    Get
      Return _Minimum
    End Get
    Set(ByVal Value As Integer)
      _Minimum = Value
      ' Make sure that the minimum value is never set >= the maximum value.
      If _Minimum >= Maximum Then _Minimum = Maximum - 1
      ' Prevent a negative
      If _Minimum < 0 Then _Minimum = 0
      ' Force a redraw
      Me.Invalidate()
    End Set
  End Property

  Private _Maximum As Integer
  <Category("Behavior"), RefreshProperties(RefreshProperties.All), DefaultValue(100), Description("The upper bound of the range this ProgressBar is working with.")> _
  Public Property Maximum() As Integer
    Get
      Return _Maximum
    End Get
    Set(ByVal Value As Integer)
      _Maximum = Value
      ' Make sure that the maximum value is never set <= the minimum value.
      If _Maximum <= Minimum Then _Maximum = Minimum + 1
      ' Prevent a negative
      If _Maximum < 0 Then _Maximum = 2
      ' Force a redraw
      Me.Invalidate()
    End Set
  End Property

  Private _Value As Integer
  <Category("Behavior"), RefreshProperties(RefreshProperties.All), DefaultValue(0), Description("The current value for the ProgressBar, in the range specified by the Minimum and Maximum properties.")> _
  Public Property Value() As Integer
    Get
      Return _Value
    End Get
    Set(ByVal Value As Integer)
      _Value = Value
      ' Make sure that the value does not stray outside the valid range.
      If _Value < Minimum Then _Value = Minimum
      If _Value > Maximum Then _Value = Maximum
      ' Force a redraw
      Me.Invalidate()
    End Set
  End Property

#End Region

#Region " Appearance Properties "

  Private _BarColorMain As Color = SystemColors.Highlight     ' Color of progress meter
  <Category("Appearance"), DefaultValue(GetType(Color), "Highlight"), Description("The foreground color used to display the current progress value of the ProgressBar.")> _
  Public Property BarColorMain() As Color
    Get
      Return _BarColorMain
    End Get
    Set(ByVal Value As Color)
      _BarColorMain = Value
      Me.Invalidate()
    End Set
  End Property

  Private _BarColorGradient As Color = SystemColors.HighlightText ' End color for gradient fill
  <Category("Appearance"), DefaultValue(GetType(Color), "HighlightText"), Description("The foreground end color used to display the current progress value of the ProgressBar in Gradient mode.")> _
  Public Property BarColorGradient() As Color
    Get
      Return _BarColorGradient
    End Get
    Set(ByVal Value As Color)
      _BarColorGradient = Value
      Me.Invalidate()
    End Set
  End Property

  Private _BackColorMain As Color = SystemColors.ControlDark
  <Category("Appearance"), DefaultValue(GetType(Color), "ControlDark"), Description("The background color of the ProgressBar.")> _
  Public Property BackColorMain() As Color
    Get
      Return _BackColorMain
    End Get
    Set(ByVal Value As Color)
      _BackColorMain = Value
      Me.Invalidate()
    End Set
  End Property

  Private _BackColorGradient As Color = SystemColors.ControlLight
  <Category("Appearance"), DefaultValue(GetType(Color), "ControlLight"), Description("The background end color of the ProgressBar in Gradient mode.")> _
  Public Property BackColorGradient() As Color
    Get
      Return _BackColorGradient
    End Get
    Set(ByVal Value As Color)
      _BackColorGradient = Value
      Me.Invalidate()
    End Set
  End Property

  Private _BorderStyle As System.Windows.Forms.BorderStyle = Windows.Forms.BorderStyle.Fixed3D
  <Category("Appearance"), DefaultValue(GetType(System.Windows.Forms.BorderStyle), "Fixed3D"), Description("Indicates whether or not the ProgressBar should have a border.")> _
  Public Shadows Property BorderStyle() As System.Windows.Forms.BorderStyle
    Get
      Return _BorderStyle
    End Get
    Set(ByVal Value As System.Windows.Forms.BorderStyle)
      _BorderStyle = Value
      Me.Invalidate()
    End Set
  End Property

  Private _BarColorStyle As ColorStyle = ColorStyle.Solid
  <Category("Appearance"), DefaultValue(GetType(ColorStyle), "Solid"), Description("The foreground color style of the ProgressBar.")> _
  Public Property BarColorStyle() As ColorStyle
    Get
      Return _BarColorStyle
    End Get
    Set(ByVal Value As ColorStyle)
      _BarColorStyle = Value
      Me.Invalidate()
    End Set
  End Property

  Private _BackColorStyle As ColorStyle = ColorStyle.Solid
  <Category("Appearance"), DefaultValue(GetType(ColorStyle), "Solid"), Description("The background color style of the ProgressBar.")> _
  Public Property BackColorStyle() As ColorStyle
    Get
      Return _BackColorStyle
    End Get
    Set(ByVal Value As ColorStyle)
      _BackColorStyle = Value
      Me.Invalidate()
    End Set
  End Property

#End Region

#Region " Shadows Properties "

  <Browsable(False)> _
  Shadows ReadOnly Property ForeColor() As Color
    Get
      Return MyBase.ForeColor
    End Get
  End Property

  <Browsable(False)> _
  Shadows ReadOnly Property Font() As Font
    Get
      Return MyBase.Font
    End Get
  End Property

  <Browsable(False)> _
  Shadows ReadOnly Property Text() As String
    Get
      Return MyBase.Text
    End Get
  End Property

  <Browsable(False)> _
  Shadows ReadOnly Property BackColor() As Color
    Get
      Return MyBase.BackColor
    End Get
  End Property
#End Region

#End Region

End Class
