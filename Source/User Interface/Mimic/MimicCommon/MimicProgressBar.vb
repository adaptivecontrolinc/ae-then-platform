' Smooth horizontal progress bar

Imports System.ComponentModel
Imports System.Drawing.Drawing2D

Public Class MimicProgressBar

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

  Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
    'Dim rectf As RectangleF = RectangleF.op_Implicit(Me.ClientRectangle)

    Dim rectControl As RectangleF = Me.ClientRectangle
    Dim rectProgress As RectangleF = rectControl
    Dim brushGradient As LinearGradientBrush
    Dim brushSolid As SolidBrush
    Dim brushText As SolidBrush

    ' Draw the background if gradient is selected
    If BackColorStyle = ColorStyle.Solid Then
      brushSolid = New SolidBrush(BackColorMain)
      e.Graphics.FillRectangle(brushSolid, rectControl)
    Else
      brushGradient = New LinearGradientBrush(rectControl, BackColorMain, BackColorGradient, LinearGradientMode.Vertical)
      e.Graphics.FillRectangle(brushGradient, rectControl)
    End If

    ' Calculate area for drawing the progress bar
    rectProgress.Width = rectProgress.Width * Percent

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

    'If Text IsNot Nothing Then
    If ShowTextInBar Then
      Dim text As String = Value.ToString("#0") & " / " & Maximum.ToString("#0")
      brushText = New SolidBrush(TextColor)
      Dim textFont As New Font("Tahoma", 8, FontStyle.Bold)
      Dim textFormat As New StringFormat(StringFormatFlags.NoClip) With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}

      e.Graphics.DrawString(text, textFont, brushText, rectControl, textFormat)
    End If

    If Not brushGradient Is Nothing Then brushGradient.Dispose()
    If Not brushSolid Is Nothing Then brushSolid.Dispose()
    If Not brushText Is Nothing Then brushText.Dispose()

    ' Draw a border around the control.
    If Me.BorderStyle <> Windows.Forms.BorderStyle.None Then DrawBorder(e.Graphics)
  End Sub

  Protected Overrides Sub OnResize(ByVal e As EventArgs)
    ' Invalidate the control to get a repaint.
    Me.Invalidate()
  End Sub

  Private Sub DrawBorder(ByVal g As Graphics)
    Dim PenLeftTop, PenRightBottom As Pen
    Dim PenWidth As Integer = CInt(Pens.White.Width)

    If Me.BorderStyle = Windows.Forms.BorderStyle.Fixed3D Then
      PenLeftTop = New Pen(SystemColors.ControlDark, 1.0F)
      PenRightBottom = New Pen(SystemColors.ControlLightLight, 1.0F)
    Else
      PenLeftTop = New Pen(BorderColor, 1)
      PenRightBottom = PenLeftTop
    End If

    With Me.ClientRectangle
      g.DrawLine(PenLeftTop, New Point(.Left, .Top), New Point(.Width - PenWidth, .Top))
      g.DrawLine(PenLeftTop, New Point(.Left, .Top), New Point(.Left, .Height - PenWidth))
      g.DrawLine(PenRightBottom, New Point(.Left, .Height - PenWidth), New Point(.Width - PenWidth, .Height - PenWidth))
      g.DrawLine(PenRightBottom, New Point(.Width - PenWidth, .Top), New Point(.Width - PenWidth, .Height - PenWidth))
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
    Set(ByVal value As Integer)
      _Minimum = value
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
    Set(ByVal value As Integer)
      _Maximum = value
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
    Set(ByVal value As Integer)
      _Value = value
      ' Make sure that the value does not stray outside the valid range.
      If _Value < Minimum Then _Value = Minimum
      If _Value > Maximum Then _Value = Maximum
      Me.Invalidate()  ' Force a redraw
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
    Set(ByVal value As Color)
      _BarColorMain = value
      Me.Invalidate()
    End Set
  End Property

  Private _BarColorGradient As Color = SystemColors.HighlightText ' End color for gradient fill
  <Category("Appearance"), DefaultValue(GetType(Color), "HighlightText"), Description("The foreground end color used to display the current progress value of the ProgressBar in Gradient mode.")> _
  Public Property BarColorGradient() As Color
    Get
      Return _BarColorGradient
    End Get
    Set(ByVal value As Color)
      _BarColorGradient = value
      Me.Invalidate()
    End Set
  End Property

  Private _BackColorMain As Color = SystemColors.ControlDark
  <Category("Appearance"), DefaultValue(GetType(Color), "ControlDark"), Description("The background color of the ProgressBar.")> _
  Public Property BackColorMain() As Color
    Get
      Return _BackColorMain
    End Get
    Set(ByVal value As Color)
      _BackColorMain = value
      Me.Invalidate()
    End Set
  End Property

  Private _BackColorGradient As Color = SystemColors.ControlLight
  <Category("Appearance"), DefaultValue(GetType(Color), "ControlLight"), Description("The background end color of the ProgressBar in Gradient mode.")> _
  Public Property BackColorGradient() As Color
    Get
      Return _BackColorGradient
    End Get
    Set(ByVal value As Color)
      _BackColorGradient = value
      Me.Invalidate()
    End Set
  End Property

  Private _BorderColor As Color = Color.Black
  <Category("Appearance"), DefaultValue(GetType(Color), "Black"), Description("The color of the border.")> _
  Public Shadows Property BorderColor() As Color
    Get
      Return _BorderColor
    End Get
    Set(ByVal value As Color)
      _BorderColor = value
      Me.Invalidate()
    End Set
  End Property

  Private _BorderStyle As System.Windows.Forms.BorderStyle = Windows.Forms.BorderStyle.Fixed3D
  <Category("Appearance"), DefaultValue(GetType(System.Windows.Forms.BorderStyle), "Fixed3D"), Description("Indicates whether or not the ProgressBar should have a border.")> _
  Public Shadows Property BorderStyle() As System.Windows.Forms.BorderStyle
    Get
      Return _BorderStyle
    End Get
    Set(ByVal value As System.Windows.Forms.BorderStyle)
      _BorderStyle = value
      Me.Invalidate()
    End Set
  End Property

  Private _BarColorStyle As ColorStyle = ColorStyle.Solid
  <Category("Appearance"), DefaultValue(GetType(ColorStyle), "Solid"), Description("The foreground color style of the ProgressBar.")> _
  Public Property BarColorStyle() As ColorStyle
    Get
      Return _BarColorStyle
    End Get
    Set(ByVal value As ColorStyle)
      _BarColorStyle = value
      Me.Invalidate()
    End Set
  End Property

  Private _BackColorStyle As ColorStyle = ColorStyle.Solid
  <Category("Appearance"), DefaultValue(GetType(ColorStyle), "Solid"), Description("The background color style of the ProgressBar.")> _
  Public Property BackColorStyle() As ColorStyle
    Get
      Return _BackColorStyle
    End Get
    Set(ByVal value As ColorStyle)
      _BackColorStyle = value
      Me.Invalidate()
    End Set
  End Property

  Private _TextColor As Color = Color.White
  <Category("Appearance"), DefaultValue(GetType(Color), "White"), Description("The color of the text shown in the progress bar.")> _
  Public Shadows Property TextColor() As Color
    Get
      Return _TextColor
    End Get
    Set(ByVal value As Color)
      _TextColor = value
      Me.Invalidate()
    End Set
  End Property

  <Category("Appearance"), DefaultValue(GetType(Boolean), "False"), Description("Show value in the progress bar.")> _
  Public Property ShowTextInBar As Boolean

#End Region

#Region " Shadows Properties "
  ' Basically stops these values being set from the outside

  <Browsable(False)> _
  Shadows ReadOnly Property BackColor() As Color
    Get
      Return MyBase.BackColor
    End Get
  End Property

  <Browsable(False)> _
  Shadows ReadOnly Property Font() As Font
    Get
      Return MyBase.Font
    End Get
  End Property

  <Browsable(False)> _
  Shadows ReadOnly Property ForeColor() As Color
    Get
      Return MyBase.ForeColor
    End Get
  End Property

  <Browsable(False)> _
  Shadows ReadOnly Property Text() As String
    Get
      Return MyBase.Text
    End Get
  End Property
#End Region

#End Region

End Class

