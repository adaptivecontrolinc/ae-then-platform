Public Class ControlKeyboardDecimalBad

  Public Event KeyClick(ByVal key As String)
  Public Event KeyClear()

  Private keys() As String = {"7", "8", "9", _
                              "4", "5", "6", _
                              "1", "2", "3", _
                              "0", ".", "Clr"}

  Private keySize_ As New Drawing.Size(36, 32)
  Public Property KeySize() As Drawing.Size
    Get
      Return keySize_
    End Get
    Set(ByVal value As Drawing.Size)
      If value <> KeySize Then
        keySize_.Width = value.Width
        keySize_.Height = value.Height
        InitializeControl()
      End If
    End Set
  End Property

  Private textBoxAmount_ As New Windows.Forms.TextBox
  Public ReadOnly Property TextBoxAmount() As Windows.Forms.TextBox
    Get
      Return textBoxAmount_
    End Get
  End Property

  Public Property ShowAmount() As Boolean
    Get
      Return TextBoxAmount.Visible
    End Get
    Set(ByVal value As Boolean)
      If value <> TextBoxAmount.Visible Then
        TextBoxAmount.Visible = value
        InitializeControl()
      End If
    End Set
  End Property

  Public Property Amount() As String
    Get
      Return TextBoxAmount.Text
    End Get
    Set(ByVal value As String)
      TextBoxAmount.Text = value
    End Set
  End Property

  Public Property AmountInteger As Integer
    Get
      Dim tryInteger As Integer
      If Integer.TryParse(Amount, tryInteger) Then Return tryInteger
      Return 0
    End Get
    Set(ByVal value As Integer)
      Amount = value.ToString
    End Set
  End Property

  Public Property AmountDouble As Double
    Get
      Dim tryDouble As Double
      If Double.TryParse(Amount, tryDouble) Then Return tryDouble
      Return 0
    End Get
    Set(ByVal value As Double)
      Amount = value.ToString
    End Set
  End Property

  Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    InitializeControl()
  End Sub

  Private Sub InitializeControl()
    'Clear any existing controls
    Me.Controls.Clear()

    'Setup text box 
    With TextBoxAmount
      .Location = New Drawing.Point(0, 0)
      .Name = "textBoxAmount"
      .Size = New Drawing.Size(KeySize.Width * 3, 22)
      .TextAlign = HorizontalAlignment.Right
    End With
    Me.Controls.Add(TextBoxAmount)

    'Add number pad keys
    For i As Integer = keys.GetLowerBound(0) To keys.GetUpperBound(0)
      Dim key As Windows.Forms.Button = GetKey(i)
      Me.Controls.Add(key)
    Next
  End Sub

  Private Function GetKey(ByVal index As Integer) As Windows.Forms.Button
    Dim row As Integer = GetRow(index)
    Dim column As Integer = GetColumn(index)
    Dim key As New Windows.Forms.Button
    With key
      .Name = "button" & keys(index)
      .Text = keys(index)
      .Size = KeySize  'New Drawing.Size(keyWidth, keyHeight)
      .Location = GetKeyPosition(index)
      AddHandler key.Click, AddressOf KeyClickHandler
    End With
    Return key
  End Function

  Private Function GetRow(ByVal index As Integer) As Integer
    Dim totalItems As Integer = keys.GetLength(0)
    Dim itemsPerRow As Integer = 3

    Dim row As Integer = (index \ itemsPerRow) + 1
    Return row
  End Function

  Private Function GetColumn(ByVal index As Integer) As Integer
    Dim totalItems As Integer = keys.GetLength(0)
    Dim itemsPerColumn As Integer = 4

    Dim column As Integer = (index Mod 3) + 1
    Return column
  End Function

  Private Function GetKeyPosition(ByVal index As Integer) As Drawing.Point
    Dim amountOffset As Integer = 0
    If ShowAmount Then amountOffset = 26
    Select Case index
      Case 0 To 2
        Return New Drawing.Point(KeySize.Width * index, 0 + amountOffset)

      Case 3 To 5
        Return New Drawing.Point(KeySize.Width * (index - 3), KeySize.Height + amountOffset)

      Case 6 To 8
        Return New Drawing.Point(KeySize.Width * (index - 6), (KeySize.Height * 2) + amountOffset)

      Case 9 To 11
        Return New Drawing.Point(KeySize.Width * (index - 9), (KeySize.Height * 3) + amountOffset)

    End Select
  End Function

  Private Sub KeyClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    With DirectCast(sender, Windows.Forms.Button)
      Select Case .Text
        Case "Clr"
          TextBoxAmount.Text = Nothing
          RaiseEvent KeyClear()
        Case Else
          TextBoxAmount.Text &= .Text
          RaiseEvent KeyClick(.Text)
      End Select
    End With
  End Sub

End Class
