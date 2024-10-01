Public Class KeysNumeric

  Public Event KeyClick(ByVal text As String)

  Public Property KeySpacing As Integer = 2
  Private keyWidth As Integer = 48
  Private keyHeight As Integer = 36

  Private keys() As String = {"7", "8", "9", "4", "5", "6", "1", "2", "3", "0", ".", "Del"}

  Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    InitializeControl()
  End Sub

  Private Sub KeysNumeric_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
    LayoutKeys()
  End Sub

  Private Sub InitializeControl()
    For i As Integer = keys.GetLowerBound(0) To keys.GetUpperBound(0)
      Dim key As ControlButtonNoFocus = GetKey(i)
      key.TabStop = False
      Me.Controls.Add(key)
    Next
    LayoutKeys()
  End Sub

  Private Function GetKey(ByVal index As Integer) As ControlButtonNoFocus
    Dim key As New ControlButtonNoFocus
    With key
      .Location = New Drawing.Point(0, 0)
      .Name = GetKeyName(index)
      .Size = New Drawing.Size(keyWidth, keyHeight)
      .Tag = index.ToString
      .Text = keys(index)
      .Font = New System.Drawing.Font("Tahoma", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      AddHandler key.KeyClick, AddressOf KeyClickHandler
    End With
    Return key
  End Function

  Private Function GetKeyName(ByVal index As Integer) As String
    Dim tryInteger As Integer
    Select Case keys(index)
      Case "." : Return "aiButtonDecimalPoint"
      Case "Del" : Return "aiButtonDelete"
      Case Else
        If Integer.TryParse(keys(index).ToString, tryInteger) Then
          Return "aiButtonNumber" & keys(index)
        End If
    End Select
    Return Nothing
  End Function

  Private Sub LayoutKeys()
    keyWidth = Convert.ToInt32((Me.Width - (KeySpacing * 2)) / 3)
    keyHeight = Convert.ToInt32((Me.Height - (KeySpacing * 3)) / 4)

    For Each key As ControlButtonNoFocus In Me.Controls
      If key.Name.Contains("aiButton") Then LayoutKey(key)
    Next
  End Sub

  Private Sub LayoutKey(ByVal key As ControlButtonNoFocus)
    ' Get key index - helpfully stored in the tag property
    Dim index As Integer = Integer.Parse(key.Tag.ToString)
    ' No offset rows
    Dim keyOrder As Integer
    Select Case index
      Case 0 To 2
        keyOrder = index
        key.Size = New Drawing.Size(keyWidth, keyHeight)
        key.Location = New Drawing.Point((keyOrder * KeySpacing) + (keyWidth * keyOrder), 0)

      Case 3 To 5
        keyOrder = index - 3
        key.Size = New Drawing.Size(keyWidth, keyHeight)
        key.Location = New Drawing.Point((keyOrder * KeySpacing) + (keyWidth * keyOrder), keyHeight + KeySpacing)

      Case 6 To 8
        keyOrder = index - 6
        key.Size = New Drawing.Size(keyWidth, keyHeight)
        key.Location = New Drawing.Point((keyOrder * KeySpacing) + (keyWidth * keyOrder), (keyHeight + KeySpacing) * 2)

      Case 9 To 11
        keyOrder = index - 9
        key.Size = New Drawing.Size(keyWidth, keyHeight)
        key.Location = New Drawing.Point((keyOrder * KeySpacing) + (keyWidth * keyOrder), (keyHeight + KeySpacing) * 3)

    End Select
  End Sub

  Private Sub KeyClickHandler(ByVal sender As Object, ByVal text As String)
    Select Case text
      Case "Del"
        RaiseEvent KeyClick("{BACKSPACE}")
      Case "Ent"
        RaiseEvent KeyClick("{ENTER}")
      Case Else
        RaiseEvent KeyClick(text)
    End Select
  End Sub

End Class
