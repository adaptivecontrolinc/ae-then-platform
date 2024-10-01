Public Class KeysQwerty

  Public Event KeyClick(ByVal text As String)

  Private keyWidth As Integer = 34
  Private keyHeight As Integer = 30
  Private keySpacing As Integer = 2

  Dim keys() As String = {"1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "-", "+", _
                          "Q", "W", "E", "R", "T", "Y", "U", "I", "O", "P", "%", _
                          "A", "S", "D", "F", "G", "H", "J", "K", "L", "Ent", _
                          "Z", "X", "C", "V", "B", "N", "M", "(", ")", _
                          "Caps", "Space", "Del"}

  Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    InitializeControl()
  End Sub

  Private Sub InitializeControl()
    For i As Integer = keys.GetLowerBound(0) To keys.GetUpperBound(0)
      Dim key As ControlButtonNoFocus = GetKey(i)
      key.TabStop = False
      Me.Controls.Add(key)
    Next
    Me.ActiveControl = Me.Controls("buttonSpace")
  End Sub

  Private Function GetKey(ByVal index As Integer) As ControlButtonNoFocus
    Dim key As New ControlButtonNoFocus
    With key
      .Location = GetKeyPosition(index)
      .Name = GetKeyName(index)
      .Size = GetKeySize(index)
      .Text = keys(index)
      AddHandler key.KeyClick, AddressOf KeyClickHandler
    End With
    Return key
  End Function

  Private Function GetKeyName(ByVal index As Integer) As String
    Dim tryInteger As Integer
    Select Case keys(index)
      Case "-" : Return "buttonMinus"
      Case "+" : Return "buttonPlus"
      Case "%" : Return "buttonPercent"
      Case "Ent" : Return "buttonEnter"
      Case "(" : Return "buttonOpenBracket"
      Case "(" : Return "buttonCloseBracket"
      Case "Caps" : Return "buttonCaps"
      Case "Space" : Return "buttonSpace"
      Case "Del" : Return "buttonDelete"
      Case Else
        If Integer.TryParse(keys(index).ToString, tryInteger) Then
          Return "buttonNumber" & keys(index)
        Else
          Return "buttonLetter" & keys(index)
        End If
    End Select
  End Function

  Private Function GetKeySize(ByVal index As Integer) As Size
    Select Case keys(index)
      Case "Ent" : Return New Drawing.Size(keyWidth + (keyWidth \ 2) + keySpacing, keyHeight)
      Case "Caps" : Return New Drawing.Size(keyWidth + (keyWidth \ 2) + keySpacing, keyHeight)
      Case "Space" : Return New Drawing.Size((keyWidth * 5) + (keySpacing * 4), keyHeight)
      Case "Del" : Return New Drawing.Size(keyWidth + (keyWidth \ 2) + keySpacing, keyHeight)
      Case Else
        Return New Drawing.Size(keyWidth, keyHeight)
    End Select
  End Function

  Private Function GetKeyPosition(ByVal index As Integer) As Drawing.Point
    Dim rowOffset As Integer, rowCount As Integer
    Select Case index
      Case 0 To 11
        rowOffset = 0
        rowCount = index
        Return New Drawing.Point(rowOffset + (rowCount * keySpacing) + (keyWidth * rowCount), 0)

      Case 12 To 22
        rowOffset = keyWidth \ 2
        rowCount = index - 12
        Return New Drawing.Point(rowOffset + (rowCount * keySpacing) + (keyWidth * rowCount), keyHeight + keySpacing)

      Case 23 To 32
        rowOffset = keyWidth
        rowCount = index - 23
        Return New Drawing.Point(rowOffset + (rowCount * keySpacing) + (keyWidth * rowCount), (keyHeight + keySpacing) * 2)

      Case 33 To 41
        rowOffset = keyWidth + (keyWidth \ 2)
        rowCount = index - 33
        Return New Drawing.Point(rowOffset + (rowCount * keySpacing) + (keyWidth * rowCount), (keyHeight + keySpacing) * 3)

      Case 42 'Caps key
        rowOffset = (keyWidth * 2)
        Return New Drawing.Point(rowOffset, (keyHeight + keySpacing) * 4)

      Case 43 'Space key
        rowOffset = (keyWidth * 3) + (keySpacing * 2) + keyWidth \ 2
        Return New Drawing.Point(rowOffset, (keyHeight + keySpacing) * 4)

      Case 44 'Del key
        rowOffset = (keyWidth * 8) + (keySpacing * 7) + keyWidth \ 2
        Return New Drawing.Point(rowOffset, (keyHeight + keySpacing) * 4)

    End Select
  End Function

  Private Sub KeyClickHandler(ByVal sender As Object, ByVal text As String)
    Select Case text
      Case "Caps"
        ToggleCaps()
      Case "Del"
        RaiseEvent KeyClick("{BACKSPACE}")
      Case "Ent"
        RaiseEvent KeyClick("{ENTER}")
      Case "Space"
        RaiseEvent KeyClick(" ")
      Case Else
        RaiseEvent KeyClick(text)
    End Select
  End Sub

  Private Sub ToggleCaps()
    Try
      For Each button As ControlButtonNoFocus In Me.Controls
        With button
          If .Name.Contains("buttonLetter") Then
            If Char.IsUpper(.Text.Chars(0)) Then
              .Text = .Text.ToLower()
            Else
              .Text = .Text.ToUpper()
            End If
          End If
        End With
      Next
    Catch ex As Exception
      'Ignore errors...
    End Try
  End Sub

End Class
