Public NotInheritable Class Memo
  Private Shared ReadOnly cr_ As Char = Convert.ToChar(13), lf_ As Char = Convert.ToChar(10)
  Private Sub New()
  End Sub
  Public Shared Function Split(ByVal value As String) As Collections.ObjectModel.ReadOnlyCollection(Of String)
    ' Assemble into a collection
    Dim coll As New List(Of String)
    If Not String.IsNullOrEmpty(value) Then
      Dim startIndex As Integer, valueLength As Integer = value.Length
      Do While startIndex < valueLength
        Dim find As Integer = startIndex
        Do
          Dim ch As Char = value.Chars(find)
          If ch = cr_ OrElse ch = lf_ Then Exit Do
          find += 1 : If find = valueLength Then coll.Add(value.Substring(startIndex)) : GoTo retNow
        Loop
        coll.Add(value.Substring(startIndex, find - startIndex))
        startIndex = find + 1
        If startIndex = valueLength Then Exit Do
        If value.Chars(startIndex) = lf_ Then startIndex += 1
      Loop
      ' And return it as a simple string array
retNow:
    End If
    Return coll.AsReadOnly
  End Function

  Public Shared Function Join(ByVal ParamArray str() As String) As String
    Return Join(DirectCast(str, IEnumerable))
  End Function

  Public Shared Function Join(ByVal value As IEnumerable) As String
    If value Is Nothing Then Return String.Empty
    With New System.Text.StringBuilder
      Dim isSubsequent As Boolean
      For Each obj As Object In value
        If isSubsequent Then
          .Append(Convert.ToChar(13) & Convert.ToChar(10))
        Else
          isSubsequent = True
        End If
        If obj IsNot Nothing Then .Append(obj.ToString)
      Next obj
      Return .ToString
    End With
  End Function
End Class
