Public Class ListViewBuffered : Inherits Windows.Forms.ListView
  
  Private saveSelectedItem As ListViewItem
  Private saveTopItem As ListViewItem

  Public Sub New()
    MyBase.New()
    MyBase.DoubleBuffered = True
  End Sub

  Public Sub SaveSelections(hide As Boolean)
    If hide Then MyBase.Visible = False
    SaveSelections()
  End Sub

  Public Sub SaveSelections()
    saveSelectedItem = Nothing
    saveTopItem = MyBase.TopItem
    If Me.SelectedItems.Count > 0 Then saveSelectedItem = MyBase.SelectedItems(0)
  End Sub

  Public Sub ResetScroll()
    saveSelectedItem = Nothing
  End Sub

  Public Sub RestoreSelections(show As Boolean)
    RestoreSelections()
    If show Then MyBase.Visible = True
  End Sub

  Public Sub RestoreSelections()
    'Restore selection first
    If saveSelectedItem IsNot Nothing Then
      For Each item As ListViewItem In MyBase.Items
        If item.Name = saveSelectedItem.Name Then
          item.Selected = True
          Exit For
        End If
      Next
    End If

    ' then restore top item
    If saveTopItem IsNot Nothing Then
      For Each item As ListViewItem In MyBase.Items
        If item.Name = saveTopItem.Name Then
          For i As Integer = 0 To 7  ' try this up to 8 times
            MyBase.TopItem = item
            If MyBase.TopItem Is item Then Exit Sub
          Next
        End If
      Next
    End If
  End Sub

End Class
