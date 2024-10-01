Public Class ControlDataRowEdit
  'Edit a data row using a list view control and a text box that is automatically moved to the item to be edited
  '  Uses a separate scroll bar control 'cos there are no events associated with the 'built-in' ListView scroll bar
  '    and we need to detect scrolls to hide the text box and stop some ugly behaviour

  Private dataTable_ As System.Data.DataTable
  Public Property DataTable() As System.Data.DataTable
    Get
      Return dataTable_
    End Get
    Set(ByVal value As System.Data.DataTable)
      dataTable_ = value
    End Set
  End Property

  Private dataRow_ As System.Data.DataRow
  Public Property DataRow() As System.Data.DataRow
    Get
      Return dataRow_
    End Get
    Set(ByVal value As System.Data.DataRow)
      dataRow_ = value
    End Set
  End Property

  Public Sub New()
    ' This call is required by the Windows Form Designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    InitializeControl()
  End Sub

  Private Sub InitializeControl()
    SetupListView()
    SetupTextBox()
  End Sub

  Public Sub Fill(ByVal dataTable As System.Data.DataTable, ByVal dataRow As System.Data.DataRow)
    Me.DataTable = dataTable
    Me.DataRow = dataRow
    FillListView()
    ShowTextBox()
    ShowScrollBars()
  End Sub

  Public Sub Refill()
    FillListView()
  End Sub

#Region " List View "

  Private Sub SetupListView()
    With Me.ListView
      .View = View.Details
      .BorderStyle = Windows.Forms.BorderStyle.Fixed3D
      .GridLines = True
      .FullRowSelect = True
      .HeaderStyle = ColumnHeaderStyle.Nonclickable
      .LabelEdit = False
      .LabelWrap = False
      .MultiSelect = False
      .Scrollable = False

      .Columns.Clear()
      .Columns.Add("Description", 160, HorizontalAlignment.Left)
      .Columns.Add("Value", 224, HorizontalAlignment.Left)
    End With
  End Sub

  Private Sub FillListView()
    With Me.ListView
      .Items.Clear()
      For Each dc As System.Data.DataColumn In Me.DataTable.Columns
        Dim newItem As New ListViewItem(dc.ColumnName)
        newItem.SubItems.Add(Me.DataRow(dc.ColumnName).ToString)
        .Items.Add(newItem)
      Next
    End With
  End Sub

  Private Sub UpdateListView()
    'This may seem a little over the top but I was getting some weird behaviour....
    '  the text would not update correctly
    With Me.ListView
      If .SelectedItems.Count > 0 Then
        Dim item As ListViewItem = .SelectedItems(0)
        With item.SubItems(1)
          .Name = Nothing
          .Text = Nothing

          .Name = Me.TextBox.Text.Trim
          .Text = Me.TextBox.Text.Trim
        End With
      End If
    End With
  End Sub

  Private Sub ListView_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles ListView.KeyDown
    If e.KeyCode = Keys.Enter Then
      e.Handled = True
      e.SuppressKeyPress = True
      ShowTextBox()
    End If
  End Sub

  Private Sub ListView_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles ListView.MouseClick
    ShowTextBox()
  End Sub

  Private Sub ListView_ColumnWidthChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnWidthChangedEventArgs) Handles ListView.ColumnWidthChanged
    ShowTextBox()
  End Sub

#End Region

#Region " Text Box "

  Private Sub SetupTextBox()
    With Me.TextBox
      .BorderStyle = Windows.Forms.BorderStyle.Fixed3D
      .Multiline = True
      .Visible = False
    End With
  End Sub

  Private Sub ShowTextBox()
    With Me.TextBox
      .Visible = False
      If Me.ListView.SelectedItems.Count > 0 Then
        'Get the location - make sure have one
        Dim location As Drawing.Point = GetLocation()
        If location = Nothing Then Exit Sub

        'Get the size - make sure have one
        Dim size As Drawing.Size = GetSize()
        If size = Nothing Then Exit Sub

        'Set location size and text
        .Location = location
        .Size = size
        .Text = Me.ListView.SelectedItems(0).SubItems(1).Text
        .Select(.Text.Length, 0)

        'Set the selection point

        'Finally make the text box visible
        .Visible = True
        .Focus()
      End If
    End With
  End Sub

  Private Function GetBounds() As System.Drawing.Rectangle
    With Me.ListView
      ' Make sure we have a selected item
      If .SelectedItems.Count <= 0 Then Return Nothing

      ' Return the bounds of the sub item - tweak it a bit
      Dim rect As System.Drawing.Rectangle = .SelectedItems(0).SubItems(1).Bounds
      rect.X += 2 : rect.Y += 1
      rect.Height += 1 : rect.Width += 1

      Return rect
    End With
  End Function

  Private Function GetLocation() As Drawing.Point
    Dim rect As System.Drawing.Rectangle = GetBounds()

    If rect = Nothing Then Return Nothing
    Return New System.Drawing.Point(rect.X, rect.Y - 1)
  End Function

  Private Function GetSize() As Drawing.Size
    Dim rect As System.Drawing.Rectangle = GetBounds()

    If rect = Nothing Then Return Nothing
    Return New System.Drawing.Size(rect.Width, rect.Height + 2)
  End Function

  Private Sub TextBox_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox.KeyDown
    'Intercept the enter key - stops multiline behaviour...
    If e.KeyCode = Keys.Enter Then
      e.Handled = True
      e.SuppressKeyPress = True
      TextBox.Visible = False
    End If
  End Sub

  Private Sub TextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox.TextChanged
    UpdateListView()
  End Sub

#End Region

#Region " Scroll bars "

  Private Sub ListView_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView.SizeChanged
    'Hide the edit text bar during resize
    TextBox.Visible = False
    'Check to see if we should show a scroll bar
    ShowScrollBars()
  End Sub

  Private Sub ShowScrollBars()
    'Check to see if we need to show scrollbars...
    Dim listViewItemHeight As Integer = 15
    Dim listViewHeaderHeight As Integer = 15
    With Me.ListView
      Dim itemHeight As Integer = .Items.Count * listViewItemHeight
      If itemHeight > .Height - listViewHeaderHeight Then
        .Scrollable = True
        With verticalScrollBar
          .LargeChange = 4
          .Maximum = Me.ListView.Items.Count - 1
          .Minimum = 0
          .SmallChange = 2
          .Visible = True
        End With
      Else
        .Scrollable = False
        verticalScrollBar.Visible = False
      End If
    End With
  End Sub

  Private Sub verticalScrollBar_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles verticalScrollBar.ValueChanged
    'Hide the edit text bar if any scrolling is going on
    TextBox.Visible = False

    'Make sure the scroll bar is visible
    If Not verticalScrollBar.Visible Then Exit Sub

    'Save current top item (if any)
    Dim curTopItem As ListViewItem = ListView.TopItem
    If curTopItem Is Nothing Then Exit Sub

    'Try to scroll from the current top item to the scrollbar value
    Dim curTopIndex As Integer = curTopItem.Index
    Dim targetIndex As Integer = verticalScrollBar.Value
    Dim direction As Integer = 1 : If targetIndex < curTopIndex Then direction = -1
    With Me.ListView
      If targetIndex < 0 Or targetIndex > ListView.Items.Count - 1 Then Exit Sub
      For i As Integer = curTopIndex To targetIndex Step direction
        Me.ListView.TopItem = Me.ListView.Items(i)
      Next
    End With
  End Sub

#End Region

End Class
