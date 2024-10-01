Public Class DataGridViewX : Inherits System.Windows.Forms.DataGridView

  Public Sub New()
    MyBase.New()
    MyBase.DoubleBuffered = True
  End Sub
  Public Sub New(ByVal name As String)
    MyBase.Name = name
    MyBase.DoubleBuffered = True
  End Sub

  Public Sub XSetupReadOnly()
    With Me
      .AllowUserToAddRows = False
      .AllowUserToDeleteRows = False
      .AllowUserToOrderColumns = False
      .AllowUserToResizeColumns = False
      .AllowUserToResizeRows = False
      .AutoGenerateColumns = False
      .BackgroundColor = Color.Silver
      .BorderStyle = System.Windows.Forms.BorderStyle.None
      .CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.Single
      .ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
      .ColumnHeadersHeight = 24
      .EditMode = DataGridViewEditMode.EditProgrammatically
      .MultiSelect = False
      .ReadOnly = True
      .RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
      .RowHeadersWidth = 28
      .ScrollBars = ScrollBars.Vertical
      .SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
    End With
  End Sub

  Public Sub XSetupEdit()
    With Me
      .AllowUserToAddRows = False
      .AllowUserToDeleteRows = True
      .AllowUserToOrderColumns = False
      .AllowUserToResizeColumns = False
      .AllowUserToResizeRows = False
      .AutoGenerateColumns = False
      .BackgroundColor = Color.Silver
      .BorderStyle = System.Windows.Forms.BorderStyle.None
      .CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.Single
      .ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
      .ColumnHeadersHeight = 24
      .EditMode = DataGridViewEditMode.EditOnEnter
      .MultiSelect = False
      .ReadOnly = False
      .RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
      .RowHeadersWidth = 28
      .ScrollBars = ScrollBars.Vertical
      .SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
    End With
  End Sub

  Public Sub XSetupEditAll()
    XSetupEdit()
    With Me
      For Each column As DataGridViewColumn In .Columns
        column.ReadOnly = False
      Next
    End With
  End Sub

  Public Sub XSetupAlternatingColor()
    Dim AlternatingRowColor As New System.Drawing.Color
    AlternatingRowColor = System.Drawing.Color.FromArgb(255, 224, 255, 255)
    With Me
      .AlternatingRowsDefaultCellStyle.BackColor = AlternatingRowColor
    End With
  End Sub

  Public Sub XSetupAlternatingColor(ByVal AlternatingRowColor As System.Drawing.Color)
    With Me
      .AlternatingRowsDefaultCellStyle.BackColor = AlternatingRowColor
    End With
  End Sub

  Public Sub XHideSelectionBar()
    With Me
      .DefaultCellStyle.SelectionBackColor = .DefaultCellStyle.BackColor
      .DefaultCellStyle.SelectionForeColor = .DefaultCellStyle.ForeColor

      .RowHeadersDefaultCellStyle.SelectionBackColor = .RowHeadersDefaultCellStyle.BackColor
      .RowHeadersDefaultCellStyle.SelectionForeColor = .RowHeadersDefaultCellStyle.ForeColor

      For Each column As DataGridViewColumn In .Columns
        column.DefaultCellStyle.SelectionBackColor = .DefaultCellStyle.BackColor
        column.DefaultCellStyle.SelectionForeColor = .DefaultCellStyle.ForeColor
      Next
    End With
  End Sub

  Public Sub XAddColorColumn(ByVal name As String, ByVal width As Integer)
    Dim column As New ColorColumn
    With column
      .DataPropertyName = name
      .DefaultCellStyle = Nothing
      .HeaderText = "#"
      .Name = name
      .ReadOnly = True
      .Width = width
    End With
    MyBase.Columns.Add(column)
  End Sub

  Public Sub XAddComboColumn(ByVal name As String, ByVal headerText As String, ByVal width As Integer, ByVal dt As DataTable, ByVal valueMember As String, ByVal displayMember As String)
    Dim column As New System.Windows.Forms.DataGridViewComboBoxColumn
    With column
      .DataPropertyName = name
      .DataSource = dt
      .DisplayMember = displayMember
      .DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
      .HeaderText = headerText
      .Name = name
      .ReadOnly = True
      .ValueMember = valueMember
      .Width = width
    End With
    MyBase.Columns.Add(column)
  End Sub

  Public Sub XAddComboColumn(ByVal name As String, ByVal headerText As String, ByVal width As Integer, ByVal dt As DataTable, ByVal valueMember As String, ByVal displayMember As String, ByVal alignment As System.Windows.Forms.DataGridViewContentAlignment)
    Dim column As New System.Windows.Forms.DataGridViewComboBoxColumn
    With column
      .DataPropertyName = name
      .DataSource = dt
      .DefaultCellStyle.Alignment = alignment
      .DisplayMember = displayMember
      .DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
      .HeaderText = headerText
      .Name = name
      .ReadOnly = True
      .ValueMember = valueMember
      .Width = width
    End With
    MyBase.Columns.Add(column)
  End Sub

  Public Sub XAddComboColumn(ByVal name As String, ByVal headerText As String, ByVal width As Integer, ByVal dt As DataTable, ByVal valueMember As String, ByVal displayMember As String, ByVal dataPropertyName As String)
    Dim column As New System.Windows.Forms.DataGridViewComboBoxColumn
    With column
      .DataPropertyName = dataPropertyName
      .DataSource = dt
      .DisplayMember = displayMember
      .DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
      .HeaderText = headerText
      .Name = name
      .ReadOnly = True
      .ValueMember = valueMember
      .Width = width
    End With
    MyBase.Columns.Add(column)
  End Sub

  Public Sub XAddComboColumn(ByVal name As String, ByVal headerText As String, ByVal width As Integer, ByVal dt As DataTable, ByVal valueMember As String, ByVal displayMember As String, ByVal dataPropertyName As String, ByVal alignment As System.Windows.Forms.DataGridViewContentAlignment)
    Dim column As New System.Windows.Forms.DataGridViewComboBoxColumn
    With column
      .DataPropertyName = dataPropertyName
      .DataSource = dt
      .DefaultCellStyle.Alignment = alignment
      .DisplayMember = displayMember
      .DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
      .HeaderText = headerText
      .Name = name
      .ReadOnly = True
      .ValueMember = valueMember
      .Width = width
    End With
    MyBase.Columns.Add(column)
  End Sub

  Public Sub XAddTextColumn(ByVal name As String, ByVal width As Integer)
    Dim column As New System.Windows.Forms.DataGridViewTextBoxColumn
    With column
      .DataPropertyName = name
      .HeaderText = name
      .Name = name
      .ReadOnly = True
      .Width = width
    End With
    MyBase.Columns.Add(column)
  End Sub

  Public Sub XAddTextColumn(ByVal name As String, ByVal headerText As String, ByVal width As Integer)
    Dim column As New System.Windows.Forms.DataGridViewTextBoxColumn
    With column
      .DataPropertyName = name
      .HeaderText = headerText
      .Name = name
      .ReadOnly = True
      .Width = width
    End With
    MyBase.Columns.Add(column)
  End Sub

  Public Sub XAddTextColumn(ByVal name As String, ByVal headerText As String, ByVal width As Integer, ByVal alignment As System.Windows.Forms.DataGridViewContentAlignment)
    Dim column As New System.Windows.Forms.DataGridViewTextBoxColumn
    With column
      .DataPropertyName = name
      .DefaultCellStyle.Alignment = alignment
      .HeaderText = headerText
      .Name = name
      .ReadOnly = True
      .Width = width
    End With
    MyBase.Columns.Add(column)
  End Sub

  Public Sub XAddTextColumn(ByVal name As String, ByVal headerText As String, ByVal width As Integer, ByVal alignment As System.Windows.Forms.DataGridViewContentAlignment, ByVal format As String)
    Dim column As New System.Windows.Forms.DataGridViewTextBoxColumn
    With column
      .DataPropertyName = name
      .DefaultCellStyle.Alignment = alignment
      .DefaultCellStyle.Format = format
      .HeaderText = headerText
      .Name = name
      .ReadOnly = True
      .Width = width
    End With
    MyBase.Columns.Add(column)
  End Sub

  Private xAutosizeColumn_ As String
  Public Property XAutosizeColumn() As String
    Get
      Return xAutosizeColumn_
    End Get
    Set(ByVal value As String)
      For Each column As System.Windows.Forms.DataGridViewColumn In Me.Columns
        If column.Name.ToLower = value.ToLower Then
          xAutosizeColumn_ = column.Name
          column.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        End If
      Next
    End Set
  End Property

  Private xBorderColor_ As System.Drawing.Color = Color.DarkGray
  Public Property XBorderColor() As System.Drawing.Color
    Get
      Return xBorderColor_
    End Get
    Set(ByVal value As System.Drawing.Color)
      xBorderColor_ = value
    End Set
  End Property

#Region " Events "

  Public Event RowDoubleClick(ByVal sender As Object, ByVal dr As System.Data.DataRow)

  Private Sub ControlDataGridViewX_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles Me.CellDoubleClick
    'Make sure we have some data
    If Me.DataSource Is Nothing Then Exit Sub

    'It has to be data table 'cos we're gonna pass a data row
    If Not TypeOf Me.DataSource Is System.Data.DataTable Then Exit Sub

    'Make sure the row index is valid
    If e.RowIndex < 0 Or e.RowIndex >= Me.Rows.Count Then Exit Sub

    'The DataGrid event returns a row index - we're gonna get the underlying data row
    If Me.Rows(e.RowIndex) IsNot Nothing Then
      'Get the underlying DataRowView from the row index 
      Dim drv As System.Data.DataRowView = DirectCast(Me.Rows(e.RowIndex).DataBoundItem, System.Data.DataRowView)
      RaiseEvent RowDoubleClick(Me, drv.Row)
    End If
  End Sub

  Private Sub ControlDataGridViewX_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
    'Make sure we have some data
    If Me.DataSource Is Nothing Then Exit Sub

    'It has to be data table 'cos we're gonna pass a data row
    If Not TypeOf Me.DataSource Is System.Data.DataTable Then Exit Sub

    'Catch enter key
    If e.KeyCode = Keys.Enter Then
      If Me.SelectedRows IsNot Nothing Then
        If Me.SelectedRows.Count >= 1 Then
          'Get the underlying DataRowView from the selected row
          Dim drv As System.Data.DataRowView = DirectCast(Me.SelectedRows(0).DataBoundItem, System.Data.DataRowView)
          RaiseEvent RowDoubleClick(Me, drv.Row)
        End If
      End If
      e.Handled = True
    End If
  End Sub

  Private Sub ControlDataGridViewX_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles Me.DataError
    Utilities.Log.LogError(e.Exception)
    e.Cancel = True
  End Sub

#End Region

#Region " Color Column "

  Public Class ColorColumn
    Inherits DataGridViewColumn

    Public Sub New()
      MyBase.New(New ColorCell())
    End Sub

    Public Overrides Property CellTemplate() As DataGridViewCell
      Get
        Return MyBase.CellTemplate
      End Get
      Set(ByVal value As DataGridViewCell)

        ' Ensure that the cell used for the template is a ColorCell.
        If (value IsNot Nothing) AndAlso Not value.GetType().IsAssignableFrom(GetType(ColorCell)) Then
          Throw New InvalidCastException("Must be a ColorCell")
        End If
        MyBase.CellTemplate = value

      End Set
    End Property
  End Class

  Public Class ColorCell
    Inherits DataGridViewCell

    Protected Overrides Sub Paint(ByVal graphics As System.Drawing.Graphics, ByVal clipBounds As System.Drawing.Rectangle, ByVal cellBounds As System.Drawing.Rectangle, ByVal rowIndex As Integer, ByVal cellState As System.Windows.Forms.DataGridViewElementStates, ByVal value As Object, ByVal formattedValue As Object, ByVal errorText As String, ByVal cellStyle As System.Windows.Forms.DataGridViewCellStyle, ByVal advancedBorderStyle As System.Windows.Forms.DataGridViewAdvancedBorderStyle, ByVal paintParts As System.Windows.Forms.DataGridViewPaintParts)
      Try
        'Fill in with a white background
        Using pBack As New System.Drawing.Pen(Color.White, 1)
          graphics.FillRectangle(pBack.Brush, cellBounds)
        End Using

        'Set color if we have one
        Dim tryInteger As Integer
        If Integer.TryParse(value.ToString, tryInteger) Then
          Using pColor As New System.Drawing.Pen(Utilities.Color.ConvertRgbToColor(tryInteger), 1)
            'Dim x As Integer = cellBounds.Location.X + ((cellBounds.Width - 16) \ 2) - 1
            'Dim y As Integer = cellBounds.Location.Y + ((cellBounds.Height - 16) \ 2)

            Dim x As Integer = cellBounds.Location.X + 4
            Dim y As Integer = cellBounds.Location.Y + 4
            Dim width As Integer = cellBounds.Width - 8 - 1
            Dim height As Integer = cellBounds.Height - 8 - 1

            Dim rColor As New System.Drawing.Rectangle(x, y, width, height)
            graphics.FillRectangle(pColor.Brush, rColor)
          End Using
        End If

        'Finally draw grid lines
        Using pBorder As New System.Drawing.Pen(Me.DataGridView.GridColor, 1)
          Dim pt1 As New System.Drawing.Point(cellBounds.Location.X, cellBounds.Location.Y + cellBounds.Height - 1)  'bottom left
          Dim pt2 As New System.Drawing.Point(pt1.X + cellBounds.Width - 1, pt1.Y)                                   'bottom right
          Dim pt3 As New System.Drawing.Point(pt2.X, cellBounds.Location.Y)                                          'top right
          graphics.DrawLine(pBorder, pt1, pt2)
          graphics.DrawLine(pBorder, pt2, pt3)
        End Using
      Catch ex As Exception
      End Try
    End Sub
  End Class

#End Region

End Class
