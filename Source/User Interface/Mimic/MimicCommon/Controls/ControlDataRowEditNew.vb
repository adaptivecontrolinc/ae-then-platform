Public Class ControlDataRowEditNew
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
    'InitializeControl()
  End Sub

  
End Class
