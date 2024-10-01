Module Tables

  ' Make a blank BulkedRecipe table with only the columns we use
  Public Function BulkedRecipe() As System.Data.DataTable
    Dim table As New System.Data.DataTable
    With table
      .TableName = "DyelotsBulkedRecipe"

      .Columns.Add("ID", GetType(Integer))
      .Columns.Add("CentralID", GetType(Integer))

      .Columns.Add("DyelotID", GetType(Integer))
      .Columns.Add("Dyelot", GetType(String))
      .Columns.Add("ReDye", GetType(Integer))

      .Columns.Add("StepNumber", GetType(Integer))
      .Columns.Add("ProductID", GetType(Integer))
      .Columns.Add("ProductCode", GetType(String))
      .Columns.Add("ProductName", GetType(String))

      .Columns.Add("Amount", GetType(Double))
      .Columns.Add("Units", GetType(String))
      .Columns.Add("Grams", GetType(Double))
      .Columns.Add("Pounds", GetType(Double))
      .Columns.Add("Liters", GetType(Double))
      .Columns.Add("Cost", GetType(Double))

      .Columns.Add("Created", GetType(Date))
      .Columns.Add("CreatedBy", GetType(String))
      .Columns.Add("Notes", GetType(String))

      'Set primary key - except we don't really need to
      '.PrimaryKey = New DataColumn() {.Columns("ID")}
    End With
    Return table
  End Function
End Module
