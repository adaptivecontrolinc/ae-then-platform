Imports System.Data
Imports System.Globalization.CultureInfo

Namespace Utilities.Sql
  Public Module Sql

#Region " Select "

    <System.Diagnostics.DebuggerStepThrough()>
    Public Function GetDataRow(ByVal connectionString As String, ByVal selectString As String) As DataRow
      'Get a data row from the passed connection string and select statement
      '  return nothing if no row found, more than one row found or an exception is thrown
      Try
        Dim dt As New DataTable
        Dim da As New SqlClient.SqlDataAdapter

        da.SelectCommand = New SqlClient.SqlCommand
        With da.SelectCommand
          .CommandText = selectString
          .Connection = New SqlClient.SqlConnection(connectionString)
          .Connection.Open() : da.Fill(dt) : .Connection.Close()
          If dt.Rows.Count = 1 Then Return dt.Rows(0)
        End With

      Catch ex As Exception
        Debug.Print("GetDataRow: " & selectString)
        Debug.Print(ex.Message)
        Log.LogError(ex, selectString)
      End Try
      Return Nothing
    End Function

    <System.Diagnostics.DebuggerStepThrough()>
    Public Function GetDataRow(ByVal connectionString As String, ByVal selectString As String, ByVal tableName As String) As DataRow
      'Get a data row from the passed connection string and select statement 
      '  set the table name so this row can be used with the update and insert functions here...
      '  return nothing if no row found, more than one row found or an exception is thrown
      Try
        Dim dt As New DataTable(tableName)
        Dim da As New SqlClient.SqlDataAdapter

        da.SelectCommand = New SqlClient.SqlCommand
        With da.SelectCommand
          .CommandText = selectString
          .Connection = New SqlClient.SqlConnection(connectionString)
          .Connection.Open() : da.Fill(dt) : .Connection.Close()
          If dt.Rows.Count = 1 Then Return dt.Rows(0)
        End With

      Catch ex As Exception
        Debug.Print("GetDataRow: " & selectString)
        Debug.Print(ex.Message)
        Log.LogError(ex, selectString)
      End Try
      Return Nothing
    End Function

    <System.Diagnostics.DebuggerStepThrough()>
    Public Function GetDataRow(ByVal connectionString As String, ByVal tableName As String, ByVal id As Integer) As DataRow
      'Get a data row from the passed connection string table name and id
      '  set the table name so this row can be used with the update and insert functions here...
      '  return nothing if no row found, more than one row found or an exception is thrown
      Try
        Dim dt As New DataTable(tableName)
        Dim da As New SqlClient.SqlDataAdapter

        da.SelectCommand = New SqlClient.SqlCommand
        With da.SelectCommand
          .CommandText = "SELECT * FROM " & tableName & " WHERE ID=" & id.ToString(InvariantCulture)
          .Connection = New SqlClient.SqlConnection(connectionString)
          .Connection.Open() : da.Fill(dt) : .Connection.Close()
          If dt.Rows.Count = 1 Then Return dt.Rows(0)
        End With

      Catch Ex As Exception
        Debug.Print("GetDataRow " & tableName & " " & id.ToString(InvariantCulture))
        Debug.Print(Ex.Message)
        Log.LogError(Ex)
      End Try
      Return Nothing
    End Function

    <System.Diagnostics.DebuggerStepThrough()>
    Public Function GetDataTable(ByVal connectionString As String, ByVal selectString As String) As DataTable
      Try
        'Get a data table from the passed connection string and select string
        '  return nothing if no table found or an exception is thrown
        Dim dt As New DataTable
        Dim da As New SqlClient.SqlDataAdapter

        da.SelectCommand = New SqlClient.SqlCommand
        With da.SelectCommand
          .CommandText = selectString
          .Connection = New SqlClient.SqlConnection(connectionString)
          .Connection.Open() : da.Fill(dt) : .Connection.Close()
          Return dt
        End With

      Catch ex As Exception
        Debug.Print("GetDataTable: " & selectString)
        Debug.Print(ex.Message)
        Log.LogError(ex, selectString)
      End Try
      Return Nothing
    End Function

    <System.Diagnostics.DebuggerStepThrough()>
    Public Function GetDataTable(ByVal connectionString As String, ByVal selectString As String, ByVal tableName As String) As DataTable
      'Get a data table from the passed connection string and select string
      '  set the table name so this table can be used with the update and insert functions here...
      '  return nothing if no table found or an exception is thrown
      Try
        Dim dt As New DataTable(tableName)
        Dim da As New SqlClient.SqlDataAdapter

        da.SelectCommand = New SqlClient.SqlCommand
        With da.SelectCommand
          .CommandText = selectString
          .Connection = New SqlClient.SqlConnection(connectionString)
          .Connection.Open() : da.Fill(dt) : .Connection.Close()
          Return dt
        End With

      Catch ex As Exception
        Debug.Print("GetDataTable: " & selectString)
        Debug.Print(ex.Message)
        Log.LogError(ex, selectString)
      End Try
      Return Nothing
    End Function

    <System.Diagnostics.DebuggerStepThrough()>
    Public Function GetDataTableSchema(ByVal connectionString As String, ByVal tableName As String) As DataTable
      'Get the Table Schema for the target table
      '  this will give the structure of table so we can use DataTable.NewRow
      '  assumes the table has an ID column that is AutoIncrement - not sure this has much effect
      '  return nothing if no table found or an exception is thrown
      Try
        Dim dt As New DataTable(tableName)
        Dim da As New SqlClient.SqlDataAdapter

        da.SelectCommand = New SqlClient.SqlCommand
        With da.SelectCommand
          .CommandText = "SELECT * FROM " & tableName
          .Connection = New SqlClient.SqlConnection(connectionString)
          .Connection.Open() : da.FillSchema(dt, SchemaType.Source) : .Connection.Close()
        End With

        'Set the AutoIncrement properties - this is not imported with the schema so we must manually set it
        dt.Columns("ID").AutoIncrement = True
        dt.Columns("ID").AutoIncrementSeed = 1
        dt.Columns("ID").AutoIncrementStep = 1

        Return dt
      Catch ex As Exception
        Debug.Print("GetDataTableSchema: " & tableName)
        Debug.Print(ex.Message)
        Log.LogError(ex)
      End Try
      Return Nothing
    End Function

    <System.Diagnostics.DebuggerStepThrough()>
    Public Function GetDataTableSchema(ByVal connectionString As String, ByVal tableName As String, ByVal primaryKey As String) As DataTable
      Try
        'Get the Schema for the target table
        '  this will give the structure of table so we can use DataTable.NewRow
        '  doing this here and passing a reference to reduce the number of calls to the database
        Dim dt As New DataTable(tableName)
        Dim da As New SqlClient.SqlDataAdapter

        da.SelectCommand = New SqlClient.SqlCommand
        With da.SelectCommand
          .CommandText = "SELECT * FROM " & tableName
          .Connection = New SqlClient.SqlConnection(connectionString)
          .Connection.Open() : da.FillSchema(dt, SchemaType.Source) : .Connection.Close()
        End With

        'Set primary key if it has been passed
        dt.PrimaryKey = New DataColumn() {dt.Columns(primaryKey)}

        Return dt
      Catch ex As Exception
        Debug.Print("GetDataTableSchema: " & tableName)
        Debug.Print(ex.Message)
        Log.LogError(ex)
      End Try
      Return Nothing
    End Function

    <System.Diagnostics.DebuggerStepThrough()>
    Public Function SqlSelect(ByVal connectionString As String, ByVal selectString As String) As DataTable
      Return GetDataTable(connectionString, selectString)
    End Function

#End Region

#Region " Update "
    ' Remember to include table name in data row...

    <System.Diagnostics.DebuggerStepThrough()>
    Public Function UpdateDataRow(ByVal connectionString As String, ByVal dr As DataRow) As Boolean
      'Write changes in a data row back to the database
      '  return false if write fails or an exception is thrown
      Try
        'Make sure the data row is not empty
        If dr Is Nothing Then Return False
        Dim dt As DataTable = dr.Table

        Dim da As New SqlClient.SqlDataAdapter
        da.UpdateCommand = New SqlClient.SqlCommand

        With da.UpdateCommand
          'Add parameters and build insert string based on table schema
          AddSqlParameters(dt, .Parameters)
          BuildSqlUpdateString(dt, .CommandText)

          'Set values of all parameters from the data row
          For Each dc As DataColumn In dt.Columns
            .Parameters("@" & dc.ColumnName).Value = dr(dc.ColumnName)
          Next

          .Connection = New SqlClient.SqlConnection(connectionString)
          .Connection.Open()
          Dim rows As Integer = .ExecuteNonQuery()
          .Connection.Close()
        End With

        Return True
      Catch ex As Exception
        Debug.Print("UpdateDataRow: " & dr.Table.TableName)
        Debug.Print(ex.Message)
        Log.LogError(ex)
      End Try
      Return False
    End Function

    <System.Diagnostics.DebuggerStepThrough()>
    Public Sub BuildSqlUpdateString(ByVal dt As DataTable, ByRef commandText As String)
      'Build a sql update string based on the table schema
      Dim sql As String = Nothing
      Try
        'Make sure we have a DataTable to work with
        If dt Is Nothing Then Exit Sub

        'Make the sql string
        For Each dc As DataColumn In dt.Columns
          If dc.ColumnName.ToLower <> "id" Then
            sql &= "[" & dc.ColumnName.ToString & "]=@" & dc.ColumnName.ToString & ","
          End If
        Next
        'remove the last comma
        sql = sql.Substring(0, sql.Length - 1)

        commandText = "UPDATE " & dt.TableName & " SET " & sql & " WHERE ID=@ID"
      Catch ex As Exception
        Debug.Print("MakeUpdateString: " & dt.TableName)
        Debug.Print(ex.Message)
        Log.LogError(ex, sql)
      End Try
    End Sub

    <System.Diagnostics.DebuggerStepThrough()>
    Public Function SqlUpdate(ByVal connectionString As String, ByVal updateString As String) As Integer
      Try
        Dim rows As Integer
        Dim da As New SqlClient.SqlDataAdapter
        da.UpdateCommand = New SqlClient.SqlCommand
        With da.UpdateCommand
          .CommandText = updateString
          .Connection = New SqlClient.SqlConnection(connectionString)
          .Connection.Open() : rows = .ExecuteNonQuery() : .Connection.Close()
        End With

        Return rows
      Catch ex As Exception
        Debug.Print("SqlUpdate: " & updateString)
        Debug.Print(ex.Message)
        Log.LogError(ex, updateString)
      End Try
      Return -1
    End Function

#End Region

#Region " Delete "

    <System.Diagnostics.DebuggerStepThrough()>
    Public Function SqlDelete(ByVal connectionString As String, ByVal deleteString As String) As Integer
      Try
        Dim rows As Integer
        Dim da As New SqlClient.SqlDataAdapter
        da.DeleteCommand = New SqlClient.SqlCommand
        With da.DeleteCommand
          .CommandText = deleteString
          .Connection = New SqlClient.SqlConnection(connectionString)
          .Connection.Open() : rows = .ExecuteNonQuery() : .Connection.Close()
        End With

        Return rows
      Catch ex As Exception
        Debug.Print("SqlDelete: " & deleteString)
        Debug.Print(ex.Message)
        Log.LogError(ex, deleteString)
      End Try
      Return -1
    End Function

#End Region

#Region " Insert "

    <System.Diagnostics.DebuggerStepThrough()>
    Public Function InsertDataRow(ByVal connectionString As String, ByVal dr As DataRow) As Boolean
      Try
        'Must have a valid data row
        If dr Is Nothing Then Return False

        Dim dt As DataTable = dr.Table
        Dim da As New SqlClient.SqlDataAdapter

        da.InsertCommand = New SqlClient.SqlCommand
        With da.InsertCommand

          'Add parameters and build insert string based on table schema
          AddSqlParameters(dt, .Parameters)
          BuildSqlInsertString(dt, .CommandText)

          'Set values of all parameters from the data row
          For Each dc As DataColumn In dt.Columns
            .Parameters("@" & dc.ColumnName).Value = dr(dc.ColumnName)
          Next

          Dim autoIncrement As Integer = -1
          .CommandText = .CommandText & "; SELECT CAST(scope_identity() AS int);"
          .Connection = New SqlClient.SqlConnection(connectionString)
          .Connection.Open() : autoIncrement = Integer.Parse(.ExecuteScalar().ToString) : .Connection.Close()
          dr("ID") = autoIncrement
        End With

        Return True
      Catch ex As Exception
        Debug.Print("InsertDataRow: " & dr.Table.TableName)
        Debug.Print(ex.Message)
        Log.LogError(ex)
      End Try
      Return False
    End Function

    <System.Diagnostics.DebuggerStepThrough()>
    Public Function InsertDataTable(ByVal connectionString As String, ByVal dt As DataTable) As Boolean
      Try
        'Must have a valid data table
        If dt Is Nothing Then Return False

        Dim rows As Integer
        Dim da As New SqlClient.SqlDataAdapter
        With da
          .InsertCommand = New SqlClient.SqlCommand

          With .InsertCommand
            AddSqlParameters(dt, .Parameters)
            BuildSqlInsertString(dt, .CommandText)

            .Connection = New SqlClient.SqlConnection(connectionString)
            For Each dr As DataRow In dt.Rows
              For Each dc As DataColumn In dt.Columns
                .Parameters("@" & dc.ColumnName).Value = dr(dc.ColumnName)
              Next
              .Connection.Open() : rows = .ExecuteNonQuery() : .Connection.Close()
            Next
          End With
        End With

        Return True
      Catch ex As Exception
        Debug.Print("InsertDataRow: " & dt.TableName)
        Debug.Print(ex.Message)
        Log.LogError(ex)
      End Try
      Return False
    End Function

    <System.Diagnostics.DebuggerStepThrough()>
    Public Function InsertDataTable(ByVal connectionString As String, ByVal dt As DataTable, ByVal sqlDelete As String) As Boolean
      'Insert a datatable into the database - this is used to update detail tables using a transactiion
      '  we start the transaction, delete the original records then insert the new records
      '  the transaction will be rolled back if any errors occur
      'Create the connection and declare the command and transaction objects
      Try
        Dim connection As New SqlClient.SqlConnection(connectionString)
        Dim deleteCommand As SqlClient.SqlCommand
        Dim insertCommand As SqlClient.SqlCommand
        Dim transaction As SqlClient.SqlTransaction

        'Open the connection and make the command and transaction objects on this connection
        connection.Open()
        transaction = connection.BeginTransaction
        deleteCommand = connection.CreateCommand : deleteCommand.Transaction = transaction
        insertCommand = connection.CreateCommand : insertCommand.Transaction = transaction

        Dim sql As String = Nothing
        Try
          'Delete any existing records
          With deleteCommand
            sql = sqlDelete ' just so we get the sql statement if we throw an error
            .CommandText = sql
            .ExecuteNonQuery()
          End With

          'Insert new or modified records
          With insertCommand
            'Create the parameters we will use for the insert command
            AddSqlParameters(dt, .Parameters)

            'Insert all the datarows preserving ID where it is set
            For Each dr As DataRow In dt.Rows
              If dr.RowState <> DataRowState.Deleted Then
                'Set values of all parameters from the data row
                For Each dc As DataColumn In dt.Columns
                  insertCommand.Parameters("@" & dc.ColumnName).Value = dr(dc.ColumnName)
                Next
                'Make the insert string this will be different if ID is null (new record)
                If dr.IsNull("ID") Then
                  BuildSqlInsertString(dt, sql)
                  sql &= "; SELECT CAST(scope_identity() AS int);"
                  .CommandText = sql
                  dr("ID") = Integer.Parse(.ExecuteScalar().ToString)
                Else
                  MakeSqlInsertStringWithID(dt, sql)
                  sql = "SET IDENTITY_INSERT " & dt.TableName & " ON; " & sql & "; SET IDENTITY_INSERT " & dt.TableName & " OFF;"
                  .CommandText = sql
                  .ExecuteNonQuery()
                End If
              End If
            Next
          End With

          'Commit the changes 
          transaction.Commit()
          Return True
        Catch ex As Exception
          Log.LogError(ex, sql)
        End Try
        transaction.Rollback()

      Catch ex As Exception
        Log.LogError(ex)
      End Try
      Return False
    End Function

    <System.Diagnostics.DebuggerStepThrough()>
    Public Sub BuildSqlInsertString(ByVal dt As DataTable, ByRef commandText As String)
      'Make a Sql insert (using sql parameters) to insert a row into the database
      '  assumes ID is an AutoIncrement field and will be filled in by the database
      Try
        If dt Is Nothing Then Exit Sub

        Dim columns As String = "", values As String = ""
        For Each dc As DataColumn In dt.Columns
          If dc.ColumnName <> "ID" Then
            If columns.Length = 0 Then
              columns = "[" & dc.ColumnName.ToString & "]"
            Else
              columns &= ",[" & dc.ColumnName.ToString & "]"
            End If
            If values.Length = 0 Then
              values = "@" & dc.ColumnName.ToString
            Else
              values &= ",@" & dc.ColumnName.ToString
            End If
          End If
        Next
        commandText = "INSERT INTO " & dt.TableName & " (" & columns & ") VALUES(" & values & ")"
      Catch ex As Exception
        Debug.Print("MakeInsertString: " & dt.TableName)
        Debug.Print(ex.Message)
        Log.LogError(ex)
      End Try
    End Sub

    <System.Diagnostics.DebuggerStepThrough()>
    Public Sub MakeSqlInsertStringWithID(ByVal dt As DataTable, ByRef commandText As String)
      'Make a Sql insert (using sql parameters) to insert a row into the database
      '  assumes ID is an AutoIncrement field but trys to insert the existing ID value 
      '  this can be used to delete and re-insert records with the original ID values
      Try
        If dt Is Nothing Then Exit Sub

        Dim columns As String = "", values As String = ""
        For Each dc As DataColumn In dt.Columns
          If columns.Length = 0 Then
            columns = "[" & dc.ColumnName.ToString & "]"
          Else
            columns &= ",[" & dc.ColumnName.ToString & "]"
          End If
          If values.Length = 0 Then
            values = "@" & dc.ColumnName.ToString
          Else
            values &= ",@" & dc.ColumnName.ToString
          End If
        Next
        commandText = "INSERT INTO " & dt.TableName & " (" & columns & ") VALUES(" & values & ")"
      Catch ex As Exception
        Debug.Print("MakeInsertString: " & dt.TableName)
        Debug.Print(ex.Message)
        Log.LogError(ex)
      End Try
    End Sub

    <System.Diagnostics.DebuggerStepThrough()>
    Public Function SqlInsert(ByVal connectionString As String, ByVal insertString As String) As Integer
      Try
        Dim rows As Integer
        Dim da As New SqlClient.SqlDataAdapter
        da.InsertCommand = New SqlClient.SqlCommand
        With da.InsertCommand
          .CommandText = insertString
          .Connection = New SqlClient.SqlConnection(connectionString)
          .Connection.Open() : rows = .ExecuteNonQuery() : .Connection.Close()
        End With

        Return rows
      Catch ex As Exception
        Debug.Print("SqlUpdate: " & insertString)
        Debug.Print(ex.Message)
        Log.LogError(ex, insertString)
      End Try
      Return -1
    End Function
#End Region

#Region " Utilities "

    <System.Diagnostics.DebuggerStepThrough()>
    Public Sub AddSqlParameters(ByVal dt As DataTable, ByRef pc As SqlClient.SqlParameterCollection)
      Try
        If dt Is Nothing Then Exit Sub

        pc.Clear()
        For Each dc As DataColumn In dt.Columns
          'Create a parameter for every column in the table
          Dim p As New SqlClient.SqlParameter()
          p.ParameterName = "@" & dc.ColumnName
          If dc.DataType.Name = "Int16" Then p.SqlDbType = SqlDbType.SmallInt
          If dc.DataType.Name = "Int32" Then p.SqlDbType = SqlDbType.Int
          If dc.DataType.Name = "String" Then p.SqlDbType = SqlDbType.NVarChar
          If dc.DataType.Name = "DateTime" Then p.SqlDbType = SqlDbType.DateTime
          If dc.DataType.Name = "Double" Then p.SqlDbType = SqlDbType.Float
          If dc.DataType.Name = "Decimal" Then p.SqlDbType = SqlDbType.Decimal
          pc.Add(p)
        Next
      Catch ex As Exception
        Debug.Print("MakeParameters: " & dt.TableName)
        Debug.Print(ex.Message)
        Log.LogError(ex)
      End Try
    End Sub

    <System.Diagnostics.DebuggerStepThrough()>
    Public Function SqlString(ByVal value As String) As String
      Try
        If value.Length = 0 Then Return "'Null'"

        'Wrap single quotes round string or return "Null" if string is empty
        Dim returnString As String = ""

        'Look for single quotes (') and double up
        For i As Integer = 0 To value.Length - 1
          Select Case value.Substring(i, 1)
            Case "'" : returnString &= value.Substring(i, 1) & "'"
            Case Else : returnString &= value.Substring(i, 1)
          End Select
        Next i
        Return "'" & returnString & "'"
      Catch ex As Exception
        Return "Null"
      End Try
    End Function

    <System.Diagnostics.DebuggerStepThrough()>
    Public Function SqlDateString(ByVal value As String) As String
      Try
        If value = Nothing Then
          Return "'Null'"
        Else
          Dim tryDate As Date
          If Date.TryParse(value, tryDate) Then
            Return SqlDateString(tryDate)
          Else
            Return "'Null'"
          End If
        End If
      Catch ex As Exception
        Log.LogError(ex)
      End Try
      Return "'Null'"
    End Function

    <System.Diagnostics.DebuggerStepThrough()>
    Public Function SqlDateString(ByVal value As Date) As String
      Try
        If value = Nothing Then
          Return "'Null'"
        Else
          Return SqlString(value.ToString(InvariantCulture))
        End If
      Catch ex As Exception
        Log.LogError(ex)
      End Try
      Return "'Null'"
    End Function

    Public Function SqlDate(dateToConvert As Date) As String
      Return SqlString(dateToConvert.ToString("yyyy-MM-dd HH:mm:ss"))
    End Function

    <System.Diagnostics.DebuggerStepThrough()>
    Public Function EmptyStringToNull(ByVal dr As DataRow, ByVal column As String) As String
      Try
        If dr.IsNull(column) Then Return "'Null'"
        If dr(column).ToString.Length <= 0 Then Return "'Null'"
        Return SqlString(dr(column).ToString)
      Catch ex As Exception
        Log.LogError(ex)
      End Try
      Return "'Null'"
    End Function

    <System.Diagnostics.DebuggerStepThrough()>
    Public Function NullToNothing(ByVal value As String) As String
      Try
        'Check to see if it's null
        If String.IsNullOrEmpty(value) Then Return Nothing

        'String isn't null so just return the string
        Return value
      Catch ex As Exception
        'No code
      End Try
      Return Nothing
    End Function

    Public Function NullToZeroInteger(ByVal value As Object) As Integer
      Try
        'Make sure we have something
        If value Is Nothing Then Return 0

        'If this is an integer then cast it and return it
        If TypeOf value Is Integer Then Return DirectCast(value, Integer)

        'See if we can parse it
        Dim tryInteger As Integer
        If Integer.TryParse(value.ToString, tryInteger) Then Return tryInteger
      Catch ex As Exception
        ' Do nothing
      End Try
      Return 0
    End Function

    Public Function NullToZeroDouble(ByVal value As Object) As Double
      Try
        'Make sure we have something
        If value Is Nothing Then Return 0

        'If this is a double then cast it and return it
        If TypeOf value Is Double Then Return DirectCast(value, Double)

        'See if we can parse it
        Dim tryDouble As Double
        If Double.TryParse(value.ToString, tryDouble) Then Return tryDouble
      Catch ex As Exception
        ' Do nothing
      End Try
      Return 0
    End Function

#End Region

  End Module

End Namespace
