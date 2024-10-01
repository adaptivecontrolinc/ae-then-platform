Public Class DbCentral
  Private ReadOnly connectionString_ As String
  Public Sub New(ByVal connectionString As String)
    connectionString_ = connectionString
  End Sub

  Public Function DbGetDataTable(ByVal sql As String) As DataTable
    Using dbConnection As New SqlClient.SqlConnection(connectionString_)
      dbConnection.Open()
      Using da As New SqlClient.SqlDataAdapter(sql, dbConnection)
        With da
          Try
            Dim ret As New DataTable : ret.Locale = InvariantCulture
            .Fill(ret)
            Return ret
          Catch ex As Exception
            With ex.Data : .Add("ConnectionString", connectionString_) : .Add("CommandText", sql) : End With
            Throw
          End Try
        End With
      End Using
    End Using
  End Function

  Public Function DbExecute(ByVal sql As String) As Integer
    Using dbConnection As New SqlClient.SqlConnection(connectionString_)
      dbConnection.Open()
      With dbConnection.CreateCommand
        .CommandText = sql : .CommandTimeout = 0
        Try
          Return .ExecuteNonQuery()
        Catch ex As Exception
          With ex.Data : .Add("ConnectionString", connectionString_) : .Add("CommandText", sql) : End With
          Throw
        End Try
      End With
    End Using
  End Function


End Class

' --------------------------------------------------------
Public NotInheritable Class SqlConvert
  Public Shared Function ToSqlString(ByVal value As DateTime) As String
    Dim ret As String = String.Format(InvariantCulture, _
                         "{0:0}-{1:0}-{2:0} {3:00}:{4:00}:{5:00}", _
                         New Object() {value.Year, value.Month, value.Day, value.Hour, value.Minute, value.Second})
    Return "'" & ret & "'"
  End Function

  Public Shared Function ToSqlString(ByVal value As String) As String
    If value Is Nothing Then Return "Null"

    With New System.Text.StringBuilder
      .Append("'"c)
      value = value.Replace("’"c, "'"c)  ' some strange other sort of quote
      value = value.Replace("'", "''")
      .Append(value)
      .Append("'"c)
      Return .ToString
    End With
  End Function

  Public Shared Function ToSqlString(ByVal value As Boolean) As String
    If value Then Return "1"
    Return "Null"
  End Function

  Public Shared Function ToSqlString(ByVal value As Short) As String
    Return value.ToString(InvariantCulture)
  End Function

  Public Shared Function ToSqlString(ByVal value As Integer) As String
    Return value.ToString(InvariantCulture)
  End Function

  Public Shared Function ToSqlString(ByVal value As Long) As String
    Return value.ToString(InvariantCulture)
  End Function

  Public Shared Function ToSqlString(ByVal value As Single) As String
    Return value.ToString(InvariantCulture)
  End Function

  Public Shared Function ToSqlString(ByVal value As Double) As String
    Return value.ToString(InvariantCulture)
  End Function

  Public Shared Function ToSqlString(ByVal value As Byte()) As String
    With New System.Text.StringBuilder
      .Append("0X")
      For i As Integer = 0 To value.Length - 1
        .Append(value(i).ToString("X2", InvariantCulture))
      Next i
      Return .ToString
    End With
  End Function

  Public Shared Function ToSqlString(ByVal value As Object) As String
    If TypeOf value Is DateTime Then Return ToSqlString(DirectCast(value, DateTime))
    If value Is Nothing OrElse TypeOf value Is DBNull Then Return "Null"
    Dim valueAsString As String = TryCast(value, String)
    If valueAsString IsNot Nothing Then Return ToSqlString(valueAsString)
    If TypeOf value Is Boolean Then Return ToSqlString(DirectCast(value, Boolean))
    If TypeOf value Is Short Then Return ToSqlString(DirectCast(value, Short))
    If TypeOf value Is Integer Then Return ToSqlString(DirectCast(value, Integer))
    If TypeOf value Is Long Then Return ToSqlString(DirectCast(value, Long))
    If TypeOf value Is Single Then Return ToSqlString(DirectCast(value, Single))
    If TypeOf value Is Double Then Return ToSqlString(DirectCast(value, Double))
    If TypeOf value Is Byte() Then Return ToSqlString(DirectCast(value, Byte()))
    Return value.ToString  ' this is failure, really
  End Function
End Class

' --------------------------------------------------------
Public NotInheritable Class Null
  Private Sub New()
  End Sub

  Public Shared Function NullToEmptyString(ByVal value As Object) As String
    If value Is Nothing OrElse TypeOf value Is DBNull Then Return String.Empty
    Dim str As String = TryCast(value, String) : If str IsNot Nothing Then Return str
    Return CType(value, String)
  End Function

  Public Shared Function NullToZeroInteger(ByVal value As Object) As Integer
    If value Is Nothing OrElse TypeOf value Is DBNull Then Return 0
    If value.ToString = "" Then Return 0
    If TypeOf value Is Integer Then Return DirectCast(value, Integer)
    Return CType(value, Integer)
  End Function

  Public Shared Function NullToZeroDouble(ByVal value As Object) As Double
    If value Is Nothing OrElse TypeOf value Is DBNull Then Return 0
    If TypeOf value Is Double Then Return DirectCast(value, Double)
    Return CType(value, Double)
  End Function

  Public Shared Function NullToZeroDate(ByVal value As Object) As Date
    If value Is Nothing OrElse TypeOf value Is DBNull Then Return Date.MinValue
    If TypeOf value Is Date Then Return DirectCast(value, Date)
    ' Sometimes we get a number which is actually an Ole Automation date
    Return Date.FromOADate(CType(value, Double))
  End Function

  Public Shared Function NullToFalse(ByVal value As Object) As Boolean
    If value Is Nothing OrElse TypeOf value Is DBNull Then Return False
    If TypeOf value Is Boolean Then Return DirectCast(value, Boolean)
    Return CType(value, Boolean)
  End Function

  Public Shared Function NullToNothingByteArray(ByVal value As Object) As Byte()
    If value Is Nothing OrElse TypeOf value Is DBNull Then Return Nothing
    Return DirectCast(value, Byte())
  End Function

End Class
