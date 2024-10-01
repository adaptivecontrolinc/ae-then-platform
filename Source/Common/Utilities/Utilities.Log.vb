Namespace Utilities.Log
  Public Module Log

    Public Sub LogError(ByVal message As String)
      Try
        Dim applicationPath As String = My.Application.Info.DirectoryPath
        Dim fileLocation As String = applicationPath & "\Log.txt"

        WriteToFile(fileLocation, message)
      Catch Ex As Exception
        'Some code
      End Try
    End Sub

    Public Sub LogError(ByVal message As String, ByVal detail As String)
      Try
        Dim applicationPath As String = My.Application.Info.DirectoryPath
        Dim fileLocation As String = applicationPath & "\Log.txt"

        WriteToFile(fileLocation, message, detail)
      Catch Ex As Exception
        'Some code
      End Try
    End Sub

    Public Sub LogError(ByVal err As Exception)
      Try
        Dim applicationPath As String = My.Application.Info.DirectoryPath
        Dim fileLocation As String = applicationPath & "\Log.txt"

        WriteToFile(fileLocation, err.Message, err.StackTrace)
      Catch Ex As Exception
        'Some code
      End Try
    End Sub

    Public Sub LogError(ByVal err As Exception, ByVal sql As String)
      Try
        Dim applicationPath As String = My.Application.Info.DirectoryPath
        Dim fileLocation As String = applicationPath & "\Log.txt"

        WriteToFile(fileLocation, err.Message, err.StackTrace)
      Catch Ex As Exception
        'Some code
      End Try
    End Sub

    Public Sub LogEvent(ByVal message As String)
      Try
        Dim applicationPath As String = My.Application.Info.DirectoryPath
        Dim fileLocation As String = applicationPath & "\Log.txt"

        WriteToFile(fileLocation, message)
      Catch ex As Exception
        'Some code
      End Try
    End Sub


    Public Sub LogFtpError(ByVal err As Exception)
      Try
        Dim applicationPath As String = My.Application.Info.DirectoryPath
        Dim fileLocation As String = applicationPath & "\LogFtp.txt"

        WriteToFile(fileLocation, err.Message, err.StackTrace)
      Catch Ex As Exception
        'Some code
      End Try
    End Sub


    Public Sub LogFtpEvent(ByVal message As String)
      Try
        Dim applicationPath As String = My.Application.Info.DirectoryPath
        Dim fileLocation As String = applicationPath & "\LogFtp.txt"

        System.IO.File.AppendAllText(fileLocation, message)
      Catch ex As Exception
        'Some code
      End Try
    End Sub


    Private Sub WriteToFile(ByVal fileLocation As String, ByVal message As String)
      Try
        Using sw As New System.IO.StreamWriter(fileLocation, True)
          sw.WriteLine()
          sw.WriteLine()
          sw.WriteLine(Date.Now.ToString("yyyy-MM-dd hh:mm:ss"))
          sw.WriteLine("-------------------")
          sw.WriteLine(message)
        End Using
      Catch ex As Exception
        'Some code
      End Try
    End Sub

    Private Sub WriteToFile(ByVal fileLocation As String, ByVal message As String, ByVal detail As String)
      Try
        Using sw As New System.IO.StreamWriter(fileLocation, True)
          sw.WriteLine()
          sw.WriteLine()
          sw.WriteLine(Date.Now.ToString("yyyy-MM-dd hh:mm:ss"))
          sw.WriteLine("-------------------")
          sw.WriteLine(message)
          sw.WriteLine(detail)
        End Using
      Catch ex As Exception
        'Some code
      End Try
    End Sub

    Private Sub WriteToFile(ByVal fileLocation As String, ByVal message As String, ByVal detail As String, ByVal sql As String)
      Try
        Using sw As New System.IO.StreamWriter(fileLocation, True)
          sw.WriteLine()
          sw.WriteLine()
          sw.WriteLine(Date.Now.ToString("yyyy-MM-dd hh:mm:ss"))
          sw.WriteLine("-------------------")
          sw.WriteLine(message)
          sw.WriteLine(detail)
          sw.WriteLine(sql)
        End Using
      Catch ex As Exception
        'Some code
      End Try
    End Sub

  End Module
End Namespace