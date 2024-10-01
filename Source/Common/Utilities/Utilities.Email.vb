Partial Public Class Utilities
  Public Class Email

    Private Shared _smtpHost As String = "mail.bellsouth.net"
    Public Shared Property smtpHost() As String
      Get
        Return _smtpHost
      End Get
      Set(ByVal value As String)
        _smtpHost = value
      End Set
    End Property

    Private Shared _loginUserName As String = "PspAutoclave@hotmail.com"
    Public Shared Property LogingUserName() As String
      Get
        Return _loginUserName
      End Get
      Set(ByVal value As String)
        _loginUserName = value
      End Set
    End Property

    Private Shared _loginPassword As String = "PspMarketing"
    Public Shared Property LoginPassword() As String
      Get
        Return _loginPassword
      End Get
      Set(ByVal value As String)
        _loginPassword = value
      End Set
    End Property

    Private Shared _sender As String = "adaptive@adaptivecontrol.com"
    Public Shared Property Sender() As String
      Get
        Return _sender
      End Get
      Set(ByVal value As String)
        _sender = value
      End Set
    End Property

    Private Shared _recipient1 As String = "acdalehopkins@hotmail.com"
    Public Shared Property Recipient1() As String
      Get
        Return _recipient1
      End Get
      Set(ByVal value As String)
        _recipient1 = value
      End Set
    End Property

    Private Shared _recipient2 As String = ""
    Public Shared Property Recipient2() As String
      Get
        Return _recipient2
      End Get
      Set(ByVal value As String)
        _recipient2 = value
      End Set
    End Property

    Private Shared _recipient3 As String = ""
    Public Shared Property Recipient3() As String
      Get
        Return _recipient3
      End Get
      Set(ByVal value As String)
        _recipient3 = value
      End Set
    End Property

    Private Shared _recipient4 As String = ""
    Public Shared Property Recipient4() As String
      Get
        Return _recipient4
      End Get
      Set(ByVal value As String)
        _recipient4 = value
      End Set
    End Property

    Private Shared _subjectSave As String
    Private Shared _messageSave As String
    Private Shared _fileName As String

    Public Shared Sub LoadParameters(ByVal parent As ACParent)
      Try

        Dim smtpHost As String = parent.Setting("SmtpHost")
        If Not String.IsNullOrEmpty(smtpHost) Then
          _smtpHost = smtpHost
        End If

        Dim loginUser As String = parent.Setting("UserName")
        If Not String.IsNullOrEmpty(loginUser) Then
          _loginUserName = loginUser
        End If

        Dim loginPassword As String = parent.Setting("Password")
        If Not String.IsNullOrEmpty(loginPassword) Then
          _loginPassword = loginPassword
        End If

        Dim sender As String = parent.Setting("EmailSender")
        If Not String.IsNullOrEmpty(sender) Then
          _sender = sender
        End If

        Dim recipient1 As String = parent.Setting("EmailRecipient1")
        If Not String.IsNullOrEmpty(recipient1) Then
          _recipient1 = recipient1
        End If

        Dim recipient2 As String = parent.Setting("EmailRecipient2")
        If Not String.IsNullOrEmpty(recipient2) Then
          _recipient2 = recipient2
        End If

        Dim recipient3 As String = parent.Setting("EmailRecipient3")
        If Not String.IsNullOrEmpty(recipient3) Then
          _recipient3 = recipient3
        End If

        Dim recipient4 As String = parent.Setting("EmailRecipient4")
        If Not String.IsNullOrEmpty(recipient4) Then
          _recipient4 = recipient4
        End If


      Catch ex As Exception

      End Try
    End Sub

    Public Shared Sub SendSync(ByVal subject As String, ByVal message As String)
      Try
        _subjectSave = subject
        _messageSave = message

        SendEmail()

        'Setup a delay timer to simulate email activity
        _timer.Seconds = 60

      Catch ex As Exception
        Utilities.Log.LogError("Utilities.Email.SendSync:  " & ex.Message)
      End Try
    End Sub

    Public Shared Sub SendSync(ByVal subject As String, ByVal message As String, ByVal filename As String)
      Try
        _subjectSave = subject
        _messageSave = message
        _fileName = filename

        SendEmail()

        'Setup a delay timer to simulate email activity
        _timer.Seconds = 60

      Catch ex As Exception
        Utilities.Log.LogError("Utilities.Email.SendSync:  " & ex.Message)
      End Try
    End Sub

    Public Shared Sub SendAsync(ByVal subject As String, ByVal message As String)
      Try
        _subjectSave = subject
        _messageSave = message

        Dim backgroundThread As New System.Threading.Thread(AddressOf SendEmail)
        backgroundThread.Start()

        'Setup a delay timer to simulate email activity
        _timer.Seconds = 60

      Catch ex As Exception
        Utilities.Log.LogError("Utilities.Email.SendAysnc:  " & ex.Message)
      End Try
    End Sub

    Public Shared Sub SendAsync(ByVal subject As String, ByVal message As String, ByVal filename As String)
      Try
        _subjectSave = subject
        _messageSave = message
        _fileName = filename

        Dim backgroundThread As New System.Threading.Thread(AddressOf SendEmail)
        backgroundThread.Start()

        'Setup a delay timer to simulate email activity
        _timer.Seconds = 60

      Catch ex As Exception
        Utilities.Log.LogError("Utilities.Email.SendAysnc:  " & ex.Message)
      End Try
    End Sub

    Private Shared Sub SendEmail()
      Try

        'Create the Mail Message
        Dim email As New System.Net.Mail.MailMessage
        With email
          .From = New System.Net.Mail.MailAddress(Utilities.Email.Sender)
          .To.Add(Recipient1)
          If Recipient2 <> "" Then .To.Add(Recipient2)
          If Recipient3 <> "" Then .To.Add(Recipient3)
          If Recipient4 <> "" Then .To.Add(Recipient4)
          .Subject = _subjectSave
          .Body = _messageSave

          If _fileName <> "" AndAlso (Microsoft.VisualBasic.FileIO.FileSystem.FileExists(_fileName)) Then
            .Attachments.Add(New System.Net.Mail.Attachment(_fileName))
          End If

        End With

        'Create an object to be accessed in callback method
        Dim userState As Object = email

        'Send the message
        Dim smtp As New System.Net.Mail.SmtpClient
        Dim basicAuthenticationInfo As New System.Net.NetworkCredential(_loginUserName, _loginPassword) '("username", "password")

        With smtp
          .Host = Utilities.Email.smtpHost
          .UseDefaultCredentials = False
          .Credentials = basicAuthenticationInfo
          .EnableSsl = True

          'Wire Up the Event for the the Async Send is completed
          'AddHandler smtp.SendCompleted, AddressOf SmtpClient_OnCompleted

          .Send(email)

          'Setup a delay timer to simulate email activity
          _timer.Seconds = 60

        End With

        'Dispose of the Email Object 
        email.Dispose()


      Catch ex As Exception
        Utilities.Log.LogError("Utilities.Email.SendEmail:  " & ex.Message)
      End Try

    End Sub

    Public Shared Sub SmtpClient_OnCompleted(ByVal sender As Object, ByVal e As AsyncCompletedEventArgs)
      Try
        'Get the Original MailMessage Object
        Dim mail As System.Net.Mail.MailMessage = CType(e.UserState, System.Net.Mail.MailMessage)

        'Write out the subject
        Dim subject As String = mail.Subject

        If e.Cancelled Then
          _emailComplete = False
          _emailSentFailed = True
          'Console.WriteLine("Send canceled for mail with subject [{0}].", subject)
        End If
        If Not (e.Error Is Nothing) Then
          _emailComplete = False
          _emailSentFailed = True
          'Console.WriteLine("Error {1} occurred when sending mail [{0}]",subject,e.Error.ToString())
        Else
          _emailComplete = True
          _emailSentFailed = False
          'Console.WriteLine("Message [{0}] send.", subject)
        End If

      Catch ex As Exception

      End Try
    End Sub

    Private Shared _emailComplete As Boolean
    Public Shared ReadOnly Property EmailSentComplete() As Boolean
      Get
        Return _emailComplete
      End Get
    End Property

    Private Shared _emailSentFailed As Boolean
    Public Shared ReadOnly Property EmailSentFailed() As Boolean
      Get
        Return _emailSentFailed
      End Get
    End Property

    Private Shared _timer As New Timer
    Public Shared ReadOnly Property Timer() As Timer
      Get
        Return _timer
      End Get
    End Property
    Public Shared ReadOnly Property IsActive() As Boolean
      Get
        Return Not _timer.Finished
      End Get
    End Property

    Private Shared _sentRecently As Boolean
    Public Shared Property SentRecently() As Boolean
      Get
        Return _sentRecently
      End Get
      Set(ByVal value As Boolean)
        _sentRecently = False
      End Set
    End Property

  End Class
End Class
