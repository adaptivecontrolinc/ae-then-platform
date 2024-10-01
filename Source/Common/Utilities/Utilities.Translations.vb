Imports System.Globalization
Imports System.Xml

' [2016-03-09] DT's better version of the Translations Utility.
'              If using ML's version, loops occur that can slow controlcode runs/sec from 2k to 300 (laptop testing showed 4k to 400)
'
Namespace Utilities

  Public Module Translations
    Private trans_ As Dictionary(Of String, Dictionary(Of String, String))   ' es -> (eng->foreign)

    Public Property CultureName() As String
      Get
        Return cultureName_
      End Get
      Set(ByVal value As String)
        cultureName_ = value
        cultureTwoLetterName_ = value.Substring(0, 2)
        TranslateClearCache()
      End Set
    End Property

    Public ReadOnly Property CultureTwoLetterName() As String
      Get
        Return cultureTwoLetterName_
      End Get
    End Property

    Private cultureName_, cultureTwoLetterName_ As String
    Sub New()
      ' Look for a translation 
      Dim ci = CurrentCulture
      cultureName_ = ci.Name ' e.g. zh-TW 
      cultureTwoLetterName_ = ci.TwoLetterISOLanguageName  ' e.g. zh
    End Sub

    Public ReadOnly Property CurrentCultureIsSortable() As Boolean
      Get
        Return (cultureTwoLetterName_ <> "zh")  ' don't sort Chinese
      End Get
    End Property

    Public Sub Translate(ByVal control As Windows.Forms.Control)
      control.Text = Translate(control.Text)
      For Each c As Windows.Forms.Control In control.Controls
        Dim typ = c.GetType
        ' Only translate some UI controls, because setting the Text property is not
        ' a good idea for all
        If typ Is GetType(Windows.Forms.Label) OrElse typ Is GetType(Windows.Forms.CheckBox) OrElse typ Is GetType(Windows.Forms.GroupBox) _
           OrElse typ Is GetType(Windows.Forms.Button) OrElse typ Is GetType(Windows.Forms.RadioButton) Then
          Translate(c)
        ElseIf typ Is GetType(Windows.Forms.TabControl) Then
          For Each tp As Windows.Forms.TabPage In DirectCast(c, Windows.Forms.TabControl).TabPages
            Translate(tp)
          Next tp
        ElseIf typ Is GetType(Windows.Forms.ToolStrip) Then
          For Each tp As ToolStripItem In DirectCast(c, Windows.Forms.ToolStrip).Items
            tp.Text = Translate(tp.Text)
          Next tp
        End If
      Next c
    End Sub

    Public Sub TranslateClearCache()
      trans_ = Nothing
    End Sub

    Public Function Translate(ByVal str As String) As String
      ' Make sure we have something to Translate
      If str = "" Then Return ""

      ' Load the translations from an embedded resource file
      If trans_ Is Nothing Then
        trans_ = New Dictionary(Of String, Dictionary(Of String, String))
        Dim doc = New XmlDocument
        ' Load from a file in the application directory
        doc.Load(System.IO.Path.Combine(Application.StartupPath, "Translations.xml"))
        ' Load up all translations in there
        Dim documentElement = doc.DocumentElement
        If documentElement IsNot Nothing Then
          For Each xn As XmlNode In documentElement.ChildNodes
            With xn.ChildNodes
              If .Count >= 2 Then  ' must be something interesting for us to bother
                Dim english = .Item(0).InnerText
                For i = 1 To .Count - 1
                  Dim cn = .Item(i)
                  Dim language = cn.Name
                  ' Only interested in these two
                  If language = cultureName_ OrElse language = cultureTwoLetterName_ Then
                    Dim dict As Dictionary(Of String, String) = Nothing
                    If Not trans_.TryGetValue(language, dict) Then
                      If language <> "zh" Then  ' should exist as a culture, but doesn't, so avoid an exception
                        Try
                          dict = New Dictionary(Of String, String)(StringComparer.Create(
                                        Globalization.CultureInfo.GetCultureInfo(language), True))
                        Catch : End Try
                      End If
                      If dict Is Nothing Then dict = New Dictionary(Of String, String)
                      trans_.Add(language, dict)
                    End If
                    ' Test that the translation is not already there (throws an exception otherwise)
                    If Not dict.ContainsKey(english) Then
                      dict.Add(english, cn.InnerText)
                    End If
                  End If
                Next i
              End If
            End With
          Next xn
        End If
      End If

      ' Make sure string is set to something
      If str = "" Then Return ""

      ' Look for a translation 
      Dim ret = str
      Dim worte As String = Nothing
      str = str.Replace("&", "")
      Dim endsWith3Dots = str.EndsWith("...", StringComparison.Ordinal)
      If endsWith3Dots Then str = str.Substring(0, str.Length - 3)
      Dim endsWithColon = str.EndsWith(":", StringComparison.Ordinal)
      If endsWithColon Then str = str.Substring(0, str.Length - 1)

      ' Try to match the whole e.g. zh-TW first, or failing that just fall back to the e.g. zh part
      Dim dict1 As Dictionary(Of String, String)
      If trans_.TryGetValue(cultureName_, dict1) Then dict1.TryGetValue(str, worte)
      If String.IsNullOrEmpty(worte) Then
        If trans_.TryGetValue(cultureTwoLetterName_, dict1) Then dict1.TryGetValue(str, worte)
      End If

      If Not String.IsNullOrEmpty(worte) Then
        ret = worte
        If endsWith3Dots Then ret &= "..."
        If endsWithColon Then ret &= ":"
      Else
        ' Dictionary didn't find anything matching
        ' TODO - maybe add a bit of logic to update the Translation.xml with the missing text


      End If
      Return ret
    End Function

#If 0 Then
    ' TODO: LOOK FOR A METHOD TO SAVE THE ENTIRE DICTIONARY ONCE FINISHED



    ' FOLLOWING IS EXAMPLE FROM CHECKWHEIGH CODE - USES DATATABLE, NOT DICTIONARY
  Public Shared Sub Save()
    Try

      'Create a settings dataset
      Dim dsSettings As New System.Data.DataSet("Root")

      'Create a settings table
      Dim dtSettings As New System.Data.DataTable("Settings")
      dtSettings.Columns.Add("Name", Type.GetType("System.String"))
      dtSettings.Columns.Add("Value", Type.GetType("System.String"))

      'Create rows
      Dim newRow As System.Data.DataRow

      Dim seperator As String = "  **********  "

      '******************************************************************************************************
      '****************                                  VALUES                              ****************
      '******************************************************************************************************
      dtSettings.Rows.Add(seperator & "SEPARATOR" & seperator)

      newRow = dtSettings.NewRow
      newRow("Name") = "Balance1Enable" : newRow("Value") = Balance1Enable
      dtSettings.Rows.Add(newRow)





      'Add the DataTable to the DataSet
      dsSettings.Tables.Add(dtSettings)

      'Get application and file path
      Dim appPath As String = My.Application.Info.DirectoryPath
      Dim filePath As String = appPath & "\Settings.xml"

      'Save this DataSet
      dsSettings.WriteXml(filePath)

    Catch ex As Exception
      Dim ErrorText As String = "Settings.Save(): " & ex.message
      Utilities.Log.LogError(ErrorText)
      MsgBox(ErrorText, MsgBoxStyle.Critical, "Adaptive Control")
    End Try
  End Sub

#End If

  End Module

End Namespace

