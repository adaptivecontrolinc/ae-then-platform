Partial Public NotInheritable Class Settings
  ' Version 2022-05-09
  '   Removed FillHotEnabled property - reconfigure as parameters

  Public Shared ReadOnly Property DefaultCulture() As System.Globalization.CultureInfo
    Get
      Return System.Globalization.CultureInfo.InvariantCulture
    End Get
  End Property

  Public Shared Property ConnectionStringBDC As String = "data source=adaptive-server;initial catalog=BatchDyeingCentral;user id=Adaptive;password=Control"
  Public Shared Property Demo As Integer = 0
  Public Shared Property DisableMimicButtons As Integer = 0
  Public Shared Property FillHotEnabled As Integer = 0
  Public Shared Property HighTempDrainEnabled As Integer = 0
  Public Shared Property Tank1HeatEnabled As Integer = 0

  Public Shared Property SettingsLoaded As Boolean

  Public Shared Sub Load()
    Try
      SettingsLoaded = True

      'Get application and file path
      Dim appPath As String = My.Application.Info.DirectoryPath
      Dim filePath As String = appPath & "\Settings.xml"

      'If the file doses not exist then just use defaults...
      If Not My.Computer.FileSystem.FileExists(filePath) Then
        ' Save a default copy
        Save()
        Exit Sub
      End If

      'Read the settings into a dataset
      Dim dsSettings As New System.Data.DataSet
      dsSettings.ReadXml(filePath)

      With dsSettings
        If .Tables.Contains("settings") Then
          With .Tables("settings")
            For Each dr As System.Data.DataRow In .Rows
              Select Case dr("name").ToString.ToLower

                Case "ConnectionStringBDC".ToLower
                  If Not dr.IsNull("value") Then ConnectionStringBDC = dr("value").ToString

                Case "Demo".ToLower
                  If Not dr.IsNull("value") Then Demo = Integer.Parse(dr("value").ToString)

                Case "DisableMimicButtons".ToLower
                  If Not dr.IsNull("value") Then DisableMimicButtons = Integer.Parse(dr("value").ToString)

                Case "HighTempDrainEnabled".ToLower
                  If Not dr.IsNull("value") Then HighTempDrainEnabled = Integer.Parse(dr("value").ToString)

                Case "Tank1HeatEnabled".ToLower
                  If Not dr.IsNull("value") Then Tank1HeatEnabled = Integer.Parse(dr("value").ToString)

              End Select
            Next
          End With
        End If
      End With

      ' Save a copy now
      Save()

    Catch ex As Exception
      'TODO - log error ?
    End Try
  End Sub

  Public Shared Sub Save()
    Try
      'Create a settings dataset
      Dim dsSettings As New System.Data.DataSet("Root")

      'Create a settings table
      Dim dtSettings As New System.Data.DataTable("Settings")
      dtSettings.Columns.Add("Name", Type.GetType("System.String"))
      dtSettings.Columns.Add("Help", Type.GetType("System.String"))
      dtSettings.Columns.Add("Value", Type.GetType("System.String"))

      'Create rows
      Dim newRow As System.Data.DataRow

      'ConnectionString to Central database
      newRow = dtSettings.NewRow
      newRow("Name") = "ConnectionStringBDC" : newRow("Help") = "Central Database Connection" : newRow("Value") = ConnectionStringBDC
      dtSettings.Rows.Add(newRow)

      ' Demo Mode
      newRow = dtSettings.NewRow
      newRow("Name") = "Demo" : newRow("Help") = "0=Off, 1=Demo" : newRow("Value") = Demo.ToString
      dtSettings.Rows.Add(newRow)

      ' Disable Mimic Pushbuttons
      newRow = dtSettings.NewRow
      newRow("Name") = "DisableMimicButtons" : newRow("Help") = "0=Off, 1=Disabled" : newRow("Value") = DisableMimicButtons.ToString
      dtSettings.Rows.Add(newRow)

      ' Fill Hot Enabled
      newRow = dtSettings.NewRow
      newRow("Name") = "FillHotEnabled" : newRow("Help") = "0=Disabled, 1=Enabled" : newRow("Value") = FillHotEnabled.ToString
      dtSettings.Rows.Add(newRow)

      ' High Temp Drain Enabled
      newRow = dtSettings.NewRow
      newRow("Name") = "HighTempDrainEnabled" : newRow("Help") = "0=Disabled, 1=Enabled" : newRow("Value") = HighTempDrainEnabled.ToString
      dtSettings.Rows.Add(newRow)

      ' Tank 1 Heat Enabled
      newRow = dtSettings.NewRow
      newRow("Name") = "Tank1HeatEnabled" : newRow("Help") = "0=Disabled, 1=Enabled" : newRow("Value") = Tank1HeatEnabled.ToString
      dtSettings.Rows.Add(newRow)


      'Add the DataTable to the DataSet
      dsSettings.Tables.Add(dtSettings)

      'Get application and file path
      Dim appPath As String = My.Application.Info.DirectoryPath
      Dim filePath As String = appPath & "\Settings.xml"

      'Save this DataSet
      dsSettings.WriteXml(filePath)

    Catch ex As Exception
      Utilities.Log.LogError(ex)
    End Try
  End Sub
End Class
