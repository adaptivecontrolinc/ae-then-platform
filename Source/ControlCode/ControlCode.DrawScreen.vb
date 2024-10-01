Imports Utilities.Translations

Partial Class ControlCode
  ' Up to 15 rows (1-15) on one screen on an 800x600 display
  '   Use \ to get an integer value (no tenths) (round off)
  '   Use / to get a double and follow with .ToString("#0.0") to show one decimal place (doesn't round off)
  <ScreenButton("Main", 1, ButtonImage.Vessel),
  ScreenButton("Data", 2, ButtonImage.Information),
  ScreenButton("Tank 1", 3, ButtonImage.SideVessel),
  ScreenButton("Reserve", 4, ButtonImage.SideVessel),
  ScreenButton("Add", 5, ButtonImage.SideVessel),
  ScreenButton("Temp", 6, ButtonImage.Thermometer),
  ScreenButton("Flow", 7, ButtonImage.Information),
  ScreenButton("Lid", 8, ButtonImage.Information),
  ScreenButton("Dispense", 9, ButtonImage.Information)>
  Public Sub DrawScreen(ByVal screen As Integer, ByVal row() As String) Implements ACControlCode.DrawScreen

    Select Case screen
      Case 1 : DrawScreenMain(row)
      Case 2 : DrawScreenData(row)
      Case 3 : DrawScreenKitchen(row)
      Case 4 : DrawScreenReserve(row)
      Case 5 : DrawScreenAddition(row)
      Case 6 : DrawScreenTemp(row)
      Case 7 : DrawScreenFlow(row)
      Case 8 : DrawScreenLid(row)
      Case 9 : DrawScreenDispense(row)

    End Select
  End Sub

  Sub DrawScreenMain(ByVal row() As String)
    row(1) = "Procedure " & Parent.ProgramNumber.ToString("#0")
    row(2) = SafetyControl.StateString
    row(3) = ""
    row(4) = Translate("Level") & (": ") & (MachineLevel / 10).ToString("#0.0") & "%"
    row(5) = Translate("Volume") & (": ") & (VolumeBasedOnLevel / 10).ToString("#0") & "Gal"
    row(6) = Translate("Pressure") & (": ") & MachinePressureDisplay
    row(7) = PumpControl.StateString '& (" ") & (IO.PumpSpeedOutput / 10).ToString("#0.0") & "%"
    row(8) = PumpControl.StateFlowString
    '    If FR.IsOn Then row(8) = ("FR: ") & FR.Status
    '    If FC.IsOn Then row(8) = ("FC: ") & FC.Status
    If FL.IsOn Then row(9) = FL.Status
    If DP.IsOn Then row(9) = DP.Status
    If Parameters.PackageDiffPressEnable >= 1 Then row(10) = ("DP: ") & Translate("Actual") & (": ") & PackageDpStr
    row(11) = Translate("Flow Rate") & (": ") & (MachineFlowRatePv).ToString("#0") & " GPM"
    row(12) = "          " & (FlowratePerWt).ToString("#0.0") & " GPM/kg"
    If IO.Fill Then
      ' Fill or CityWaterFill
      If FI.IsActive Then
        row(13) = "Desired Fill Temp: " & (FI.FillTemp / 10).ToString("#0.0") & "F"
        row(14) = "Actual Fill Temp: " & (IO.FillTemp / 10).ToString("#0.0") & "F"
        row(15) = "Blend fill valve: " & (IO.BlendFillOutput / 10).ToString("#0.0") & "%"
      End If
      If RC.IsActive Then
        row(13) = "Desired Rinse Temp: " & (RC.RinseTemp / 10).ToString("#0.0") & "F"
        row(14) = "Actual Rinse Temp: " & (IO.FillTemp / 10).ToString("#0.0") & "F"
        row(15) = "Blend fill valve: " & (IO.BlendFillOutput / 10).ToString("#0.0") & "%"
      End If
      If RH.IsActive Then
        row(13) = "Desired Rinse Temp: " & (RH.RinseTemp / 10).ToString("#0.0") & "F"
        row(14) = "Actual Rinse Temp: " & (IO.FillTemp / 10).ToString("#0.0") & "F"
        row(15) = "Blend fill valve: " & (IO.BlendFillOutput / 10).ToString("#0.0") & "%"
      End If
      If RI.IsActive Then
        row(13) = "Desired Rinse Temp: " & (RI.RinseTemp / 10).ToString("#0.0") & "F"
        row(14) = "Actual Rinse Temp: " & (IO.FillTemp / 10).ToString("#0.0") & "F"
        row(15) = "Blend fill valve: " & (IO.BlendFillOutput / 10).ToString("#0.0") & "%"
      End If
    End If

    row(13) = ""
    row(14) = TemperatureControl.Status

    Dim TextRow16 As String
    TextRow16 = ""
    If AP.IsWaitReady Then
      If (AP.Message >= 1) Or (AP.Message <= 99) Then
        TextRow16 = Parent.Message(AP.Message)
      End If
    End If
    If RP.IsWaitReady Then
      If (RP.Message >= 1) Or (RP.Message <= 99) Then
        TextRow16 = Parent.Message(RP.Message)
      End If
    End If
    row(16) = TextRow16

    '    row(17) = Translate("Color") & (": ") & RecipeCode
    '    row(18) = Translate("Name") & (": ") & RecipeName
  End Sub

  Sub DrawScreenData(ByVal row() As String)
    If Parent.IsProgramRunning Then
      row(1) = Parent.ProgramNumberAsString & " - " & Parent.ProgramName
      row(2) = Translate("Cycle Time") & (": ") & CycleTimeDisplay
      row(3) = Translate("Batch weight") & (": ") & BatchWeight.ToString("#0") & " lbs"
      row(4) = Translate("Package Height") & (": ") & PackageHeight.ToString("#0")
      row(5) = Translate("Package Type") & (": ") & PackageType.ToString
      row(6) = Translate("") ' TODO ("Contacts: ") & TotalNumberOfContacts
      If IO.CityWaterSw Then row(7) = "Using City Water"
      If Delay <> DelayValue.NormalRunning Then
        row(8) = Translate("Delay") & (": ") & Delay.ToString
      Else
        row(8) = Translate("Normal Running")
      End If
      row(10) = Translate("Start Time") & (": ")
      row(11) = "    " & StartTime
      row(12) = Translate("End Time") & (": ")
      row(13) = "    " & EndTime

    Else
      row(1) = Translate("Machine Idle") & (": ") & ProgramStoppedTimer.ToString
      row(2) = ""
      row(3) = ""
      row(4) = Translate("Previous Cycle Time") & (": ") & TimerString(LastProgramCycleTime)
    End If

    ' Row 15
    Dim TextRow15 As String
    TextRow15 = ""
    If AP.IsWaitReady Then
      If (AP.Message >= 1) Or (AP.Message <= 99) Then
        TextRow15 = Parent.Message(AP.Message)
      End If
    End If
    If RP.IsWaitReady Then
      If (RP.Message >= 1) Or (RP.Message <= 99) Then
        TextRow15 = Parent.Message(RP.Message)
      End If
    End If
    row(15) = TextRow15
  End Sub

  Sub DrawScreenDispense(ByVal row() As String)
    row(1) = RecipeCode
    row(2) = RecipeName
    row(3) = ""
    DrawRecipeSteps(row, 4, Tank1Calloff)


  End Sub

  Sub DrawRecipeSteps(row() As String, startRow As Integer, dropNumber As Integer)
    Try
      Dim rowNumber = startRow
      If DrugroomPreview.BulkedRecipeHost Is Nothing Then Exit Sub
      If DrugroomPreview.BulkedRecipeHost.Rows.Count <= 0 Then Exit Sub

      For Each recipeRow As DataRow In DrugroomPreview.BulkedRecipeHost.Rows
        Dim stepNumber = Utilities.Sql.NullToZeroInteger(recipeRow("ADD_NUMBER"))
        If stepNumber = dropNumber Then
          Dim stepTypeID = Utilities.Sql.NullToZeroInteger(recipeRow("INGREDIENT_ID"))
          ' Dim stepCode = recipeRow("StepCode").ToString.Trim
          Dim stepDescription = recipeRow("INGREDIENT_DESC").ToString.Trim
          Dim stepUnits = Utilities.Sql.NullToZeroDouble(recipeRow("UNITS")).ToString("#0")
          Dim stepAmount = Utilities.Sql.NullToZeroDouble(recipeRow("AMOUNT")).ToString("#0") & (" ") & stepUnits
          Dim stepDAmount = Utilities.Sql.NullToZeroDouble(recipeRow("DAMOUNT")).ToString("#0.00") & (" ") & stepUnits

          ' If stepTypeID = 1 OrElse stepTypeID = 2 Then
          row(rowNumber) = stepDescription & " " & stepAmount
          rowNumber += 1
          ' End If
        End If
      Next
    Catch ex As Exception
      ' Ignore errors
    End Try
  End Sub

  Sub DrawRecipeStep(row() As String, startRow As Integer, dropNumber As Integer)
    Try
      Dim rowNumber = startRow
      If DrugroomPreview.BulkedRecipeHost Is Nothing Then Exit Sub
      If DrugroomPreview.BulkedRecipeHost.Rows.Count <= 0 Then Exit Sub

      For Each recipeRow As DataRow In DrugroomPreview.BulkedRecipeHost.Rows
        Dim stepNumber = Utilities.Sql.NullToZeroInteger(recipeRow("StepNumber"))
        If stepNumber = dropNumber Then
          Dim stepTypeID = Utilities.Sql.NullToZeroInteger(recipeRow("INGREDIENT_ID"))
          ' Dim stepCode = recipeRow("StepCode").ToString.Trim
          Dim stepDescription = recipeRow("INGREDIENT_DESC").ToString.Trim
          Dim stepUnits = Utilities.Sql.NullToZeroDouble(recipeRow("UNITS")).ToString("#0")
          Dim stepAmount = Utilities.Sql.NullToZeroDouble(recipeRow("AMOUNT")).ToString("#0") & (" ") & stepUnits
          Dim stepDAmount = Utilities.Sql.NullToZeroDouble(recipeRow("DAMOUNT")).ToString("#0.00") & (" ") & stepUnits

          If stepTypeID = 1 OrElse stepTypeID = 2 Then
            row(rowNumber) = stepDescription & " " & stepAmount
            rowNumber += 1
          End If
        End If
      Next
    Catch ex As Exception
      ' Ignore errors
    End Try
  End Sub




  Sub DrawScreenKitchen(ByVal row() As String)
    row(1) = Tank1Status
    ' KP.KP1.Status, LA.KP1.Status, KA.Status, CK.Status
    ' If LA.KP1.IsOn then Row(#) = "Looking ahead to lot & LA.Job

    row(2) = ("   ") & Translate("Level") & (": ") & (Tank1Level / 10).ToString("#0.0") & "%"

    If Parameters.Tank1TempProbe > 0 Then
      row(3) = ("   ") & Translate("Temp:") & (" ") & (IO.Tank1Temp / 10).ToString("#0.0") & "F"
      row(4) = ("   ") & Translate("Calloff") & (": ") & Tank1Calloff.ToString
    Else
      row(3) = ("   ") & Translate("Calloff") & (": ") & Tank1Calloff.ToString
      row(4) = ""
    End If
    row(5) = ""
    If IO.Tank1ManualSw Then row(5) = "Drugroom in Manual"
    If CK.IsOn Then row(6) = CK.Status
    row(7) = ""
    row(8) = RecipeCode
    row(9) = RecipeName
    row(10) = ""
    row(11) = ""
    row(12) = ""
    row(13) = ""

    ' Row 15
    Dim TextRow14 As String
    TextRow14 = ""
    If AP.IsWaitReady Then
      If (AP.Message >= 1) Or (AP.Message <= 99) Then
        TextRow14 = Parent.Message(AP.Message)
      End If
    End If
    If RP.IsWaitReady Then
      If (RP.Message >= 1) Or (RP.Message <= 99) Then
        TextRow14 = Parent.Message(RP.Message)
      End If
    End If
    row(14) = TextRow14
    row(15) = SafetyControl.StateString
  End Sub

  Sub DrawScreenAddition(ByVal row() As String)
    row(1) = Translate("Add Level") & (": ") & (AddLevel / 10).ToString("#0.0") & "%"
    row(2) = AddControl.Status
    If AC.IsBackground Then row(3) = AC.Status
    If AT.IsBackground Then row(3) = AT.Status
    row(4) = ""
    row(5) = ""
    If KP.IsOn AndAlso KP.Tank1Destination = EKitchenDestination.Add Then
      row(6) = KP.KP1.Status ' KP.KP1.Status
    End If
    row(7) = ""
    row(8) = ""
    row(9) = ""
    If AP.IsWaitReady Then
      Dim Case1Text As String
      If (AP.Message >= 1) And (AP.Message <= 99) Then
        Case1Text = Translate("Add Prepare") & (": ") & Parent.Message(AP.Message)
      End If
      If Case1Text IsNot Nothing Then
        row(10) = WordWrap(Case1Text, 10)
      End If
    End If
    row(11) = ""
    row(12) = ""
    row(13) = ""
    row(14) = ""
    Dim TextRow15 As String
    TextRow15 = ""
    If AP.IsWaitReady Then
      If (AP.Message >= 1) Or (AP.Message <= 99) Then
        TextRow15 = Parent.Message(AP.Message)
      End If
    End If
    row(15) = TextRow15
    row(16) = ""
  End Sub

  Sub DrawScreenReserve(ByVal row() As String)
    row(1) = Translate("Reserve Level") & (": ") & (ReserveLevel / 10).ToString("#0.0") & "%"
    row(2) = Translate("Reserve Temp") & (": ") & (IO.ReserveTemp / 10).ToString("#0.0") & "F"
    row(3) = ReserveControl.Status
    If RF.IsOn Then row(4) = RF.Status
    If RD.IsOn Then row(4) = RD.Status
    If RW.IsOn Then row(4) = RW.Status
    If RT.IsOn Then row(5) = RT.Status
    row(6) = ""
    If KP.IsOn AndAlso (KP.Tank1Destination = EKitchenDestination.Reserve) AndAlso Not RT.IsOn Then
      row(7) = KP.KP1.Status
    End If
    row(8) = ""
    row(9) = ""
    If RP.IsWaitReady Then
      Dim Case1Text As String
      If (RP.Message >= 1) And (RP.Message <= 99) Then
        Case1Text = Translate("Reserve Prepare") & (": ") & Parent.Message(RP.Message)
      End If
      If Case1Text IsNot Nothing Then
        row(10) = WordWrap(Case1Text, 10)
      End If
    End If
    row(11) = ""
    row(12) = ""
    row(13) = ""
    row(14) = ""
    row(15) = SafetyControl.StateString
  End Sub

  Sub DrawScreenFlow(ByVal row() As String)
    row(1) = PumpControl.StateString & (" ") & (IO.PumpSpeedOutput / 10).ToString("#0.0") & "%"
    row(2) = PumpControl.StateFlowString
    row(3) = ""
    If FL.IsOn Then row(3) = FL.Status
    If DP.IsOn Then row(3) = DP.Status
    row(4) = ""
    If Parameters.PackageDiffPressEnable >= 1 Then row(4) = ("DP ") & Translate("Actual") & (": ") & PackageDpStr

    row(5) = Translate("Flow Rate") & (": ") & (MachineFlowRatePv).ToString("#0") & " gpm"
    row(6) = "Flow Percent" & (": ") & (FlowRatePercent / 10).ToString("#0.0") & "%"
    row(6) = "  Per Pound: " & (FlowratePerWt).ToString("#0.0") & " gpm/lb"
    row(7) = ""
    row(8) = ""
    If Parent.IsProgramRunning Then
      row(9) = "Flow In-Out: " & PumpControl.CountInToOut.ToString("#0")
      row(10) = "Flow Out-In: " & PumpControl.CountOutToIn.ToString("#0")
    Else
      row(9) = "Previous Flow In-Out: " & PumpControl.CountInToOut.ToString("#0")
      row(10) = "Previous Flow Out-In: " & PumpControl.CountOutToIn.ToString("#0")
    End If
    row(9) = ""
    row(10) = ""
    row(11) = ""
    row(12) = ""
    row(13) = ""
    row(14) = ""
    ' Row 15
    Dim TextRow15 As String
    TextRow15 = ""
    If AP.IsWaitReady Then
      If (AP.Message >= 1) Or (AP.Message <= 99) Then
        TextRow15 = Parent.Message(AP.Message)
      End If
    End If
    If RP.IsWaitReady Then
      If (RP.Message >= 1) Or (RP.Message <= 99) Then
        TextRow15 = Parent.Message(RP.Message)
      End If
    End If
    row(15) = TextRow15
    row(16) = SafetyControl.StateString
  End Sub

  Sub DrawScreenTemp(ByVal row() As String)
    row(1) = TemperatureControl.Status
    row(2) = ("  ") & Translate("Gradient") & (": ") & (TemperatureControl.Gradient / 10).ToString("#0.0") & " F/m"
    row(3) = ("  ") & Translate("Final Temp") & (": ") & (TemperatureControl.FinalTemp / 10).ToString("#0") & " F"
    row(4) = ("  ") & Translate("Setpoint") & (": ") & (TemperatureControl.Pid.PidSetpoint / 10).ToString("#0") & " F"
    row(5) = ("  ") & Translate("Output") & (": ") & (TemperatureControl.IoOutput / 10).ToString("#0") & "%"
    row(6) = ""
    row(7) = ""
    If (IO.CondensateTemp > 320) AndAlso (IO.CondensateTemp < 3000) Then
      row(8) = "Condensate Temp: " & (IO.CondensateTemp / 10).ToString("#0") & " F"
    ElseIf Parameters.TempCondensateLimit > 0 Then
      row(8) = "Condensate Temp: " & (IO.CondensateTemp / 10).ToString("#0") & " F"
    Else : row(8) = ""
    End If
    row(9) = ""
    row(10) = ""
    row(11) = ""
    row(12) = ""
    row(13) = SafetyControl.StateString
    ' Row 14
    Dim TextRow14 As String
    TextRow14 = ""
    If AP.IsWaitReady Then
      If (AP.Message >= 1) Or (AP.Message <= 99) Then
        TextRow14 = Parent.Message(AP.Message)
      End If
    End If
    If RP.IsWaitReady Then
      If (RP.Message >= 1) Or (RP.Message <= 99) Then
        TextRow14 = Parent.Message(RP.Message)
      End If
    End If
    row(14) = TextRow14
  End Sub

  Sub DrawScreenLid(ByVal row() As String)
    If IO.OpenLid OrElse Not IO.KierLidClosed Then
      row(1) = "Lid Open"
    ElseIf IO.CloseLid OrElse IO.KierLidClosed Then
      row(1) = "Lid Closed"
    End If
    If IO.OpenLockingBand Then
      row(2) = "Locking Band Open"
    ElseIf IO.CloseLockingBand Then
      row(2) = "Locking Band Closed"
    End If
    If IO.LockingPinRaised Then
      row(3) = "Locking Pin Raised"
    Else : row(3) = "Locking Pin Lowered"
    End If
    row(4) = LidControl.StateString
    row(5) = ""
    row(6) = ""
    row(7) = ""
    row(8) = ""
    row(9) = ""
    row(10) = ""
    row(11) = ""
    row(12) = ""
    row(13) = ""
    row(14) = ""
    row(15) = ""
  End Sub

  ' TODO
  Public Function GetDispenseInformation(ByVal dispenseinformation As String, ByVal calloff As Integer) As String
    Try
      'These separators are used by programs and prefixed steps
      Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
      Dim separator1 As String = Convert.ToChar(255)
      Dim separator2 As String = ","

      'strings
      Dim steps() As String
      Dim StepDetailSplit() As String
      Dim stepnumbersplit() As String
      Dim NumberOfPrepsFound As Integer = 0
      Dim stepnumber As String = ""
      Dim scaleInfo As String = ""
      'Split the program string into an array of program steps - one step per array element
      steps = dispenseinformation.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
      'Make sure we've have something to check
      If steps.GetUpperBound(0) <= 0 Then Return ""

      For i = 1 To steps.GetUpperBound(0)
        StepDetailSplit = steps(i).Split(CChar(separator2))
        If StepDetailSplit(1) Like "Fnc=Prep*" Then
          NumberOfPrepsFound += 1
          If NumberOfPrepsFound = calloff Then
            stepnumbersplit = StepDetailSplit(0).Split(CChar("="))
            stepnumber = stepnumbersplit(1)
            scaleInfo = StepDetailSplit(3).Substring(StepDetailSplit(3).Length - 5, 5)
            Return stepnumber & "," & scaleInfo
          End If
        End If
      Next

      Return ""
    Catch ex As Exception
      Return ""
    End Try

  End Function


End Class
