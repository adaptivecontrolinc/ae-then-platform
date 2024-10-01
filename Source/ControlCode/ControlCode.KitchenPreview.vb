Imports Utilities.Translations

Partial Class ControlCode

  Public Dyelot As String
  Public ReDye As Integer
  Public Job As String

  Public RecipeCode As String
  Public RecipeName As String
  Public RecipeProducts As String
  Public RecipeRows As System.Data.DataTable

  '  Public StepNumberWas As Integer  ' In ControlCode for moment 

  Public StyleCode As String


  Public KitchenCalloffRefreshed As Boolean
  Public KitchenJob As String
  Public KitchenManualProduct As Boolean

  'Private kitchenDestination_ As String
  'Public Property KitchenDestination As String
  '  Private Set(value As String)
  '    kitchenDestination_ = value
  '  End Set
  '  Get
  '    If Tank1Destination = EKitchenDestination.Add Then
  '      kitchenDestination_ = "Add"
  '    ElseIf Tank1Destination = EKitchenDestination.Reserve Then
  '      kitchenDestination_ = "Reserve"
  '    Else : kitchenDestination_ = ""
  '    End If
  '    Return kitchenDestination_
  '  End Get
  'End Property

  Private kitchenStatus_ As String
  Public Property KitchenStatus As String
    Private Set(value As String)
      kitchenStatus_ = value
    End Set
    Get
      Return kitchenStatus_
    End Get
  End Property

  Private kitchenTimeBeforeTransfer_ As New Timer
  Public Property KitchenTimeBeforeTransfer As Integer
    Get
      Return kitchenTimeBeforeTransfer_.Seconds
    End Get
    Set(value As Integer)
      kitchenTimeBeforeTransfer_.Seconds = value
    End Set
  End Property

  Public Property Tank1Calloff As Integer

  Private tank1Destination_ As EKitchenDestination
  Public Property Tank1Destination As EKitchenDestination
    Private Set(value As EKitchenDestination)
      tank1Destination_ = value
    End Set
    Get
      Return tank1Destination_
    End Get
  End Property

  Private tank1Status_ As String
  Public ReadOnly Property Tank1Status As String
    Get
      tank1Status_ = Translate("Idle")
      If Tank1Ready Then tank1Status_ = Translate("Idle")
      If CK.IsOn Then tank1Status_ = CK.Status
      If KP.KP1.IsOn Then tank1Status_ = KP.KP1.Status
      If LA.KP1.IsOn Then tank1Status_ = LA.KP1.Status
      If KA.IsOn Then tank1Status_ = KA.Status
      Return tank1Status_
    End Get
  End Property

  Private tank1TimeWaiting_ As Integer
  Public ReadOnly Property Tank1TimeWaiting As Integer
    Get
      tank1TimeWaiting_ = 0
      If KP.KP1.IsOn AndAlso KP.KP1.TimerUpWaitReady.IsRunning Then tank1TimeWaiting_ = KP.KP1.TimerUpWaitReady.Seconds
      If LA.KP1.IsOn AndAlso KP.KP1.TimerUpWaitReady.IsRunning Then tank1TimeWaiting_ = KP.KP1.TimerUpWaitReady.Seconds

      Return tank1TimeWaiting_
    End Get
  End Property

  Private addStatus_ As String
  Public ReadOnly Property AddStatus As String
    Get
      addStatus_ = Translate(" ")
      If AddReady Then addStatus_ = Translate("Ready")
      If AD.IsOn Then addStatus_ = AD.StateString
      If AF.IsOn Then addStatus_ = AF.Status
      If AddControl.IsActive OrElse AddControl.MixIsActive Then addStatus_ = AddControl.Status
      Return addStatus_
    End Get
  End Property

  Private reserveStatus_ As String
  Public ReadOnly Property ReserveStatus As String
    Get
      reserveStatus_ = Translate(" ")
      If RD.IsOn Then reserveStatus_ = RD.Status
      If RF.IsOn Then reserveStatus_ = RF.Status
      If RT.IsOn Then reserveStatus_ = RT.Status
      If ReserveControl.IsActive OrElse ReserveControl.IsHeatActive Then reserveStatus_ = ReserveControl.Status
      Return reserveStatus_
    End Get
  End Property

  Public Sub CheckTank1Status()
    Try

      ' No current tank preparation active
      If Not (KP.KP1.IsOn OrElse LA.KP1.IsOn OrElse KA.IsOn) Then
        KitchenJob = Parent.Job
      End If

      ' Which Kitchen Command is active: KP/LA/KA
      If KP.KP1.IsOn Then
        ' Update local Calloff
        Tank1Calloff = KP.KP1.Param_Calloff ' KP.KP1.AddCallOff 

        ' Waiting for Tank1?
        If Not Tank1Ready Then
          Tank1Destination = CType(KP.KP1.DispenseTank, EKitchenDestination) ' KP.KP1.Param_Destination
          If RecipeProducts.Length > 0 Then

          End If
          kitchenStatus_ = Tank1Status 'Add Status for Lack of Recipe
          If Not KitchenCalloffRefreshed Then
            KitchenTimeBeforeTransfer = 0 ' TODO GetTimeBeforeTransfer(KP.KP1.Param_Destination)
            KitchenCalloffRefreshed = True
          End If
        Else
          ' Clear tank activity
          Tank1Destination = EKitchenDestination.Drain
          kitchenStatus_ = ""
          KitchenTimeBeforeTransfer = 0
        End If

      ElseIf LA.KP1.IsOn Then
        ' Update local Calloff
        Tank1Calloff = LA.KP1.Param_Calloff

        If Not Tank1Ready Then
          Tank1Destination = LA.KP1.Param_Destination
          kitchenStatus_ = Tank1Status
          If Not KitchenCalloffRefreshed Then
            KitchenTimeBeforeTransfer = GetTimeBeforeTransfer(LA.KP1.Param_Destination)
            KitchenCalloffRefreshed = True
          End If
        Else
          ' Clear tank activity
          Tank1Destination = EKitchenDestination.Drain
          kitchenStatus_ = ""
          Tank1Calloff = 0
          KitchenTimeBeforeTransfer = 0
        End If

      ElseIf LAActive Then
        ' Clear tank activity
        Tank1Destination = EKitchenDestination.Drain
        kitchenStatus_ = ""
        Tank1Calloff = 0
        KitchenTimeBeforeTransfer = 0

      ElseIf Not KitchenCalloffRefreshed Then

        If KitchenManualProduct Then '  If DrugroomPreview.IsManualProduct(GetKPCalloff(Me), Parameters_DDSEnabled, Parameters_DispenseEnabled) And (Not LAActive) Then 
          '                               DrugroomPreview.GetNextDestinationTank Me
          '                               DrugroomPreview.GetTimeBeforeTransfer(Me, DrugroomPreview.DestinationTank)
          '                               If DrugroomPreview.CurrentRecipe Then
          '                               DrugroomPreview.PreviewStatus = "Manual add for calloff " & DrugroomPreview.Tank1Calloff
          '                               Else
          '                               DrugroomPreview.PreviewStatus = "No Recipe manual add for calloff " & DrugroomPreview.Tank1Calloff
          '                               End If
          '                               DrugroomCalloffRefreshed = True

        Else
          ' Clear tank activity
          Tank1Destination = EKitchenDestination.Drain
          kitchenStatus_ = ""
          Tank1Calloff = 0
          KitchenTimeBeforeTransfer = 0

        End If
      Else
        ' Clear tank activity
        Tank1Destination = EKitchenDestination.Drain
        kitchenStatus_ = ""
        Tank1Calloff = 0
        KitchenTimeBeforeTransfer = 0
      End If

      ' Re-Scan if step number changed
      If StepNumberWas <> Parent.StepNumber Then
        StepNumberWas = Parent.StepNumber
        KitchenCalloffRefreshed = False
      End If

      ' Look-Ahead Enable
      ' If Parameters_LookAheadEnabled = 0 Then LAActive = False
      ' Maybe use Setting instead?

    Catch ex As Exception
      ' Just for debugging
      Dim message As String = ex.Message
    End Try
  End Sub


#Region " LOAD DYELOT AND DYELOTSBULKEDRECIPE DETAILS "

  '***********************************************************************************************
  '******                Update the Dyelot & Redye from the local database                  ******
  '***********************************************************************************************
  Public Sub LoadRecipeDetailsActiveJob()
    Dim sql As String = ("SELECT Dyelot,Redye, RecipeCode, RecipeName, StyleCode FROM Dyelots WHERE State=2 ORDER BY StartTime")
    Try
      Dim dt As System.Data.DataTable = Parent.DbGetDataTable(sql)
      If (dt IsNot Nothing) AndAlso (dt.Rows.Count > 0) Then
        For Each dr As System.Data.DataRow In dt.Rows
          'Get updated Dyelot/Redye values from Datatable
          Dyelot = Utilities.Null.NullToEmptyString(dr("Dyelot"))
          ReDye = Utilities.Null.NullToZeroInteger(dr("ReDye"))

          Job = Dyelot
          If ReDye > 0 Then Job += "@" & ReDye.ToString("#0")

          RecipeCode = Utilities.Null.NullToEmptyString(dr("RecipeCode"))
          RecipeName = Utilities.Null.NullToEmptyString(dr("RecipeName"))
          StyleCode = Utilities.Null.NullToEmptyString(dr("StyleCode"))
        Next
      End If

    Catch ex As Exception
      'Clear if there's a problem
      RecipeCode = ""
      RecipeName = ""
      StyleCode = ""

      Utilities.Log.LogError(ex, sql)
    End Try
  End Sub

#End Region

#Region " GET NEXT CALLOFF "

  '***********************************************************************************************
  '****                      This function returns the next tank calloff                      ****
  '****                   based on the destination tank declared in KP command                ****
  '***********************************************************************************************
  Public Function GetNextCalloff() As Integer
    Dim returnValue As Integer = 0
    Try

      ' These separators are used by programs and prefixed steps
      Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
      Dim separator1 As String = Convert.ToChar(255)
      Dim separator2 As String = ","

      ' Get current program and step number
      Dim programNumber As String = Parent.ProgramNumberAsString
      Dim stepNumber As Integer : stepNumber = Parent.StepNumber

      ' Get all program steps for this program
      Dim PrefixedSteps As String : PrefixedSteps = Parent.PrefixedSteps
      Dim ProgramSteps() As String : ProgramSteps = PrefixedSteps.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

      'Loop through each step starting from the beginning of the program
      Dim StartChecking As Boolean = False
      Dim Command() As String

      'Look for the next AP, KP, or RP command
      For i = ProgramSteps.GetLowerBound(0) To ProgramSteps.GetUpperBound(0)
        'Ignore step 0
        If i > 0 Then
          ' Current Step
          Dim [Step] = ProgramSteps(i).Split(separator1.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
          If [Step].GetUpperBound(0) >= 1 Then
            If ([Step](0)) = programNumber AndAlso (CDbl([Step](1)) = stepNumber - 1) Then StartChecking = True
            Command = [Step](5).Split(separator2.ToCharArray)

            If StartChecking Then
              If (Command(0) = "AP") OrElse (Command(0) = "KP") OrElse (Command(0) = "RP") Then
                returnValue += 1
                Exit For
              End If
            Else
              'TODO need this portion?
              If (Command(0) = "AP") OrElse (Command(0) = "KP") OrElse (Command(0) = "RP") Then
                returnValue += 1
              End If
            End If
          End If
        End If
        ' Reached the upper bound of the program steps array
        If i = ProgramSteps.GetUpperBound(0) Then
          returnValue = 0
        End If

      Next i

    Catch ex As Exception
      Utilities.Log.LogError(ex.ToString)
    End Try

    'Return Default Value 
    Return returnValue
  End Function

#End Region

#Region " GET TIME BEFORE TRANSFER "

  '***********************************************************************************************
  '****                   This function returns the time before next transfer                 ****
  '****                   based on the destination tank declared in KP command                ****
  '***********************************************************************************************
  Public Function GetTimeBeforeTransfer(ByVal destination As EKitchenDestination) As Integer
    Dim returnValue As Integer = 0
    Try

      ' These separators are used by programs and prefixed steps
      Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
      Dim separator1 As String = Convert.ToChar(255)
      Dim separator2 As String = ","

      ' Get current program and step number
      Dim programNumber As String = Parent.ProgramNumberAsString
      Dim stepNumber As Integer : stepNumber = Parent.StepNumber

      ' Get all program steps for this program
      Dim PrefixedSteps As String : PrefixedSteps = Parent.PrefixedSteps
      Dim ProgramSteps() As String : ProgramSteps = PrefixedSteps.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

      ' Loop through each step starting from the beginning of the program
      Dim StartChecking As Boolean = False
      Dim Command() As String

      ' Only check for a destination transfer if Destination Tank is set in controlcode
      If destination = EKitchenDestination.Add Then
        'Look for the next AT or AC command
        For i = ProgramSteps.GetLowerBound(0) To ProgramSteps.GetUpperBound(0)
          'Ignore step 0
          If i > 0 Then
            Dim [Step] = ProgramSteps(i).Split(separator1.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

            If [Step].GetUpperBound(0) >= 1 Then
              If ([Step](0)) = programNumber AndAlso (CDbl([Step](1)) = stepNumber - 1) Then StartChecking = True
              Command = [Step](5).Split(separator2.ToCharArray)

              If StartChecking Then
                If (Command(0) = "AT") OrElse (Command(0) = "AC") Then
                  returnValue += CInt([Step](3))
                  StartChecking = False
                End If
              End If
            End If
          End If
          ' Reached the upper bound of the program steps array
          If i = ProgramSteps.GetUpperBound(0) Then returnValue = 0
        Next i

      ElseIf destination = EKitchenDestination.Reserve Then
        'Look for the next RT command
        For i = ProgramSteps.GetLowerBound(0) To ProgramSteps.GetUpperBound(0)
          'Ignore step 0
          If i > 0 Then
            Dim [Step] = ProgramSteps(i).Split(separator1.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

            If [Step].GetUpperBound(0) >= 1 Then
              If ([Step](0)) = programNumber AndAlso (CDbl([Step](1)) = stepNumber - 1) Then StartChecking = True
              Command = [Step](5).Split(separator2.ToCharArray)

              If StartChecking Then
                If (Command(0) = "RT") Then
                  returnValue += CInt([Step](3))
                  StartChecking = False
                End If
              End If
            End If
          End If
          If i = ProgramSteps.GetUpperBound(0) Then returnValue = 0
        Next i
      End If

    Catch ex As Exception
      Utilities.Log.LogError(ex)
    End Try

    'Return Default Value 
    Return returnValue
  End Function

#End Region

  '***********************************************************************************************
  '****           This function returns a True if the next Calloff has a manual product       ****
  '****                   "Dispensed" Field value for product row is set to "M"               ****
  '***********************************************************************************************
  Public Function IsManualProduct(ByVal Calloff As Integer, DyeDispenserEnabled As Boolean,
                                  ChemicalDispenserEnabled As Boolean) As Boolean
    'Default Return Value
    Dim returnValue As Boolean = True
    Try
      ' Calloff not defined, return manual product
      If Calloff = 0 Then
        Return returnValue
        Exit Function
      End If
      ' No dispensers enabled, return manual product
      If Not (DyeDispenserEnabled AndAlso ChemicalDispenserEnabled) Then
        Return returnValue
        Exit Function
      End If

      ' Goal to determine if this calloff has manual products
      Dim stepCalloff As Integer = 0
      Dim chemicalId As Integer = 0
      Dim dyeId As Integer = 0
      Dim RowCountCheck As Integer 'just in case the calloff does not exist in the recipe.

      ' Check each row - Make sure we've got something to check
      If RecipeRows.Rows.Count > 0 Then
        For Each dr As System.Data.DataRow In RecipeRows.Rows
          ' Look for the current calloff
          stepCalloff = Utilities.Null.NullToZeroInteger(dr("StepNumber"))
          If stepCalloff = Calloff Then
            chemicalId = Utilities.Null.NullToZeroInteger(dr("ChemicalID"))
            dyeId = Utilities.Null.NullToZeroInteger(dr("DyeId"))
            If chemicalId > 0 AndAlso Not ChemicalDispenserEnabled Then returnValue = True
            If dyeId > 0 AndAlso Not DyeDispenserEnabled Then returnValue = True
          Else
            'StepNumber <> calloff
            RowCountCheck += 1
            If RowCountCheck = RecipeRows.Rows.Count Then
              ' We have checked each row and the calloff is not in the recipe
              returnValue = True
            End If

          End If
        Next
      End If

      ' No Manual products found
      returnValue = False

    Catch ex As Exception

    End Try
    Return returnValue
  End Function

  Public Sub SelectFromCentralBulkedRecipeTable() ' Runs on a background thread to query the central BDC Database for recipe steps
    If Parent.IsProgramRunning Then
      ' Clear products to begin
      RecipeProducts = ""

#If 0 Then
            Dim sql As String = "SELECT * FROM DyelotsBulkedRecipe WHERE Dyelot='" & Dyelot.ToString & "' AND " &
                        "ReDye=" & ReDye.ToString("#0") & " AND " &
                        "(ChemicalID IS NOT NULL OR DyeID IS NOT NULL) " &
                        "ORDER BY StepNumber "
#End If

      Dim sql As String = "SELECT * FROM DyelotsBulkedRecipe WHERE Dyelot='" & Dyelot.ToString & "' AND " &
                        "ReDye=" & ReDye.ToString("#0") & " AND " &
                        "(INGREDIENT_ID IS NOT NULL) " &
                        "ORDER BY ADD_NUMBER "

      Dim dtProducts As System.Data.DataTable = Utilities.Sql.GetDataTable(Settings.ConnectionStringBDC, sql, "DyelotsBulkedRecipe")
      If (dtProducts IsNot Nothing) AndAlso (dtProducts.Rows.Count > 0) Then
        ' Keep a local copy 
        RecipeRows = dtProducts
        ' Step through each and update the products string
#If 0 Then
                For Each drProduct As System.Data.DataRow In dtProducts.Rows
          RecipeProducts = RecipeProducts & drProduct("StepNumber").ToString & "|" &
                                          drProduct("Code").ToString & "|" &
                                          drProduct("Description").ToString & "|" &
                                          drProduct("Kilograms").ToString & "|" &
                                          drProduct("DKilograms").ToString & "|" &
                                          drProduct("DState").ToString & ";"

        Next
#End If
        For Each drProduct As System.Data.DataRow In dtProducts.Rows
          RecipeProducts = RecipeProducts & drProduct("ADD_NUMBER").ToString & "|" &
                                          drProduct("INGREDIENT_ID").ToString & "|" &
                                          drProduct("INGREDIENT_DESC").ToString & "|" &
                                          drProduct("Amount").ToString & "|" &
                                          drProduct("DAmount").ToString & "|" &
                                          drProduct("DState").ToString & ";"

        Next
      End If
    Else
      'Clear if there's a problem
      RecipeProducts = ""
    End If
  End Sub





#If 0 Then
  ' VB6 Remove



  
'Dispenser & Drugroom Preview Determination
'if a program is running look to see if the next drop has a manual.
  If Parent.IsProgramRunning Then



     'for dispenser
     If Not (KP.KP1.IsOn Or LA.KP1.IsOn Or KA.IsOn) Then
       DrugroomDisplayJob1 = Parent.Job
     End If




     If KP.KP1.IsOn Then


        If KP.KP1.AddDispenseError And Not Tank1Ready Then
           DrugroomPreview.Tank1Calloff = KP.KP1.AddCallOff
           DrugroomPreview.DestinationTank = KP.KP1.DispenseTank
            If DrugroomCalloffRefreshed = False Then DrugroomPreview.GetTimeBeforeTransfer Me, DrugroomPreview.DestinationTank
            DrugroomCalloffRefreshed = True
           DrugroomPreview.PreviewStatus = "Dispense error for calloff " & DrugroomPreview.Tank1Calloff
        ElseIf KP.KP1.ManualAdd And Not Tank1Ready Then
           DrugroomPreview.Tank1Calloff = KP.KP1.AddCallOff
           DrugroomPreview.DestinationTank = KP.KP1.DispenseTank
            If DrugroomCalloffRefreshed = False Then DrugroomPreview.GetTimeBeforeTransfer Me, DrugroomPreview.DestinationTank
            DrugroomCalloffRefreshed = True
           DrugroomPreview.PreviewStatus = "Manual add for calloff " & DrugroomPreview.Tank1Calloff
        ElseIf DrugroomPreview.IsManualProduct(KP.KP1.AddCallOff, Parameters_DDSEnabled, Parameters_DispenseEnabled) And Not Tank1Ready Then
           DrugroomPreview.DestinationTank = KP.KP1.DispenseTank
           If DrugroomCalloffRefreshed = False Then DrugroomPreview.GetTimeBeforeTransfer Me, DrugroomPreview.DestinationTank
           DrugroomCalloffRefreshed = True
           If DrugroomPreview.CurrentRecipe Then
              DrugroomPreview.PreviewStatus = "Manual add for calloff " & DrugroomPreview.Tank1Calloff
           Else
              DrugroomPreview.PreviewStatus = "No Recipe manual add for calloff " & DrugroomPreview.Tank1Calloff
           End If
        Else
           DrugroomPreview.DestinationTank = 0
           DrugroomPreview.PreviewStatus = ""
           DrugroomPreview.Tank1Calloff = 0
           DrugroomPreview.TimeToTransferTimer.TimeRemaining = 0
        End If


     ElseIf LA.KP1.IsOn Then


        If LA.KP1.AddDispenseError And Not Tank1Ready Then
           DrugroomPreview.Tank1Calloff = LA.KP1.AddCallOff
           DrugroomPreview.DestinationTank = LA.KP1.DispenseTank
            If DrugroomCalloffRefreshed = False Then DrugroomPreview.GetTimeBeforeTransfer Me, DrugroomPreview.DestinationTank
            DrugroomCalloffRefreshed = True
           DrugroomPreview.PreviewStatus = "Dispense error for calloff " & DrugroomPreview.Tank1Calloff
        ElseIf LA.KP1.ManualAdd And Not Tank1Ready Then
           DrugroomPreview.Tank1Calloff = LA.KP1.AddCallOff
           DrugroomPreview.DestinationTank = LA.KP1.DispenseTank
            If DrugroomCalloffRefreshed = False Then DrugroomPreview.GetTimeBeforeTransfer Me, DrugroomPreview.DestinationTank
            DrugroomCalloffRefreshed = True
           DrugroomPreview.PreviewStatus = "Manual add for calloff " & DrugroomPreview.Tank1Calloff
        ElseIf DrugroomPreview.IsManualProduct(LA.KP1.AddCallOff, Parameters_DDSEnabled, Parameters_DispenseEnabled) And Not Tank1Ready Then
           DrugroomPreview.DestinationTank = LA.KP1.DispenseTank
           If DrugroomCalloffRefreshed = False Then DrugroomPreview.GetTimeBeforeTransfer Me, DrugroomPreview.DestinationTank
           DrugroomCalloffRefreshed = True
           If DrugroomPreview.CurrentRecipe Or DrugroomPreview.NextRecipe Then
              DrugroomPreview.PreviewStatus = "Manual add for calloff " & DrugroomPreview.Tank1Calloff
           Else
              DrugroomPreview.PreviewStatus = "No Recipe manual add for calloff " & DrugroomPreview.Tank1Calloff
           End If
        Else
           DrugroomPreview.DestinationTank = 0
           DrugroomPreview.PreviewStatus = ""
           DrugroomPreview.Tank1Calloff = 0
           DrugroomPreview.TimeToTransferTimer.TimeRemaining = 0
        End If


     ElseIf LAActive Then
           DrugroomPreview.DestinationTank = 0
           DrugroomPreview.PreviewStatus = ""
           DrugroomPreview.Tank1Calloff = 0
           DrugroomPreview.TimeToTransferTimer.TimeRemaining = 0


     ElseIf Not DrugroomCalloffRefreshed Then
        If DrugroomPreview.IsManualProduct(GetKPCalloff(Me), Parameters_DDSEnabled, Parameters_DispenseEnabled) And (Not LAActive) Then
           DrugroomPreview.GetNextDestinationTank Me
           DrugroomPreview.GetTimeBeforeTransfer Me, DrugroomPreview.DestinationTank
           If DrugroomPreview.CurrentRecipe Then
              DrugroomPreview.PreviewStatus = "Manual add for calloff " & DrugroomPreview.Tank1Calloff
           Else
              DrugroomPreview.PreviewStatus = "No Recipe manual add for calloff " & DrugroomPreview.Tank1Calloff
           End If
           DrugroomCalloffRefreshed = True
        Else
           DrugroomCalloffRefreshed = True
           DrugroomPreview.DestinationTank = 0
           DrugroomPreview.PreviewStatus = ""
           DrugroomPreview.Tank1Calloff = 0
           DrugroomPreview.TimeToTransferTimer.TimeRemaining = 0
        End If
     End If
    
    'rescan if step number has changed
    If StepNumberWas <> Parent.StepNumber Then
        StepNumberWas = Parent.StepNumber
        DrugroomCalloffRefreshed = False
    End If
    
  Else
    DrugroomCalloffRefreshed = False
  End If



#End If



















End Class
