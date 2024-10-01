Imports Utilities.Sql

Public Class DrugroomPreview : Inherits MarshalByRefObject
  Private ReadOnly controlCode As ControlCode

  'Table populated with Central DyelotsBulkedRecipe for current dyelot
  Public SqlConnection As String = "data source=10.1.21.200;initial catalog=BatchDyeingCentral;user id=Adaptive;password=Control;app=ControlCode"

  Public Dyelot As String
  Public ReDye As Integer
  Public Job As String

  Public CycleTime As Integer
  Public CycleWater As Integer

  Public StepTime As Integer
  Public TimeInStep As Integer

  Public BatchWeight As Double
  Public LiquorRatio As Double
  Public LiquorVolume As Double

  Public StyleCode As String
  Public ColorCode As String
  '  Public NozzleCode As String

  Public RecipeStepCount As Integer

  Public BatchLocal As DataRow                       ' local dyelot row
  Public BatchHost As System.Data.DataRow            ' host dyelot row 
  Public BulkedRecipeHost As System.Data.DataTable   ' host bulked recipe rows

  Public BatchString As String                       ' Sub set of batch row as a delimited string for remote mimics
  Public BulkedRecipeString As String                ' Sub set of bulked recipe table as a delimited string for remote mimics
  Public BatchNotes As String                        ' Batch notes for remote mimics

  Public DataRead As Boolean



  Public RecipeCode As String
  Public RecipeName As String
  Public RecipeProducts As String
  Public RecipeRows As System.Data.DataTable




  'Private values - KP command parameters
  Private pTimeToTransfer As Long
  Public TimeToTransferTimer As New Timer
  Public WaitingAtTransferTimer As New TimerUp
  Public Tank1Calloff As Long
  Public DestinationTank As Integer
  Public IsManualCalloff As Boolean
  Public PreviewStatus As String

  Public ReadOnly Property KitchenDestination As String
    Get
      If DestinationTank = 1 Then
        Return "Add"
      ElseIf DestinationTank = 2 Then
        Return "Reserve"
      Else
        Return ""
      End If
    End Get
  End Property


  Property Records As Integer
  Property CollectionRows As Integer
  Property CurrentRecipe As Boolean
  Property NextRecipe As Boolean


  ReadOnly Property TimeToTransferString As String
    Get
      Return TimeToTransferTimer.ToString
    End Get
  End Property
  ReadOnly Property WaitingAtTransferString As String
    Get
      If WaitingAtTransferTimer.Seconds > 0 Then
        If WaitingAtTransferTimer.IsPaused Then
          Return "0"
        Else
          Return WaitingAtTransferTimer.ToString
        End If
      Else
        Return ""
      End If
    End Get
  End Property
  ReadOnly Property WaitingAtTransferSeconds As Integer
    Get
      Return WaitingAtTransferTimer.Seconds
    End Get
  End Property
  ReadOnly Property TimeToTransferSeconds As Integer
    Get
      Return TimeToTransferTimer.Seconds
    End Get
  End Property

  Public Sub New(ByVal controlCode As ControlCode)
    Me.controlCode = controlCode
    Me.controlCode = controlCode
    SqlConnection = DefaultSetting(controlCode.Parent.Setting("SqlConnection"), SqlConnection)

    '   CurrentrecipeRows = New RecipeRowCollection
  End Sub

  Sub LoadDyelot()
    DataRead = False

    With controlCode
      If Not .Parent.IsProgramRunning Then Exit Sub

      ' Get the dyelot and redye from the job 
      SetDyelotReDye(.Parent.Job)

      ' Get the local row on the main thread and the host row on a worker thread
      Dim table = .Parent.DbGetDataTable("SELECT * FROM Dyelots WHERE Dyelot=" & SqlString(Dyelot) & " AND Redye=" & ReDye.ToString)
      If table IsNot Nothing AndAlso table.Rows.Count = 1 Then
        BatchLocal = table.Rows(0)
        UpdateVariables(BatchLocal)
        BatchString = GetBatchString(BatchLocal)
      End If
    End With

    ' Run the SQL commands on a separate thread so we dont't hold up the control system
    Threading.ThreadPool.QueueUserWorkItem(AddressOf ReadAsync)
  End Sub

  Private Sub ReadAsync(ByVal args As Object)
    BatchHost = GetDyelotRowFromHost(Dyelot, ReDye)
    BulkedRecipeHost = GetBulkedRecipeTableFromHost(Dyelot, ReDye)

    If BatchHost IsNot Nothing Then
      UpdateVariables(BatchHost)
      BatchString = GetBatchString(BatchHost)
      BulkedRecipeString = GetBulkedRecipeString(BulkedRecipeHost)
    End If

    If BulkedRecipeHost IsNot Nothing AndAlso BulkedRecipeHost.Rows.Count > 0 Then
      RecipeStepCount = BulkedRecipeHost.Rows.Count
    Else
      RecipeStepCount = 0
    End If

    'Fill these in for the remote mimic
    '    controlCode.MimicBatch = BatchString
    '    controlCode.MimicBatchNotes = BatchNotes
    '    controlCode.MimicRecipeSteps = BulkedRecipeString


    ' Fill these in for remote display
    controlCode.DispenseProducts = BulkedRecipeString



    DataRead = True
  End Sub

  Private Function GetDyelotRowFromHost(dyelot As String, redye As Integer) As DataRow
    Dim sql As String = Nothing
    Try
      If String.IsNullOrEmpty(SqlConnection) Then Return Nothing
      sql = "SELECT * FROM Dyelots WHERE Dyelot=" & SqlString(dyelot) & " AND Redye=" & redye.ToString

      Return GetDataRow(SqlConnection, sql)
    Catch ex As Exception
      ' Log error  ??
    End Try
    Return Nothing
  End Function

  Private Function GetBulkedRecipeTableFromHost(dyelot As String, redye As Integer) As DataTable
    Dim sql As String = Nothing
    Try
      If String.IsNullOrEmpty(SqlConnection) Then Return Nothing
      'sql = "SELECT * FROM DyelotsBulkedRecipe WHERE Dyelot=" & SqlString(dyelot) & " AND Redye=" & redye.ToString & " ORDER BY DisplayOrder"
      sql = "SELECT * FROM DyelotsBulkedRecipe WHERE Batch_Number=" & SqlString(dyelot) & " AND Redye=" & redye.ToString & " ORDER BY ADD_NUMBER, ADD_SEQUENCE"

      Return GetDataTable(SqlConnection, sql)
    Catch ex As Exception
      ' Log error  ??
    End Try
    Return Nothing
  End Function

  Private Sub UpdateVariables(row As DataRow)
    If row Is Nothing Then Exit Sub

    '  BatchWeight = NullToZeroDouble(row("AE_ActualWeight"))  'BatchWeight
    '  StyleCode = NullToNothing(row("StyelCode").ToString)
    '  StyleName
    '  RecipeCode = NullToNothing(row("RecipeCode").ToString)
    '  RecipeName = NullToNothing(row("RecipeName").ToString)

    '  LiquorRatio = NullToZeroDouble(row("LiquorRatio"))
    '  LiquorVolume = NullToZeroDouble(row("LiquorVolume"))


    ' Some fields: AE_ColorNumber (int), AE_CycleNumber (int), AE_ColorCode (Text)
    '   PurchaseOrdrers "PurchaseOrders, vnarchar(300)"

    BatchNotes = row("Notes").ToString

    With controlCode
      '  If BatchWeight > 0 Then .BatchWeight = CInt(BatchWeight)
      '  If LiquorRatio > 0 Then .LiquorRatio = CInt(LiquorRatio)
      '  If LiquorVolume > 0 Then .WorkingVolume = CInt(LiquorVolume)
    End With
  End Sub

  Private Sub ClearVariables()
    BatchWeight = 0
    LiquorRatio = 0
    LiquorVolume = 0

    StyleCode = Nothing
    ColorCode = Nothing
    '   NozzleCode = Nothing

    BatchNotes = Nothing
  End Sub

  ' Get delimited string of interesting columns from the batch row so we can populate remote mimic controls (can't send objects to remote mimics)
  Private Function GetBatchString(row As DataRow) As String
    If row Is Nothing Then Return Nothing
    Try
      Return row("Dyelot").ToString & "|" & row("ReDye").ToString & "|" &
             row("StyleCode").ToString & "|" &
             row("RecipeCode").ToString & "|" & row("Program").ToString & "|" &
             row("LiquorRatio").ToString & "|" & row("LiquorVolume").ToString
    Catch ex As Exception
      ' Do nothing
      Return Nothing
    End Try
    'Return row("ID").ToString & "|" & row("Dyelot").ToString & "|" & row("ReDye").ToString & "|" & row("Orders").ToString & "|" &
    '       row("StyleCode").ToString & "|" & row("SubstrateCode").ToString & "|" & row("ColorCode").ToString & "|" &
    '       row("RecipeCode").ToString & "|" & row("Program").ToString & "|" & row("BatchWeight").ToString & "|" &
    '       row("LiquorRatio").ToString & "|" & row("LiquorVolume").ToString
  End Function

  ' Get delimited string of interesting columns from the bulked recipe table so we can populate remote mimic controls (can't send objects to remote mimics)
  Private Function GetBulkedRecipeString(table As DataTable) As String
    If table Is Nothing OrElse table.Rows.Count <= 0 Then Return Nothing

    Dim data As String
    Dim str As String
    For Each row As DataRow In table.Rows
      'str = row("ID").ToString & "|" & row("StepTypeID").ToString & "|" & row("StepID").ToString & "|" & row("StepNumber").ToString & "|" &
      '      row("StepCode").ToString & "|" & row("StepDescription").ToString & "|" & row("Grams").ToString & "|" & row("Pounds").ToString & "|" &
      '      row("DispenseGrams").ToString & "|" & row("DispensePounds").ToString & "|" & row("DispenseSource").ToString & "|" & row("DispenseTime").ToString

      str = row("ID").ToString & "|" & row("ADD_NUMBER").ToString & "|" & row("INGREDIENT_ID").ToString & "|" & row("INGREDIENT_DESC").ToString & "|" &
            row("DISPENSED").ToString & "|" & row("AMOUNT").ToString & "|" & row("DAMOUNT").ToString & "|" & row("UNITS").ToString & "|" &
            row("DSTATE").ToString

      If data Is Nothing Then
        data = str
      Else
        data &= ";" & str
      End If
    Next
    Return data
  End Function

#If 0 Then
  CREATE TABLE [dbo].[DyelotsBulkedRecipe](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[BATCH_NUMBER] [varchar](50) NOT NULL,
	[Redye] [int] NULL CONSTRAINT [DF_DyelotsBulkedRecipe_Redye]  DEFAULT (0),
	[ADD_NUMBER] [int] NOT NULL,
	[ADD_SEQUENCE] [int] NOT NULL,
	[INGREDIENT_ID] [int] NOT NULL,
	[INGREDIENT_DESC] [nvarchar](500) NOT NULL,
	[DISPENSED] [varchar](20) NOT NULL,
	[AMOUNT] [float] NULL,
	[DAMOUNT] [float] NULL,
	[UNITS] [nvarchar](20) NULL,
	[GRAVITY] [float] NULL,
	[DispenseID] [int] NULL,
	[DProductID] [int] NULL,
	[DState] [int] NULL,
	[DGrams] [float] NULL,
	[ProductDispensedState] [int] NULL,
	[AddDispensedState] [int] NULL,
	[Created] [datetime] NULL CONSTRAINT [DF_DyelotsBulkedRecipe_Created]  DEFAULT (getdate()),
	[DTime] [datetime] NULL,
	[DSource] [nchar](20) NULL,
	[DTimeCorrected] [int] NULL,
	[Coupled] [int] NOT NULL DEFAULT (0),
	[Notes] [text] NULL,
 CONSTRAINT [PK_DyelotsBulkedRecipe] PRIMARY KEY NONCLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

ALTER TABLE [dbo].[DyelotsBulkedRecipe]  WITH CHECK ADD  CONSTRAINT [FKDyelotsBulkedRecipeDyelots] FOREIGN KEY([BATCH_NUMBER], [Redye], [Coupled])
REFERENCES [dbo].[Dyelots] ([Dyelot], [ReDye], [Coupled])
ON UPDATE CASCADE
ON DELETE CASCADE
GO

ALTER TABLE [dbo].[DyelotsBulkedRecipe] CHECK CONSTRAINT [FKDyelotsBulkedRecipeDyelots]
GO
#End If

  Private Sub SetDyelotReDye(job As String)
    Me.Job = job
    Dim data = job.Split("@".ToCharArray)
    Me.Dyelot = data(0)
    Me.ReDye = 0
    If data.Length = 2 Then Me.ReDye = CInt(data(1))
  End Sub

  Sub UpdateNozzleCode(nozzleCode As String)
    With controlCode
      If BatchLocal IsNot Nothing Then
        BatchLocal("NozzleCode") = nozzleCode
        BatchString = GetBatchString(BatchLocal)
      End If

      If BatchHost IsNot Nothing Then
        BatchHost("NozzleCode") = nozzleCode
        BatchString = GetBatchString(BatchHost)
      End If

      controlCode.MimicBatch = BatchString

      .Parent.DbExecute("UPDATE Dyelots SET NozzleCode ='" & nozzleCode & "' WHERE Dyelot=" & SqlString(Dyelot) & " AND Redye=" & ReDye.ToString)
    End With
  End Sub


  Public Sub ResetAll()
    Dyelot = ""
    ReDye = 0
    Job = ""

    DestinationTank = 0
    CurrentRecipe = False
    Records = 0

    RecipeRows = Nothing
    RecipeStepCount = 0

    '  collectionRows_ = Nothing

    '   CurrentrecipeRows.Clear() ' TODO
    '    collectionRows_ = CurrentrecipeRows.Count ' TODO
    NextRecipe = False
  End Sub

  Public Sub Run()
    With controlCode

      'Dispenser & Drugroom Preview Determination
      'if a program is running look to see if the next drop has a manual.
      If .Parent.IsProgramRunning Then
        'for dispenser
        If Not (.KP.KP1.IsOn Or .LA.KP1.IsOn Or .KA.IsOn) Then
          Job = Dyelot
          If ReDye > 0 Then Job += "@" & ReDye.ToString("#0")
          .DrugroomDisplayJob1 = Job
        End If


        If .KP.KP1.IsOn Then

          If .KP.KP1.AddDispenseError AndAlso Not .Tank1Ready Then
            Tank1Calloff = .KP.KP1.DispenseCalloff
            DestinationTank = .KP.KP1.DispenseTank
            If Not .DrugroomCalloffRefreshed Then
              pTimeToTransfer = .GetTimeBeforeTransfer(CType(DestinationTank, EKitchenDestination))
              .DrugroomCalloffRefreshed = True
            End If
            PreviewStatus = "Dispense error for calloff " & Tank1Calloff
          ElseIf .KP.KP1.ManualAdd AndAlso Not .Tank1Ready Then
            Tank1Calloff = .KP.KP1.Param_Calloff
            DestinationTank = .KP.KP1.DispenseTank
            If Not .DrugroomCalloffRefreshed Then
              pTimeToTransfer = .GetTimeBeforeTransfer(CType(DestinationTank, EKitchenDestination))
              .DrugroomCalloffRefreshed = True
            End If
            PreviewStatus = "Manual add for calloff "
          ElseIf IsManualProduct(.KP.KP1.Param_Calloff, .Parameters.DDSEnabled = 1, .Parameters.DispenseEnabled = 1) And Not .Tank1Ready Then
            Tank1Calloff = .KP.KP1.Param_Calloff
            DestinationTank = .KP.KP1.DispenseTank
            If .DrugroomCalloffRefreshed = False Then pTimeToTransfer = GetTimeBeforeTransfer(CType(DestinationTank, EKitchenDestination))
            .DrugroomCalloffRefreshed = True
            If CurrentRecipe Then
              PreviewStatus = "Manual add for calloff " & Tank1Calloff
            Else
              PreviewStatus = "No Recipe manual add for calloff " & Tank1Calloff
            End If
          Else
            DestinationTank = 0
            PreviewStatus = ""
            Tank1Calloff = 0
            TimeToTransferTimer.TimeRemaining = 0
          End If

        ElseIf .LA.KP1.IsOn Then

          If .LA.KP1.AddDispenseError And Not .Tank1Ready Then
            Tank1Calloff = .LA.KP1.Param_Calloff
            DestinationTank = .LA.KP1.DispenseTank
            If .DrugroomCalloffRefreshed = False Then pTimeToTransfer = GetTimeBeforeTransfer(CType(DestinationTank, EKitchenDestination))
            .DrugroomCalloffRefreshed = True
            PreviewStatus = "Dispense error for calloff " & Tank1Calloff
          ElseIf .LA.KP1.ManualAdd And Not .Tank1Ready Then
            Tank1Calloff = .LA.KP1.Param_Calloff
            DestinationTank = .LA.KP1.DispenseTank
            If .DrugroomCalloffRefreshed = False Then pTimeToTransfer = GetTimeBeforeTransfer(CType(DestinationTank, EKitchenDestination))
            .DrugroomCalloffRefreshed = True
            PreviewStatus = "Manual add for calloff " & Tank1Calloff
          ElseIf IsManualProduct(.LA.KP1.Param_Calloff, .Parameters.DDSEnabled = 1, .Parameters.DispenseEnabled = 1) And Not .Tank1Ready Then
            DestinationTank = .LA.KP1.DispenseTank
            Tank1Calloff = .LA.KP1.Param_Calloff ' Check This

            If .DrugroomCalloffRefreshed = False Then pTimeToTransfer = GetTimeBeforeTransfer(CType(DestinationTank, EKitchenDestination))
            .DrugroomCalloffRefreshed = True
            If CurrentRecipe Or NextRecipe Then
              PreviewStatus = "Manual add for calloff " & Tank1Calloff
            Else
              PreviewStatus = "No Recipe manual add for calloff " & Tank1Calloff
            End If
          Else
            DestinationTank = 0
            PreviewStatus = ""
            Tank1Calloff = 0
            TimeToTransferTimer.TimeRemaining = 0
          End If

        ElseIf .LAActive Then
          DestinationTank = 0
          PreviewStatus = ""
          Tank1Calloff = 0
          TimeToTransferTimer.TimeRemaining = 0

        ElseIf Not .DrugroomCalloffRefreshed Then
          ' TODO Why is this returning Calloff = 1 
          If IsManualProduct(.LA.KP1.Param_Calloff, .Parameters.DDSEnabled = 1, .Parameters.DispenseEnabled = 1) And Not .Tank1Ready Then
            Tank1Calloff = .LA.KP1.Param_Calloff ' Check This
            DestinationTank = GetNextDestinationTank()
            pTimeToTransfer = GetTimeBeforeTransfer(CType(DestinationTank, EKitchenDestination))
            If CurrentRecipe Then
              PreviewStatus = "Manual add for calloff " & Tank1Calloff
            ElseIf Tank1Calloff > 0 Then
              PreviewStatus = "No Recipe manual add for calloff " & Tank1Calloff
            Else
              PreviewStatus = ""
            End If
            .DrugroomCalloffRefreshed = True

          Else
            .DrugroomCalloffRefreshed = True
            DestinationTank = 0
            PreviewStatus = ""
            Tank1Calloff = 0
            TimeToTransferTimer.TimeRemaining = 0
          End If
        End If

        'rescan if step number has changed
        If .StepNumberWas <> .Parent.StepNumber Then
          .StepNumberWas = .Parent.StepNumber
          .DrugroomCalloffRefreshed = False
        End If

      Else
        .DrugroomCalloffRefreshed = False
      End If

    End With
  End Sub


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
      Dim programNumber As String = controlCode.Parent.ProgramNumberAsString
      Dim stepNumber As Integer : stepNumber = controlCode.Parent.StepNumber

      ' Get all program steps for this program
      Dim PrefixedSteps As String : PrefixedSteps = controlCode.Parent.PrefixedSteps
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

  Public Function GetNextDestinationTank() As Integer
    Dim returnValue As Integer = 0
    Try

      ' These separators are used by programs and prefixed steps
      Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
      Dim separator1 As String = Convert.ToChar(255)
      Dim separator2 As String = ","

      ' Get current program and step number
      Dim programNumber As String = controlCode.Parent.ProgramNumberAsString
      Dim stepNumber As Integer : stepNumber = controlCode.Parent.StepNumber

      'Use this to split out command parameters for a program step
      Dim parameters() As String
      Dim checkingfornotes() As String

      ' Get all program steps for this program
      Dim PrefixedSteps As String : PrefixedSteps = controlCode.Parent.PrefixedSteps
      Dim ProgramSteps() As String : ProgramSteps = PrefixedSteps.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

      'Loop through each step starting from the beginning of the program
      Dim StartChecking As Boolean = False
      Dim KPFound As Boolean = False
      Dim KP1 As String
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

            If StartChecking AndAlso (Not KPFound) Then
              Select Case Command(0).ToUpper.Trim
                Case "AP"
                  returnValue = 1 '   DestinationTank = 1
                  ' TODO - Set destination to EType?
                  Exit For

                Case "KP"
                  ' Only find first KP after LA
                  KPFound = True
                  KP1 = [Step](5)
                  Exit For

                Case "RP"
                  returnValue = 2 ' DestinationTank = 2
                  Exit For

              End Select
            End If
          End If
        End If
        ' Reached the upper bound of the program steps array
        If i = ProgramSteps.GetUpperBound(0) Then
          returnValue = 0
        End If

      Next i


      ' DestinationTank = 0
      If KP1.Length > 0 Then
        checkingfornotes = KP1.Split(CType("'", Char()))                '  Split(KP1, "'")
        parameters = checkingfornotes(0).Split(CType(",", Char()))      '  Split(checkingfornotes(0), ",")  '0 based array
        If parameters.GetUpperBound(0) >= 5 Then
          '0=KP, 1=Level, 2=Desired Temp, 3=Mix Time, 4=Mix on/off, 5=DispenseTank(1=Add/2=Reserve)
          ' VB6 TODO
          'If IsNumeric(parameters(1)) And IsNumeric(parameters(2)) And IsNumeric(parameters(3)) And IsNumeric(parameters(4)) And IsNumeric(parameters(5)) Then
          DestinationTank = CInt(parameters(5))
        End If
      End If

      returnValue = DestinationTank

    Catch ex As Exception
      Utilities.Log.LogError(ex.ToString)
    End Try
    Return returnValue
  End Function



End Class
