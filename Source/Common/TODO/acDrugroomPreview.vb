Public Class DrugroomPreviewOld
#If 0 Then

Attribute VB_Name = "acDrugroomPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***********************************************************************************************
'******      Always use Connection15 and Recordset15 to avoid ADO version nonsense        ******
'***********************************************************************************************
'Table populated with Central DyelotsBulkedRecipe for current dyelot
  Private pCurrentRecipe As Boolean
  Private pNextRecipe As Boolean

'Private storage for properties
  Private pDyelot As String
  Private pRedye As Long
  
'for the collection
  Dim pRecords As Integer
  Dim pCollectionRows As Integer
  Dim CurrentrecipeRows As acRecipeRowCollection

'Private values - KP command parameters
  Private pTimeToTransfer As Long
  Public TimeToTransferTimer As New acTimer
  Public WaitingAtTransferTimer As New acTimerUp
  Public Tank1Calloff As Long
  Public DestinationTank As Long
  Public IsManualCalloff As Boolean
  Public PreviewStatus As String
  
  'initialize
  Private Sub Class_Initialize()
    Set CurrentrecipeRows = New acRecipeRowCollection
  End Sub
  Public Property Get Records() As Integer
    Records = pRecords
  End Property
  Public Property Get CollectionRows() As Integer
    CollectionRows = pCollectionRows
  End Property
  Public Property Get CurrentRecipe() As Boolean
    CurrentRecipe = pCurrentRecipe
  End Property
  Public Property Get NextRecipe() As Boolean
    NextRecipe = pNextRecipe
  End Property
  Public Property Get Job() As String
    Job = pDyelot
    If pRedye > 0 Then Job = pDyelot & "@" & pRedye
  End Property
  Public Property Get TimeToTransferString() As String
    TimeToTransferString = TimerString(TimeToTransferTimer.TimeRemaining)
  End Property
  Public Property Get WaitingAtTransferString() As String
    If WaitingAtTransferTimer.IsPaused Then
       WaitingAtTransferString = "0"
    Else
       WaitingAtTransferString = TimerString(WaitingAtTransferTimer.TimeElapsed)
    End If
  End Property
  Public Property Get WaitingAtTransferSeconds() As Integer
   WaitingAtTransferSeconds = WaitingAtTransferTimer.TimeElapsed
  End Property
  Public Property Get TimeToTransferSeconds() As Integer
  TimeToTransferSeconds = TimeToTransferTimer.TimeRemaining
  End Property
'***********************************************************************************************
'******                          Clear all relavent values                                ******
'***********************************************************************************************
Public Sub ResetAll()
  
  pDyelot = ""
  pRedye = 0
  DestinationTank = 0
  pCurrentRecipe = False
  pRecords = 0
  CurrentrecipeRows.Clear
  pCollectionRows = CurrentrecipeRows.Count
  pNextRecipe = False
  
End Sub

'***********************************************************************************************
'******                Update the Dyelot & Redye from the local database                  ******
'***********************************************************************************************
Public Sub GetDyelotRedye()
On Error GoTo Err_GetDyelotRedye

'Set default values
  pDyelot = ""
  pRedye = 0

'Create an ADO connection to the local database
  ' get current job (dyelot and redye) from the local database instead of from parent.job
  ' carryover from PlantExplorer @1 bug
    Dim localConnection As ADODB.Connection15
    Set localConnection = New ADODB.Connection
    localConnection.Open "BatchDyeing"
    
    Dim rsDyelot As ADODB.Recordset15
    Set rsDyelot = localConnection.Execute("SELECT * FROM Dyelots WHERE State=2 order by StartTime")
    rsDyelot.MoveFirst
  
    If Not rsDyelot.EOF Then
      'Set Dyelot and Redye - these must be non-null because they are the primary key
      pDyelot = CStr(rsDyelot!Dyelot)
      pRedye = CLng(rsDyelot!Redye)
    End If

'Tidy up
    localConnection.Close
    Set localConnection = Nothing


'Avoid Error Handler
    Exit Sub

Err_GetDyelotRedye:

End Sub

Public Sub LoadRecipeSteps()
On Error GoTo Err_LoadRecipeSteps

'Load the local database values for current dyelot & redye
  GetDyelotRedye
  
'Check that Dyelot & Redye are set. If dyelot is blank there is a empty space " " not ""
  If pDyelot = " " Then Exit Sub

'clear next recipe
  pNextRecipe = False
  
'Reset records
  pRecords = 0
'clear collection
  CurrentrecipeRows.Clear

'Create ADO objects
  Dim con As ADODB.Connection
  Dim rs As ADODB.Recordset15
  Dim sql As String

  Set con = New ADODB.Connection
  Set rs = New ADODB.Recordset
  sql = "SELECT * FROM DyelotsBulkedRecipe WHERE BATCH_NUMBER='" & pDyelot & "' AND " & "Redye=" & pRedye
  
'Open the connection and get the data
  con.Open "BatchDyeingCentral", "sa", ""
  rs.Open sql, con, adOpenStatic, adLockReadOnly
 
'Make sure the recordset is fully populated then set records variable
 If rs.RecordCount > 0 Then
  rs.MoveLast: rs.MoveFirst
  pRecords = rs.RecordCount
  pCurrentRecipe = True
 Else
  pRecords = 0
  pCollectionRows = 0
  pCurrentRecipe = False
  Exit Sub
 End If
'Populate the recipe row collection
  Dim newRecipeRow As acRecipeRow
  Do While Not rs.EOF
    Set newRecipeRow = New acRecipeRow
    newRecipeRow.Fill rs
    CurrentrecipeRows.Add newRecipeRow
    rs.MoveNext
  Loop
   
 'set the number of collection rows just for info
   pCollectionRows = CurrentrecipeRows.Count
   
'FOR DEBUG ONLY - output basic step info
  Dim message As String
  rs.MoveFirst
  Do While Not rs.EOF
    message = rs("ID") & "  " & rs("INGREDIENT_ID") & "  " & rs("INGREDIENT_DESC")
    Debug.Print message
    rs.MoveNext
  Loop
  rs.MoveFirst
  
'Clean up
  con.Close
  Set con = Nothing
  Set rs = Nothing

'Avoid Error Handler
  Exit Sub

Err_LoadRecipeSteps:
  MsgBox Err.Description
  Set con = Nothing
  Set rs = Nothing
End Sub



'***********************************************************************************************
'****           This function returns a True if the next Calloff has a manual product       ****
'****                   "Dispensed" Field value for product row is set to "M"               ****
'***********************************************************************************************
Public Function IsManualProduct(ByVal Calloff As Long, DDSEnabled As Long, DispenseEnabled As Long) As Boolean
On Error GoTo Err_IsManualProduct
  If Calloff = 0 Then Exit Function
  
'Default Return Value
  IsManualCalloff = True
  IsManualProduct = True
  Tank1Calloff = Calloff
  Dim r As Integer
  Dim RowCountCheck As Integer 'just in case the calloff does not exist in the recipe.
  
'Make sure we've got something to check
  If CurrentrecipeRows.Count > 0 Then
    For r = 1 To CurrentrecipeRows.Count
      If CurrentrecipeRows(r).AddNumber = Calloff Then
         If ((CurrentrecipeRows(r).IngredientID >= 1000) And DispenseEnabled <> 1) Or _
         ((CurrentrecipeRows(r).IngredientID < 1000) And DDSEnabled <> 1) Or _
         (CurrentrecipeRows(r).Dispensed = False) Then
            IsManualCalloff = True
            IsManualProduct = True
            Tank1Calloff = Calloff
            Exit Function
         End If
      Else
        RowCountCheck = RowCountCheck + 1
        If RowCountCheck = CurrentrecipeRows.Count Then
           'we checked every row and the calloff is not in the recipe has to be manual.
           'should never happen but just in case
            IsManualCalloff = True
            IsManualProduct = True
            Tank1Calloff = Calloff
            Exit Function
        End If
      End If
    Next
    
  Else
    Exit Function
  End If

'No manual products found
  IsManualCalloff = False
  IsManualProduct = False
  Tank1Calloff = 0
 
'Avoid Error Handler
  Exit Function

Err_IsManualProduct:
  IsManualCalloff = True
End Function

'***********************************************************************************************
'****             This function returns the destination tank for the next Calloff           ****
'****     Uses IsNextDyelot to read from current program steps or from next scheduled job   ****
'***********************************************************************************************
Public Sub GetNextDestinationTank(ByVal ControlObject As Object)
On Error GoTo Err_GetNextDestinationTank


  Dim ControlCode As ControlCode: Set ControlCode = ControlObject
  With ControlCode
    'Get current program and step number
    Dim ProgNum As Long: ProgNum = .Parent.ProgramNumber
    Dim StepNum As Long: StepNum = .Parent.StepNumber
    
    'Get all program steps for this program
    Dim PrefixedSteps As String:  PrefixedSteps = .Parent.[_PrefixedSteps]
    Dim ProgramSteps() As String: ProgramSteps = Split(PrefixedSteps, vbCrLf)
    Dim Steps() As String
    'Use this to split out command parameters for a program step
    Dim parameters() As String
    Dim checkingfornotes() As String
    
    'Loop through each step starting from the beginning of the program
    Dim i As Integer
    Dim StartChecking As Boolean, KPFound As Boolean
    StartChecking = False
    KPFound = False
    Dim KP1 As String
    For i = LBound(ProgramSteps) To UBound(ProgramSteps)
      'Ignore step 0
      If i > 0 Then
        Steps = Split(ProgramSteps(i), Chr(255))
        If UBound(Steps) >= 1 Then
          If (Steps(0) = ProgNum) And (Steps(1) = StepNum - 1) Then StartChecking = True
          If StartChecking And (Not KPFound) Then
            Select Case UCase(Left(Trim(Steps(5)), 2))
    
              Case "AP"
                DestinationTank = 1
                Exit For
                
              Case "KP"
                'only find first kp after la
                KPFound = True
                KP1 = Steps(5)
                Exit For
                     
              Case "RP"
                DestinationTank = 2
                Exit For
                
              End Select
          
          End If
        End If
      End If
    Next i
    
    'DestinationTank = 0
    If Len(KP1) > 0 Then
      checkingfornotes = Split(KP1, "'")
      parameters = Split(checkingfornotes(0), ",")  '0 based array
      If UBound(parameters) >= 5 Then
        '0 = command '1 = Level, 2 = desired temp, 3 = mixing time, 4 = mixer on or off, 5 = DispenseTank
        If IsNumeric(parameters(1)) And IsNumeric(parameters(2)) And IsNumeric(parameters(3)) And _
                                        IsNumeric(parameters(4)) And IsNumeric(parameters(5)) Then
          DestinationTank = CLng(parameters(5))
        End If
      End If
    End If
    
  End With
  
  
  
'Avoid Error Handler
  Exit Sub
  
Err_GetNextDestinationTank:

End Sub

'***********************************************************************************************
'****                   This function returns the time before next transfer                 ****
'****                   based on the destination tank declared in KP command                ****
'***********************************************************************************************
Public Sub GetTimeBeforeTransfer(ByVal ControlObject As Object, Destination As Long)
On Error GoTo err_gettimebeforetransfer
  pTimeToTransfer = 0 'Return this value
  Dim ControlCode As ControlCode: Set ControlCode = ControlObject
  With ControlCode
  
    'Get current program and step number
    Dim ProgNum As Long: ProgNum = .Parent.ProgramNumber
    Dim StepNum As Long: StepNum = .Parent.StepNumber

    'Get all program steps for this program
    Dim PrefixedSteps As String: PrefixedSteps = .Parent.[_PrefixedSteps]
    Dim ProgramSteps() As String: ProgramSteps = Split(PrefixedSteps, vbCrLf)
    Dim Step() As String, Command As String
    
    'Loop through each step starting from the beginning of the program
    Dim i As Integer
    Dim StartChecking As Boolean: StartChecking = False

    'Only check for a destination transfer if Destination Tank is set in controlcode
    If Destination = 1 Then
      'Look for the next AT or AC command
      For i = LBound(ProgramSteps) To UBound(ProgramSteps)
        Step = Split(ProgramSteps(i), Chr(255))
        If UBound(Step) >= 1 Then
          If (Step(0) = ProgNum) And (Step(1) = StepNum - 1) Then StartChecking = True
          If StartChecking Then
            Command = Left(Trim(Step(5)), 2)
            If Not ((Command = "AT") Or (Command = "AC")) Then _
              pTimeToTransfer = pTimeToTransfer + CLng(Step(3))
            If (Command = "AT") Or (Command = "AC") Then Exit For
          End If
        End If
        'If i = UBound(ProgramSteps) Then pTimeToTransfer = 0
      Next i
    ElseIf Destination = 2 Then
      'Look for the next RT command
      For i = LBound(ProgramSteps) To UBound(ProgramSteps)
        Step = Split(ProgramSteps(i), Chr(255))
        If UBound(Step) >= 1 Then
          If (Step(0) = ProgNum) And (Step(1) = StepNum - 1) Then StartChecking = True
          If StartChecking Then
            Command = Left(Trim(Step(5)), 2)
            If Not (Command = "RT") Then pTimeToTransfer = pTimeToTransfer + CLng(Step(3))
            If (Command = "RT") Then Exit For
          End If
        End If
        'If i = UBound(ProgramSteps) Then pTimeToTransfer = 0
      Next i
    End If
    
  End With
  
  
  TimeToTransferTimer.TimeRemaining = pTimeToTransfer * 60
'Avoid Error Handler
  Exit Sub
  
err_gettimebeforetransfer:
  pTimeToTransfer = 0
  TimeToTransferTimer.TimeRemaining = 0
End Sub

'***********************************************************************************************
'******                Update the Dyelot & Redye from the local database                  ******
'***********************************************************************************************
Public Sub GetNextDyelotRedye()
On Error GoTo Err_GetDyelotRedye

'Set default values
  pDyelot = ""
  pRedye = 0

'Create an ADO connection to the local database
  ' get current job (dyelot and redye) from the local database instead of from parent.job
  ' carryover from PlantExplorer @1 bug
    Dim localConnection As ADODB.Connection15
    Set localConnection = New ADODB.Connection
    localConnection.Open "BatchDyeing"
    
    Dim rsDyelot As ADODB.Recordset15
    Set rsDyelot = localConnection.Execute("SELECT * FROM Dyelots WHERE State is null order by StartTime")
    rsDyelot.MoveFirst
  
    If Not rsDyelot.EOF Then
      'Set Dyelot and Redye - these must be non-null because they are the primary key
      pDyelot = CStr(rsDyelot!Dyelot)
      pRedye = CLng(rsDyelot!Redye)
    End If

'Tidy up
    localConnection.Close
    Set localConnection = Nothing


'Avoid Error Handler
    Exit Sub

Err_GetDyelotRedye:

End Sub
Public Sub LoadNextRecipeSteps()
On Error GoTo Err_LoadNextRecipeSteps

'Load the local database values for current dyelot & redye
  GetNextDyelotRedye
  
'Check that Dyelot & Redye are set. If dyelot is blank there is a empty space " " not ""
  If pDyelot = " " Then Exit Sub

'reset current record
  pCurrentRecipe = False
'Reset records
  pRecords = 0
'clear collection
  CurrentrecipeRows.Clear

'Create ADO objects
  Dim con As ADODB.Connection
  Dim rs As ADODB.Recordset15
  Dim sql As String

  Set con = New ADODB.Connection
  Set rs = New ADODB.Recordset
  sql = "SELECT * FROM DyelotsBulkedRecipe WHERE BATCH_NUMBER='" & pDyelot & "' AND " & "Redye=" & pRedye
  
'Open the connection and get the data
  con.Open "BatchDyeingCentral", "sa", ""
  rs.Open sql, con, adOpenStatic, adLockReadOnly
 
'Make sure the recordset is fully populated then set records variable
 If rs.RecordCount > 0 Then
  rs.MoveLast: rs.MoveFirst
  pRecords = rs.RecordCount
  pNextRecipe = True
 Else
  pRecords = 0
  pCollectionRows = 0
  pNextRecipe = False
  Exit Sub
 End If
'Populate the recipe row collection
  Dim newRecipeRow As acRecipeRow
  Do While Not rs.EOF
    Set newRecipeRow = New acRecipeRow
    newRecipeRow.Fill rs
    CurrentrecipeRows.Add newRecipeRow
    rs.MoveNext
  Loop
   
 'set the number of collection rows just for info
   pCollectionRows = CurrentrecipeRows.Count
   
'FOR DEBUG ONLY - output basic step info
'  Dim message As String
'  rs.MoveFirst
'  Do While Not rs.EOF
'    message = rs("ID") & "  " & rs("INGREDIENT_ID") & "  " & rs("INGREDIENT_DESC")
'    Debug.Print message
'    rs.MoveNext
'  Loop
'  rs.MoveFirst
  
'Clean up
  con.Close
  Set con = Nothing
  Set rs = Nothing

'Avoid Error Handler
  Exit Sub

Err_LoadNextRecipeSteps:
  MsgBox Err.Description
  Set con = Nothing
  Set rs = Nothing
End Sub





  VERSION 1.0 CLASS
'***********************************************************************************************
'******      Always use Connection15 and Recordset15 to avoid ADO version nonsense        ******
'***********************************************************************************************
'Table populated with Central DyelotsBulkedRecipe for current dyelot
  Private pCurrentRecipe As Boolean
  Private pNextRecipe As Boolean

'Private storage for properties
  Private pDyelot As String
  Private pRedye As Long
  
'for the collection
  Dim pRecords As Integer
  Dim pCollectionRows As Integer
  Dim CurrentrecipeRows As acRecipeRowCollection

'Private values - KP command parameters
  Private pTimeToTransfer As Long
  Public TimeToTransferTimer As New acTimer
  Public WaitingAtTransferTimer As New acTimerUp
  Public Tank1Calloff As Long
  Public DestinationTank As Long
  Public IsManualCalloff As Boolean
  Public PreviewStatus As String
  
  'initialize
  Private Sub Class_Initialize()
    Set CurrentrecipeRows = New acRecipeRowCollection
  End Sub
  Public Property Get Records() As Integer
    Records = pRecords
  End Property
  Public Property Get CollectionRows() As Integer
    CollectionRows = pCollectionRows
  End Property
  Public Property Get CurrentRecipe() As Boolean
    CurrentRecipe = pCurrentRecipe
  End Property
  Public Property Get NextRecipe() As Boolean
    NextRecipe = pNextRecipe
  End Property
  Public Property Get Job() As String
    Job = pDyelot
    If pRedye > 0 Then Job = pDyelot & "@" & pRedye
  End Property
  Public Property Get TimeToTransferString() As String
    TimeToTransferString = TimerString(TimeToTransferTimer.TimeRemaining)
  End Property
  Public Property Get WaitingAtTransferString() As String
    If WaitingAtTransferTimer.IsPaused Then
       WaitingAtTransferString = "0"
    Else
       WaitingAtTransferString = TimerString(WaitingAtTransferTimer.TimeElapsed)
    End If
  End Property
  Public Property Get WaitingAtTransferSeconds() As Integer
   WaitingAtTransferSeconds = WaitingAtTransferTimer.TimeElapsed
  End Property
  Public Property Get TimeToTransferSeconds() As Integer
  TimeToTransferSeconds = TimeToTransferTimer.TimeRemaining
  End Property
'***********************************************************************************************
'******                          Clear all relavent values                                ******
'***********************************************************************************************
Public Sub ResetAll()
  
  pDyelot = ""
  pRedye = 0
  DestinationTank = 0
  pCurrentRecipe = False
  pRecords = 0
  CurrentrecipeRows.Clear
  pCollectionRows = CurrentrecipeRows.Count
  pNextRecipe = False
  
End Sub

'***********************************************************************************************
'******                Update the Dyelot & Redye from the local database                  ******
'***********************************************************************************************
Public Sub GetDyelotRedye()
On Error GoTo Err_GetDyelotRedye

'Set default values
  pDyelot = ""
  pRedye = 0

'Create an ADO connection to the local database
  ' get current job (dyelot and redye) from the local database instead of from parent.job
  ' carryover from PlantExplorer @1 bug
    Dim localConnection As ADODB.Connection15
    Set localConnection = New ADODB.Connection
    localConnection.Open "BatchDyeing"
    
    Dim rsDyelot As ADODB.Recordset15
    Set rsDyelot = localConnection.Execute("SELECT * FROM Dyelots WHERE State=2 order by StartTime")
    rsDyelot.MoveFirst
  
    If Not rsDyelot.EOF Then
      'Set Dyelot and Redye - these must be non-null because they are the primary key
      pDyelot = CStr(rsDyelot!Dyelot)
      pRedye = CLng(rsDyelot!Redye)
    End If

'Tidy up
    localConnection.Close
    Set localConnection = Nothing


'Avoid Error Handler
    Exit Sub

Err_GetDyelotRedye:

End Sub

Public Sub LoadRecipeSteps()
On Error GoTo Err_LoadRecipeSteps

'Load the local database values for current dyelot & redye
  GetDyelotRedye
  
'Check that Dyelot & Redye are set. If dyelot is blank there is a empty space " " not ""
  If pDyelot = " " Then Exit Sub

'clear next recipe
  pNextRecipe = False
  
'Reset records
  pRecords = 0
'clear collection
  CurrentrecipeRows.Clear

'Create ADO objects
  Dim con As ADODB.Connection
  Dim rs As ADODB.Recordset15
  Dim sql As String

  Set con = New ADODB.Connection
  Set rs = New ADODB.Recordset
  sql = "SELECT * FROM DyelotsBulkedRecipe WHERE BATCH_NUMBER='" & pDyelot & "' AND " & "Redye=" & pRedye
  
'Open the connection and get the data
  con.Open "BatchDyeingCentral", "sa", ""
  rs.Open sql, con, adOpenStatic, adLockReadOnly
 
'Make sure the recordset is fully populated then set records variable
 If rs.RecordCount > 0 Then
  rs.MoveLast: rs.MoveFirst
  pRecords = rs.RecordCount
  pCurrentRecipe = True
 Else
  pRecords = 0
  pCollectionRows = 0
  pCurrentRecipe = False
  Exit Sub
 End If
'Populate the recipe row collection
  Dim newRecipeRow As acRecipeRow
  Do While Not rs.EOF
    Set newRecipeRow = New acRecipeRow
    newRecipeRow.Fill rs
    CurrentrecipeRows.Add newRecipeRow
    rs.MoveNext
  Loop
   
 'set the number of collection rows just for info
   pCollectionRows = CurrentrecipeRows.Count
   
'FOR DEBUG ONLY - output basic step info
  Dim message As String
  rs.MoveFirst
  Do While Not rs.EOF
    message = rs("ID") & "  " & rs("INGREDIENT_ID") & "  " & rs("INGREDIENT_DESC")
    Debug.Print message
    rs.MoveNext
  Loop
  rs.MoveFirst
  
'Clean up
  con.Close
  Set con = Nothing
  Set rs = Nothing

'Avoid Error Handler
  Exit Sub

Err_LoadRecipeSteps:
  MsgBox Err.Description
  Set con = Nothing
  Set rs = Nothing
End Sub



'***********************************************************************************************
'****           This function returns a True if the next Calloff has a manual product       ****
'****                   "Dispensed" Field value for product row is set to "M"               ****
'***********************************************************************************************
Public Function IsManualProduct(ByVal Calloff As Long, DDSEnabled As Long, DispenseEnabled As Long) As Boolean
On Error GoTo Err_IsManualProduct
  If Calloff = 0 Then Exit Function
  
'Default Return Value
  IsManualCalloff = True
  IsManualProduct = True
  Tank1Calloff = Calloff
  Dim r As Integer
  Dim RowCountCheck As Integer 'just in case the calloff does not exist in the recipe.
  
'Make sure we've got something to check
  If CurrentrecipeRows.Count > 0 Then
    For r = 1 To CurrentrecipeRows.Count
      If CurrentrecipeRows(r).AddNumber = Calloff Then
         If ((CurrentrecipeRows(r).IngredientID >= 1000) And DispenseEnabled <> 1) Or _
         ((CurrentrecipeRows(r).IngredientID < 1000) And DDSEnabled <> 1) Or _
         (CurrentrecipeRows(r).Dispensed = False) Then
            IsManualCalloff = True
            IsManualProduct = True
            Tank1Calloff = Calloff
            Exit Function
         End If
      Else
        RowCountCheck = RowCountCheck + 1
        If RowCountCheck = CurrentrecipeRows.Count Then
           'we checked every row and the calloff is not in the recipe has to be manual.
           'should never happen but just in case
            IsManualCalloff = True
            IsManualProduct = True
            Tank1Calloff = Calloff
            Exit Function
        End If
      End If
    Next
    
  Else
    Exit Function
  End If

'No manual products found
  IsManualCalloff = False
  IsManualProduct = False
  Tank1Calloff = 0
 
'Avoid Error Handler
  Exit Function

Err_IsManualProduct:
  IsManualCalloff = True
End Function

'***********************************************************************************************
'****             This function returns the destination tank for the next Calloff           ****
'****     Uses IsNextDyelot to read from current program steps or from next scheduled job   ****
'***********************************************************************************************
Public Sub GetNextDestinationTank(ByVal ControlObject As Object)
On Error GoTo Err_GetNextDestinationTank


  Dim ControlCode As ControlCode: Set ControlCode = ControlObject
  With ControlCode
    'Get current program and step number
    Dim ProgNum As Long: ProgNum = .Parent.ProgramNumber
    Dim StepNum As Long: StepNum = .Parent.StepNumber
    
    'Get all program steps for this program
    Dim PrefixedSteps As String:  PrefixedSteps = .Parent.[_PrefixedSteps]
    Dim ProgramSteps() As String: ProgramSteps = Split(PrefixedSteps, vbCrLf)
    Dim Steps() As String
    'Use this to split out command parameters for a program step
    Dim parameters() As String
    Dim checkingfornotes() As String
    
    'Loop through each step starting from the beginning of the program
    Dim i As Integer
    Dim StartChecking As Boolean, KPFound As Boolean
    StartChecking = False
    KPFound = False
    Dim KP1 As String
    For i = LBound(ProgramSteps) To UBound(ProgramSteps)
      'Ignore step 0
      If i > 0 Then
        Steps = Split(ProgramSteps(i), Chr(255))
        If UBound(Steps) >= 1 Then
          If (Steps(0) = ProgNum) And (Steps(1) = StepNum - 1) Then StartChecking = True
          If StartChecking And (Not KPFound) Then
            Select Case UCase(Left(Trim(Steps(5)), 2))
    
              Case "AP"
                DestinationTank = 1
                Exit For
                
              Case "KP"
                'only find first kp after la
                KPFound = True
                KP1 = Steps(5)
                Exit For
                     
              Case "RP"
                DestinationTank = 2
                Exit For
                
              End Select
          
          End If
        End If
      End If
    Next i
    
    'DestinationTank = 0
    If Len(KP1) > 0 Then
      checkingfornotes = Split(KP1, "'")
      parameters = Split(checkingfornotes(0), ",")  '0 based array
      If UBound(parameters) >= 5 Then
        '0 = command '1 = Level, 2 = desired temp, 3 = mixing time, 4 = mixer on or off, 5 = DispenseTank
        If IsNumeric(parameters(1)) And IsNumeric(parameters(2)) And IsNumeric(parameters(3)) And _
                                        IsNumeric(parameters(4)) And IsNumeric(parameters(5)) Then
          DestinationTank = CLng(parameters(5))
        End If
      End If
    End If
    
  End With
  
  
  
'Avoid Error Handler
  Exit Sub
  
Err_GetNextDestinationTank:

End Sub

'***********************************************************************************************
'****                   This function returns the time before next transfer                 ****
'****                   based on the destination tank declared in KP command                ****
'***********************************************************************************************
Public Sub GetTimeBeforeTransfer(ByVal ControlObject As Object, Destination As Long)
On Error GoTo err_gettimebeforetransfer
  pTimeToTransfer = 0 'Return this value
  Dim ControlCode As ControlCode: Set ControlCode = ControlObject
  With ControlCode
  
    'Get current program and step number
    Dim ProgNum As Long: ProgNum = .Parent.ProgramNumber
    Dim StepNum As Long: StepNum = .Parent.StepNumber

    'Get all program steps for this program
    Dim PrefixedSteps As String: PrefixedSteps = .Parent.[_PrefixedSteps]
    Dim ProgramSteps() As String: ProgramSteps = Split(PrefixedSteps, vbCrLf)
    Dim Step() As String, Command As String
    
    'Loop through each step starting from the beginning of the program
    Dim i As Integer
    Dim StartChecking As Boolean: StartChecking = False

    'Only check for a destination transfer if Destination Tank is set in controlcode
    If Destination = 1 Then
      'Look for the next AT or AC command
      For i = LBound(ProgramSteps) To UBound(ProgramSteps)
        Step = Split(ProgramSteps(i), Chr(255))
        If UBound(Step) >= 1 Then
          If (Step(0) = ProgNum) And (Step(1) = StepNum - 1) Then StartChecking = True
          If StartChecking Then
            Command = Left(Trim(Step(5)), 2)
            If Not ((Command = "AT") Or (Command = "AC")) Then _
              pTimeToTransfer = pTimeToTransfer + CLng(Step(3))
            If (Command = "AT") Or (Command = "AC") Then Exit For
          End If
        End If
        'If i = UBound(ProgramSteps) Then pTimeToTransfer = 0
      Next i
    ElseIf Destination = 2 Then
      'Look for the next RT command
      For i = LBound(ProgramSteps) To UBound(ProgramSteps)
        Step = Split(ProgramSteps(i), Chr(255))
        If UBound(Step) >= 1 Then
          If (Step(0) = ProgNum) And (Step(1) = StepNum - 1) Then StartChecking = True
          If StartChecking Then
            Command = Left(Trim(Step(5)), 2)
            If Not (Command = "RT") Then pTimeToTransfer = pTimeToTransfer + CLng(Step(3))
            If (Command = "RT") Then Exit For
          End If
        End If
        'If i = UBound(ProgramSteps) Then pTimeToTransfer = 0
      Next i
    End If
    
  End With
  
  
  TimeToTransferTimer.TimeRemaining = pTimeToTransfer * 60
'Avoid Error Handler
  Exit Sub
  
err_gettimebeforetransfer:
  pTimeToTransfer = 0
  TimeToTransferTimer.TimeRemaining = 0
End Sub

'***********************************************************************************************
'******                Update the Dyelot & Redye from the local database                  ******
'***********************************************************************************************
Public Sub GetNextDyelotRedye()
On Error GoTo Err_GetDyelotRedye

'Set default values
  pDyelot = ""
  pRedye = 0

'Create an ADO connection to the local database
  ' get current job (dyelot and redye) from the local database instead of from parent.job
  ' carryover from PlantExplorer @1 bug
    Dim localConnection As ADODB.Connection15
    Set localConnection = New ADODB.Connection
    localConnection.Open "BatchDyeing"
    
    Dim rsDyelot As ADODB.Recordset15
    Set rsDyelot = localConnection.Execute("SELECT * FROM Dyelots WHERE State is null order by StartTime")
    rsDyelot.MoveFirst
  
    If Not rsDyelot.EOF Then
      'Set Dyelot and Redye - these must be non-null because they are the primary key
      pDyelot = CStr(rsDyelot!Dyelot)
      pRedye = CLng(rsDyelot!Redye)
    End If

'Tidy up
    localConnection.Close
    Set localConnection = Nothing


'Avoid Error Handler
    Exit Sub

Err_GetDyelotRedye:

End Sub
Public Sub LoadNextRecipeSteps()
On Error GoTo Err_LoadNextRecipeSteps

'Load the local database values for current dyelot & redye
  GetNextDyelotRedye
  
'Check that Dyelot & Redye are set. If dyelot is blank there is a empty space " " not ""
  If pDyelot = " " Then Exit Sub

'reset current record
  pCurrentRecipe = False
'Reset records
  pRecords = 0
'clear collection
  CurrentrecipeRows.Clear

'Create ADO objects
  Dim con As ADODB.Connection
  Dim rs As ADODB.Recordset15
  Dim sql As String

  Set con = New ADODB.Connection
  Set rs = New ADODB.Recordset
  sql = "SELECT * FROM DyelotsBulkedRecipe WHERE BATCH_NUMBER='" & pDyelot & "' AND " & "Redye=" & pRedye
  
'Open the connection and get the data
  con.Open "BatchDyeingCentral", "sa", ""
  rs.Open sql, con, adOpenStatic, adLockReadOnly
 
'Make sure the recordset is fully populated then set records variable
 If rs.RecordCount > 0 Then
  rs.MoveLast: rs.MoveFirst
  pRecords = rs.RecordCount
  pNextRecipe = True
 Else
  pRecords = 0
  pCollectionRows = 0
  pNextRecipe = False
  Exit Sub
 End If
'Populate the recipe row collection
  Dim newRecipeRow As acRecipeRow
  Do While Not rs.EOF
    Set newRecipeRow = New acRecipeRow
    newRecipeRow.Fill rs
    CurrentrecipeRows.Add newRecipeRow
    rs.MoveNext
  Loop
   
 'set the number of collection rows just for info
   pCollectionRows = CurrentrecipeRows.Count
   
'FOR DEBUG ONLY - output basic step info
'  Dim message As String
'  rs.MoveFirst
'  Do While Not rs.EOF
'    message = rs("ID") & "  " & rs("INGREDIENT_ID") & "  " & rs("INGREDIENT_DESC")
'    Debug.Print message
'    rs.MoveNext
'  Loop
'  rs.MoveFirst
  
'Clean up
  con.Close
  Set con = Nothing
  Set rs = Nothing

'Avoid Error Handler
  Exit Sub

Err_LoadNextRecipeSteps:
  MsgBox Err.Description
  Set con = Nothing
  Set rs = Nothing
End Sub


#End If

#If 0 Then
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acDrugroomPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***********************************************************************************************
'******      Always use Connection15 and Recordset15 to avoid ADO version nonsense        ******
'***********************************************************************************************
'Table populated with Central DyelotsBulkedRecipe for current dyelot
  Private pCurrentRecipe As Boolean
  Private pNextRecipe As Boolean

'Private storage for properties
  Private pDyelot As String
  Private pRedye As Long
  
'for the collection
  Dim pRecords As Integer
  Dim pCollectionRows As Integer
  Dim CurrentrecipeRows As acRecipeRowCollection

'Private values - KP command parameters
  Private pTimeToTransfer As Long
  Public TimeToTransferTimer As New acTimer
  Public WaitingAtTransferTimer As New acTimerUp
  Public Tank1Calloff As Long
  Public DestinationTank As Long
  Public IsManualCalloff As Boolean
  Public PreviewStatus As String
  
  'initialize
  Private Sub Class_Initialize()
    Set CurrentrecipeRows = New acRecipeRowCollection
  End Sub




'***********************************************************************************************
'******                Update the Dyelot & Redye from the local database                  ******
'***********************************************************************************************
Public Sub GetDyelotRedye()
On Error GoTo Err_GetDyelotRedye

'Set default values
  pDyelot = ""
  pRedye = 0

'Create an ADO connection to the local database
  ' get current job (dyelot and redye) from the local database instead of from parent.job
  ' carryover from PlantExplorer @1 bug
    Dim localConnection As ADODB.Connection15
    Set localConnection = New ADODB.Connection
    localConnection.Open "BatchDyeing"
    
    Dim rsDyelot As ADODB.Recordset15
    Set rsDyelot = localConnection.Execute("SELECT * FROM Dyelots WHERE State=2 order by StartTime")
    rsDyelot.MoveFirst
  
    If Not rsDyelot.EOF Then
      'Set Dyelot and Redye - these must be non-null because they are the primary key
      pDyelot = CStr(rsDyelot!Dyelot)
      pRedye = CLng(rsDyelot!Redye)
    End If

'Tidy up
    localConnection.Close
    Set localConnection = Nothing


'Avoid Error Handler
    Exit Sub

Err_GetDyelotRedye:

End Sub

Public Sub LoadRecipeSteps()
On Error GoTo Err_LoadRecipeSteps

'Load the local database values for current dyelot & redye
  GetDyelotRedye
  
'Check that Dyelot & Redye are set. If dyelot is blank there is a empty space " " not ""
  If pDyelot = " " Then Exit Sub

'clear next recipe
  pNextRecipe = False
  
'Reset records
  pRecords = 0
'clear collection
  CurrentrecipeRows.Clear

'Create ADO objects
  Dim con As ADODB.Connection
  Dim rs As ADODB.Recordset15
  Dim sql As String

  Set con = New ADODB.Connection
  Set rs = New ADODB.Recordset
  sql = "SELECT * FROM DyelotsBulkedRecipe WHERE BATCH_NUMBER='" & pDyelot & "' AND " & "Redye=" & pRedye
  
'Open the connection and get the data
  con.Open "BatchDyeingCentral", "sa", ""
  rs.Open sql, con, adOpenStatic, adLockReadOnly
 
'Make sure the recordset is fully populated then set records variable
 If rs.RecordCount > 0 Then
  rs.MoveLast: rs.MoveFirst
  pRecords = rs.RecordCount
  pCurrentRecipe = True
 Else
  pRecords = 0
  pCollectionRows = 0
  pCurrentRecipe = False
  Exit Sub
 End If
'Populate the recipe row collection
  Dim newRecipeRow As acRecipeRow
  Do While Not rs.EOF
    Set newRecipeRow = New acRecipeRow
    newRecipeRow.Fill rs
    CurrentrecipeRows.Add newRecipeRow
    rs.MoveNext
  Loop
   
 'set the number of collection rows just for info
   pCollectionRows = CurrentrecipeRows.Count
   
'FOR DEBUG ONLY - output basic step info
  Dim message As String
  rs.MoveFirst
  Do While Not rs.EOF
    message = rs("ID") & "  " & rs("INGREDIENT_ID") & "  " & rs("INGREDIENT_DESC")
    Debug.Print message
    rs.MoveNext
  Loop
  rs.MoveFirst
  
'Clean up
  con.Close
  Set con = Nothing
  Set rs = Nothing

'Avoid Error Handler
  Exit Sub

Err_LoadRecipeSteps:
  MsgBox Err.Description
  Set con = Nothing
  Set rs = Nothing
End Sub



'***********************************************************************************************
'****           This function returns a True if the next Calloff has a manual product       ****
'****                   "Dispensed" Field value for product row is set to "M"               ****
'***********************************************************************************************
Public Function IsManualProduct(ByVal Calloff As Long, DDSEnabled As Long, DispenseEnabled As Long) As Boolean
On Error GoTo Err_IsManualProduct
  If Calloff = 0 Then Exit Function
  
'Default Return Value
  IsManualCalloff = True
  IsManualProduct = True
  Tank1Calloff = Calloff
  Dim r As Integer
  Dim RowCountCheck As Integer 'just in case the calloff does not exist in the recipe.
  
'Make sure we've got something to check
  If CurrentrecipeRows.Count > 0 Then
    For r = 1 To CurrentrecipeRows.Count
      If CurrentrecipeRows(r).AddNumber = Calloff Then
         If ((CurrentrecipeRows(r).IngredientID >= 1000) And DispenseEnabled <> 1) Or _
         ((CurrentrecipeRows(r).IngredientID < 1000) And DDSEnabled <> 1) Or _
         (CurrentrecipeRows(r).Dispensed = False) Then
            IsManualCalloff = True
            IsManualProduct = True
            Tank1Calloff = Calloff
            Exit Function
         End If
      Else
        RowCountCheck = RowCountCheck + 1
        If RowCountCheck = CurrentrecipeRows.Count Then
           'we checked every row and the calloff is not in the recipe has to be manual.
           'should never happen but just in case
            IsManualCalloff = True
            IsManualProduct = True
            Tank1Calloff = Calloff
            Exit Function
        End If
      End If
    Next
    
  Else
    Exit Function
  End If

'No manual products found
  IsManualCalloff = False
  IsManualProduct = False
  Tank1Calloff = 0
 
'Avoid Error Handler
  Exit Function

Err_IsManualProduct:
  IsManualCalloff = True
End Function

'***********************************************************************************************
'****             This function returns the destination tank for the next Calloff           ****
'****     Uses IsNextDyelot to read from current program steps or from next scheduled job   ****
'***********************************************************************************************
Public Sub GetNextDestinationTank(ByVal ControlObject As Object)
On Error GoTo Err_GetNextDestinationTank


  Dim ControlCode As ControlCode: Set ControlCode = ControlObject
  With ControlCode
    'Get current program and step number
    Dim ProgNum As Long: ProgNum = .Parent.ProgramNumber
    Dim StepNum As Long: StepNum = .Parent.StepNumber
    
    'Get all program steps for this program
    Dim PrefixedSteps As String:  PrefixedSteps = .Parent.[_PrefixedSteps]
    Dim ProgramSteps() As String: ProgramSteps = Split(PrefixedSteps, vbCrLf)
    Dim Steps() As String
    'Use this to split out command parameters for a program step
    Dim parameters() As String
    Dim checkingfornotes() As String
    
    'Loop through each step starting from the beginning of the program
    Dim i As Integer
    Dim StartChecking As Boolean, KPFound As Boolean
    StartChecking = False
    KPFound = False
    Dim KP1 As String
    For i = LBound(ProgramSteps) To UBound(ProgramSteps)
      'Ignore step 0
      If i > 0 Then
        Steps = Split(ProgramSteps(i), Chr(255))
        If UBound(Steps) >= 1 Then
          If (Steps(0) = ProgNum) And (Steps(1) = StepNum - 1) Then StartChecking = True
          If StartChecking And (Not KPFound) Then
            Select Case UCase(Left(Trim(Steps(5)), 2))
    
              Case "AP"
                DestinationTank = 1
                Exit For
                
              Case "KP"
                'only find first kp after la
                KPFound = True
                KP1 = Steps(5)
                Exit For
                     
              Case "RP"
                DestinationTank = 2
                Exit For
                
              End Select
          
          End If
        End If
      End If
    Next i
    
    'DestinationTank = 0
    If Len(KP1) > 0 Then
      checkingfornotes = Split(KP1, "'")
      parameters = Split(checkingfornotes(0), ",")  '0 based array
      If UBound(parameters) >= 5 Then
        '0 = command '1 = Level, 2 = desired temp, 3 = mixing time, 4 = mixer on or off, 5 = DispenseTank
        If IsNumeric(parameters(1)) And IsNumeric(parameters(2)) And IsNumeric(parameters(3)) And _
                                        IsNumeric(parameters(4)) And IsNumeric(parameters(5)) Then
          DestinationTank = CLng(parameters(5))
        End If
      End If
    End If
    
  End With
  
  
  
'Avoid Error Handler
  Exit Sub
  
Err_GetNextDestinationTank:

End Sub

'***********************************************************************************************
'****                   This function returns the time before next transfer                 ****
'****                   based on the destination tank declared in KP command                ****
'***********************************************************************************************
Public Sub GetTimeBeforeTransfer(ByVal ControlObject As Object, Destination As Long)
On Error GoTo err_gettimebeforetransfer
  pTimeToTransfer = 0 'Return this value
  Dim ControlCode As ControlCode: Set ControlCode = ControlObject
  With ControlCode
  
    'Get current program and step number
    Dim ProgNum As Long: ProgNum = .Parent.ProgramNumber
    Dim StepNum As Long: StepNum = .Parent.StepNumber

    'Get all program steps for this program
    Dim PrefixedSteps As String: PrefixedSteps = .Parent.[_PrefixedSteps]
    Dim ProgramSteps() As String: ProgramSteps = Split(PrefixedSteps, vbCrLf)
    Dim Step() As String, Command As String
    
    'Loop through each step starting from the beginning of the program
    Dim i As Integer
    Dim StartChecking As Boolean: StartChecking = False

    'Only check for a destination transfer if Destination Tank is set in controlcode
    If Destination = 1 Then
      'Look for the next AT or AC command
      For i = LBound(ProgramSteps) To UBound(ProgramSteps)
        Step = Split(ProgramSteps(i), Chr(255))
        If UBound(Step) >= 1 Then
          If (Step(0) = ProgNum) And (Step(1) = StepNum - 1) Then StartChecking = True
          If StartChecking Then
            Command = Left(Trim(Step(5)), 2)
            If Not ((Command = "AT") Or (Command = "AC")) Then _
              pTimeToTransfer = pTimeToTransfer + CLng(Step(3))
            If (Command = "AT") Or (Command = "AC") Then Exit For
          End If
        End If
        'If i = UBound(ProgramSteps) Then pTimeToTransfer = 0
      Next i
    ElseIf Destination = 2 Then
      'Look for the next RT command
      For i = LBound(ProgramSteps) To UBound(ProgramSteps)
        Step = Split(ProgramSteps(i), Chr(255))
        If UBound(Step) >= 1 Then
          If (Step(0) = ProgNum) And (Step(1) = StepNum - 1) Then StartChecking = True
          If StartChecking Then
            Command = Left(Trim(Step(5)), 2)
            If Not (Command = "RT") Then pTimeToTransfer = pTimeToTransfer + CLng(Step(3))
            If (Command = "RT") Then Exit For
          End If
        End If
        'If i = UBound(ProgramSteps) Then pTimeToTransfer = 0
      Next i
    End If
    
  End With
  
  
  TimeToTransferTimer.TimeRemaining = pTimeToTransfer * 60
'Avoid Error Handler
  Exit Sub
  
err_gettimebeforetransfer:
  pTimeToTransfer = 0
  TimeToTransferTimer.TimeRemaining = 0
End Sub

'***********************************************************************************************
'******                Update the Dyelot & Redye from the local database                  ******
'***********************************************************************************************
Public Sub GetNextDyelotRedye()
On Error GoTo Err_GetDyelotRedye

'Set default values
  pDyelot = ""
  pRedye = 0

'Create an ADO connection to the local database
  ' get current job (dyelot and redye) from the local database instead of from parent.job
  ' carryover from PlantExplorer @1 bug
    Dim localConnection As ADODB.Connection15
    Set localConnection = New ADODB.Connection
    localConnection.Open "BatchDyeing"
    
    Dim rsDyelot As ADODB.Recordset15
    Set rsDyelot = localConnection.Execute("SELECT * FROM Dyelots WHERE State is null order by StartTime")
    rsDyelot.MoveFirst
  
    If Not rsDyelot.EOF Then
      'Set Dyelot and Redye - these must be non-null because they are the primary key
      pDyelot = CStr(rsDyelot!Dyelot)
      pRedye = CLng(rsDyelot!Redye)
    End If

'Tidy up
    localConnection.Close
    Set localConnection = Nothing


'Avoid Error Handler
    Exit Sub

Err_GetDyelotRedye:

End Sub
Public Sub LoadNextRecipeSteps()
On Error GoTo Err_LoadNextRecipeSteps

'Load the local database values for current dyelot & redye
  GetNextDyelotRedye
  
'Check that Dyelot & Redye are set. If dyelot is blank there is a empty space " " not ""
  If pDyelot = " " Then Exit Sub

'reset current record
  pCurrentRecipe = False
'Reset records
  pRecords = 0
'clear collection
  CurrentrecipeRows.Clear

'Create ADO objects
  Dim con As ADODB.Connection
  Dim rs As ADODB.Recordset15
  Dim sql As String

  Set con = New ADODB.Connection
  Set rs = New ADODB.Recordset
  sql = "SELECT * FROM DyelotsBulkedRecipe WHERE BATCH_NUMBER='" & pDyelot & "' AND " & "Redye=" & pRedye
  
'Open the connection and get the data
  con.Open "BatchDyeingCentral", "sa", ""
  rs.Open sql, con, adOpenStatic, adLockReadOnly
 
'Make sure the recordset is fully populated then set records variable
 If rs.RecordCount > 0 Then
  rs.MoveLast: rs.MoveFirst
  pRecords = rs.RecordCount
  pNextRecipe = True
 Else
  pRecords = 0
  pCollectionRows = 0
  pNextRecipe = False
  Exit Sub
 End If
'Populate the recipe row collection
  Dim newRecipeRow As acRecipeRow
  Do While Not rs.EOF
    Set newRecipeRow = New acRecipeRow
    newRecipeRow.Fill rs
    CurrentrecipeRows.Add newRecipeRow
    rs.MoveNext
  Loop
   
 'set the number of collection rows just for info
   pCollectionRows = CurrentrecipeRows.Count
   
'FOR DEBUG ONLY - output basic step info
'  Dim message As String
'  rs.MoveFirst
'  Do While Not rs.EOF
'    message = rs("ID") & "  " & rs("INGREDIENT_ID") & "  " & rs("INGREDIENT_DESC")
'    Debug.Print message
'    rs.MoveNext
'  Loop
'  rs.MoveFirst
  
'Clean up
  con.Close
  Set con = Nothing
  Set rs = Nothing

'Avoid Error Handler
  Exit Sub

Err_LoadNextRecipeSteps:
  MsgBox Err.Description
  Set con = Nothing
  Set rs = Nothing
End Sub


#End If

#If 0 Then
  ' 2nd copy

  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acDrugroomPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***********************************************************************************************
'******      Always use Connection15 and Recordset15 to avoid ADO version nonsense        ******
'***********************************************************************************************
'Table populated with Central DyelotsBulkedRecipe for current dyelot
  Private pCurrentRecipe As Boolean
  Private pNextRecipe As Boolean

'Private storage for properties
  Private pDyelot As String
  Private pRedye As Long
  
'for the collection
  Dim pRecords As Integer
  Dim pCollectionRows As Integer
  Dim CurrentrecipeRows As acRecipeRowCollection

'Private values - KP command parameters
  Private pTimeToTransfer As Long
  Public TimeToTransferTimer As New acTimer
  Public WaitingAtTransferTimer As New acTimerUp
  Public Tank1Calloff As Long
  Public DestinationTank As Long
  Public IsManualCalloff As Boolean
  Public PreviewStatus As String
  
  'initialize
  Private Sub Class_Initialize()
    Set CurrentrecipeRows = New acRecipeRowCollection
  End Sub
  Public Property Get Records() As Integer
    Records = pRecords
  End Property
  Public Property Get CollectionRows() As Integer
    CollectionRows = pCollectionRows
  End Property
  Public Property Get CurrentRecipe() As Boolean
    CurrentRecipe = pCurrentRecipe
  End Property
  Public Property Get NextRecipe() As Boolean
    NextRecipe = pNextRecipe
  End Property
  Public Property Get Job() As String
    Job = pDyelot
    If pRedye > 0 Then Job = pDyelot & "@" & pRedye
  End Property
  Public Property Get TimeToTransferString() As String
    TimeToTransferString = TimerString(TimeToTransferTimer.TimeRemaining)
  End Property
  Public Property Get WaitingAtTransferString() As String
    If WaitingAtTransferTimer.IsPaused Then
       WaitingAtTransferString = "0"
    Else
       WaitingAtTransferString = TimerString(WaitingAtTransferTimer.TimeElapsed)
    End If
  End Property
  Public Property Get WaitingAtTransferSeconds() As Integer
   WaitingAtTransferSeconds = WaitingAtTransferTimer.TimeElapsed
  End Property
  Public Property Get TimeToTransferSeconds() As Integer
  TimeToTransferSeconds = TimeToTransferTimer.TimeRemaining
  End Property
'***********************************************************************************************
'******                          Clear all relavent values                                ******
'***********************************************************************************************
Public Sub ResetAll()
  
  pDyelot = ""
  pRedye = 0
  DestinationTank = 0
  pCurrentRecipe = False
  pRecords = 0
  CurrentrecipeRows.Clear
  pCollectionRows = CurrentrecipeRows.Count
  pNextRecipe = False
  
End Sub

'***********************************************************************************************
'******                Update the Dyelot & Redye from the local database                  ******
'***********************************************************************************************
Public Sub GetDyelotRedye()
On Error GoTo Err_GetDyelotRedye

'Set default values
  pDyelot = ""
  pRedye = 0

'Create an ADO connection to the local database
  ' get current job (dyelot and redye) from the local database instead of from parent.job
  ' carryover from PlantExplorer @1 bug
    Dim localConnection As ADODB.Connection15
    Set localConnection = New ADODB.Connection
    localConnection.Open "BatchDyeing"
    
    Dim rsDyelot As ADODB.Recordset15
    Set rsDyelot = localConnection.Execute("SELECT * FROM Dyelots WHERE State=2 order by StartTime")
    rsDyelot.MoveFirst
  
    If Not rsDyelot.EOF Then
      'Set Dyelot and Redye - these must be non-null because they are the primary key
      pDyelot = CStr(rsDyelot!Dyelot)
      pRedye = CLng(rsDyelot!Redye)
    End If

'Tidy up
    localConnection.Close
    Set localConnection = Nothing


'Avoid Error Handler
    Exit Sub

Err_GetDyelotRedye:

End Sub

Public Sub LoadRecipeSteps()
On Error GoTo Err_LoadRecipeSteps

'Load the local database values for current dyelot & redye
  GetDyelotRedye
  
'Check that Dyelot & Redye are set. If dyelot is blank there is a empty space " " not ""
  If pDyelot = " " Then Exit Sub

'clear next recipe
  pNextRecipe = False
  
'Reset records
  pRecords = 0
'clear collection
  CurrentrecipeRows.Clear

'Create ADO objects
  Dim con As ADODB.Connection
  Dim rs As ADODB.Recordset15
  Dim sql As String

  Set con = New ADODB.Connection
  Set rs = New ADODB.Recordset
  sql = "SELECT * FROM DyelotsBulkedRecipe WHERE BATCH_NUMBER='" & pDyelot & "' AND " & "Redye=" & pRedye
  
'Open the connection and get the data
  con.Open "BatchDyeingCentral", "sa", ""
  rs.Open sql, con, adOpenStatic, adLockReadOnly
 
'Make sure the recordset is fully populated then set records variable
 If rs.RecordCount > 0 Then
  rs.MoveLast: rs.MoveFirst
  pRecords = rs.RecordCount
  pCurrentRecipe = True
 Else
  pRecords = 0
  pCollectionRows = 0
  pCurrentRecipe = False
  Exit Sub
 End If
'Populate the recipe row collection
  Dim newRecipeRow As acRecipeRow
  Do While Not rs.EOF
    Set newRecipeRow = New acRecipeRow
    newRecipeRow.Fill rs
    CurrentrecipeRows.Add newRecipeRow
    rs.MoveNext
  Loop
   
 'set the number of collection rows just for info
   pCollectionRows = CurrentrecipeRows.Count
   
'FOR DEBUG ONLY - output basic step info
  Dim message As String
  rs.MoveFirst
  Do While Not rs.EOF
    message = rs("ID") & "  " & rs("INGREDIENT_ID") & "  " & rs("INGREDIENT_DESC")
    Debug.Print message
    rs.MoveNext
  Loop
  rs.MoveFirst
  
'Clean up
  con.Close
  Set con = Nothing
  Set rs = Nothing

'Avoid Error Handler
  Exit Sub

Err_LoadRecipeSteps:
  MsgBox Err.Description
  Set con = Nothing
  Set rs = Nothing
End Sub



'***********************************************************************************************
'****           This function returns a True if the next Calloff has a manual product       ****
'****                   "Dispensed" Field value for product row is set to "M"               ****
'***********************************************************************************************
Public Function IsManualProduct(ByVal Calloff As Long, DDSEnabled As Long, DispenseEnabled As Long) As Boolean
On Error GoTo Err_IsManualProduct
  If Calloff = 0 Then Exit Function
  
'Default Return Value
  IsManualCalloff = True
  IsManualProduct = True
  Tank1Calloff = Calloff
  Dim r As Integer
  Dim RowCountCheck As Integer 'just in case the calloff does not exist in the recipe.
  
'Make sure we've got something to check
  If CurrentrecipeRows.Count > 0 Then
    For r = 1 To CurrentrecipeRows.Count
      If CurrentrecipeRows(r).AddNumber = Calloff Then
         If ((CurrentrecipeRows(r).IngredientID >= 1000) And DispenseEnabled <> 1) Or _
         ((CurrentrecipeRows(r).IngredientID < 1000) And DDSEnabled <> 1) Or _
         (CurrentrecipeRows(r).Dispensed = False) Then
            IsManualCalloff = True
            IsManualProduct = True
            Tank1Calloff = Calloff
            Exit Function
         End If
      Else
        RowCountCheck = RowCountCheck + 1
        If RowCountCheck = CurrentrecipeRows.Count Then
           'we checked every row and the calloff is not in the recipe has to be manual.
           'should never happen but just in case
            IsManualCalloff = True
            IsManualProduct = True
            Tank1Calloff = Calloff
            Exit Function
        End If
      End If
    Next
    
  Else
    Exit Function
  End If

'No manual products found
  IsManualCalloff = False
  IsManualProduct = False
  Tank1Calloff = 0
 
'Avoid Error Handler
  Exit Function

Err_IsManualProduct:
  IsManualCalloff = True
End Function

'***********************************************************************************************
'****             This function returns the destination tank for the next Calloff           ****
'****     Uses IsNextDyelot to read from current program steps or from next scheduled job   ****
'***********************************************************************************************
Public Sub GetNextDestinationTank(ByVal ControlObject As Object)
On Error GoTo Err_GetNextDestinationTank


  Dim ControlCode As ControlCode: Set ControlCode = ControlObject
  With ControlCode
    'Get current program and step number
    Dim ProgNum As Long: ProgNum = .Parent.ProgramNumber
    Dim StepNum As Long: StepNum = .Parent.StepNumber
    
    'Get all program steps for this program
    Dim PrefixedSteps As String:  PrefixedSteps = .Parent.[_PrefixedSteps]
    Dim ProgramSteps() As String: ProgramSteps = Split(PrefixedSteps, vbCrLf)
    Dim Steps() As String
    'Use this to split out command parameters for a program step
    Dim parameters() As String
    Dim checkingfornotes() As String
    
    'Loop through each step starting from the beginning of the program
    Dim i As Integer
    Dim StartChecking As Boolean, KPFound As Boolean
    StartChecking = False
    KPFound = False
    Dim KP1 As String
    For i = LBound(ProgramSteps) To UBound(ProgramSteps)
      'Ignore step 0
      If i > 0 Then
        Steps = Split(ProgramSteps(i), Chr(255))
        If UBound(Steps) >= 1 Then
          If (Steps(0) = ProgNum) And (Steps(1) = StepNum - 1) Then StartChecking = True
          If StartChecking And (Not KPFound) Then
            Select Case UCase(Left(Trim(Steps(5)), 2))
    
              Case "AP"
                DestinationTank = 1
                Exit For
                
              Case "KP"
                'only find first kp after la
                KPFound = True
                KP1 = Steps(5)
                Exit For
                     
              Case "RP"
                DestinationTank = 2
                Exit For
                
              End Select
          
          End If
        End If
      End If
    Next i
    
    'DestinationTank = 0
    If Len(KP1) > 0 Then
      checkingfornotes = Split(KP1, "'")
      parameters = Split(checkingfornotes(0), ",")  '0 based array
      If UBound(parameters) >= 5 Then
        '0 = command '1 = Level, 2 = desired temp, 3 = mixing time, 4 = mixer on or off, 5 = DispenseTank
        If IsNumeric(parameters(1)) And IsNumeric(parameters(2)) And IsNumeric(parameters(3)) And _
                                        IsNumeric(parameters(4)) And IsNumeric(parameters(5)) Then
          DestinationTank = CLng(parameters(5))
        End If
      End If
    End If
    
  End With
  
  
  
'Avoid Error Handler
  Exit Sub
  
Err_GetNextDestinationTank:

End Sub

'***********************************************************************************************
'****                   This function returns the time before next transfer                 ****
'****                   based on the destination tank declared in KP command                ****
'***********************************************************************************************
Public Sub GetTimeBeforeTransfer(ByVal ControlObject As Object, Destination As Long)
On Error GoTo err_gettimebeforetransfer
  pTimeToTransfer = 0 'Return this value
  Dim ControlCode As ControlCode: Set ControlCode = ControlObject
  With ControlCode
  
    'Get current program and step number
    Dim ProgNum As Long: ProgNum = .Parent.ProgramNumber
    Dim StepNum As Long: StepNum = .Parent.StepNumber

    'Get all program steps for this program
    Dim PrefixedSteps As String: PrefixedSteps = .Parent.[_PrefixedSteps]
    Dim ProgramSteps() As String: ProgramSteps = Split(PrefixedSteps, vbCrLf)
    Dim Step() As String, Command As String
    
    'Loop through each step starting from the beginning of the program
    Dim i As Integer
    Dim StartChecking As Boolean: StartChecking = False

    'Only check for a destination transfer if Destination Tank is set in controlcode
    If Destination = 1 Then
      'Look for the next AT or AC command
      For i = LBound(ProgramSteps) To UBound(ProgramSteps)
        Step = Split(ProgramSteps(i), Chr(255))
        If UBound(Step) >= 1 Then
          If (Step(0) = ProgNum) And (Step(1) = StepNum - 1) Then StartChecking = True
          If StartChecking Then
            Command = Left(Trim(Step(5)), 2)
            If Not ((Command = "AT") Or (Command = "AC")) Then _
              pTimeToTransfer = pTimeToTransfer + CLng(Step(3))
            If (Command = "AT") Or (Command = "AC") Then Exit For
          End If
        End If
        'If i = UBound(ProgramSteps) Then pTimeToTransfer = 0
      Next i
    ElseIf Destination = 2 Then
      'Look for the next RT command
      For i = LBound(ProgramSteps) To UBound(ProgramSteps)
        Step = Split(ProgramSteps(i), Chr(255))
        If UBound(Step) >= 1 Then
          If (Step(0) = ProgNum) And (Step(1) = StepNum - 1) Then StartChecking = True
          If StartChecking Then
            Command = Left(Trim(Step(5)), 2)
            If Not (Command = "RT") Then pTimeToTransfer = pTimeToTransfer + CLng(Step(3))
            If (Command = "RT") Then Exit For
          End If
        End If
        'If i = UBound(ProgramSteps) Then pTimeToTransfer = 0
      Next i
    End If
    
  End With
  
  
  TimeToTransferTimer.TimeRemaining = pTimeToTransfer * 60
'Avoid Error Handler
  Exit Sub
  
err_gettimebeforetransfer:
  pTimeToTransfer = 0
  TimeToTransferTimer.TimeRemaining = 0
End Sub

'***********************************************************************************************
'******                Update the Dyelot & Redye from the local database                  ******
'***********************************************************************************************
Public Sub GetNextDyelotRedye()
On Error GoTo Err_GetDyelotRedye

'Set default values
  pDyelot = ""
  pRedye = 0

'Create an ADO connection to the local database
  ' get current job (dyelot and redye) from the local database instead of from parent.job
  ' carryover from PlantExplorer @1 bug
    Dim localConnection As ADODB.Connection15
    Set localConnection = New ADODB.Connection
    localConnection.Open "BatchDyeing"
    
    Dim rsDyelot As ADODB.Recordset15
    Set rsDyelot = localConnection.Execute("SELECT * FROM Dyelots WHERE State is null order by StartTime")
    rsDyelot.MoveFirst
  
    If Not rsDyelot.EOF Then
      'Set Dyelot and Redye - these must be non-null because they are the primary key
      pDyelot = CStr(rsDyelot!Dyelot)
      pRedye = CLng(rsDyelot!Redye)
    End If

'Tidy up
    localConnection.Close
    Set localConnection = Nothing


'Avoid Error Handler
    Exit Sub

Err_GetDyelotRedye:

End Sub
Public Sub LoadNextRecipeSteps()
On Error GoTo Err_LoadNextRecipeSteps

'Load the local database values for current dyelot & redye
  GetNextDyelotRedye
  
'Check that Dyelot & Redye are set. If dyelot is blank there is a empty space " " not ""
  If pDyelot = " " Then Exit Sub

'reset current record
  pCurrentRecipe = False
'Reset records
  pRecords = 0
'clear collection
  CurrentrecipeRows.Clear

'Create ADO objects
  Dim con As ADODB.Connection
  Dim rs As ADODB.Recordset15
  Dim sql As String

  Set con = New ADODB.Connection
  Set rs = New ADODB.Recordset
  sql = "SELECT * FROM DyelotsBulkedRecipe WHERE BATCH_NUMBER='" & pDyelot & "' AND " & "Redye=" & pRedye
  
'Open the connection and get the data
  con.Open "BatchDyeingCentral", "sa", ""
  rs.Open sql, con, adOpenStatic, adLockReadOnly
 
'Make sure the recordset is fully populated then set records variable
 If rs.RecordCount > 0 Then
  rs.MoveLast: rs.MoveFirst
  pRecords = rs.RecordCount
  pNextRecipe = True
 Else
  pRecords = 0
  pCollectionRows = 0
  pNextRecipe = False
  Exit Sub
 End If
'Populate the recipe row collection
  Dim newRecipeRow As acRecipeRow
  Do While Not rs.EOF
    Set newRecipeRow = New acRecipeRow
    newRecipeRow.Fill rs
    CurrentrecipeRows.Add newRecipeRow
    rs.MoveNext
  Loop
   
 'set the number of collection rows just for info
   pCollectionRows = CurrentrecipeRows.Count
   
'FOR DEBUG ONLY - output basic step info
'  Dim message As String
'  rs.MoveFirst
'  Do While Not rs.EOF
'    message = rs("ID") & "  " & rs("INGREDIENT_ID") & "  " & rs("INGREDIENT_DESC")
'    Debug.Print message
'    rs.MoveNext
'  Loop
'  rs.MoveFirst
  
'Clean up
  con.Close
  Set con = Nothing
  Set rs = Nothing

'Avoid Error Handler
  Exit Sub

Err_LoadNextRecipeSteps:
  MsgBox Err.Description
  Set con = Nothing
  Set rs = Nothing
End Sub




#End If

End Class
