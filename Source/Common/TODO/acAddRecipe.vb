Public Class AddRecipeOld






#If 0 Then


'***********************************************************************************************
'******      Always use Connection15 and Recordset15 to avoid ADO version nonsense        ******
'***********************************************************************************************
'Table populated with Central DyelotsBulkedRecipe for current dyelot
  Private pRecipe As ADODB.Recordset15
  Private pCurrentRecipe As Boolean
  Private pNextRecipe As Boolean

'Private storage for properties
  Private pDyelot As String
  Private pRedye As Long
  Private pProgram As String
  
'Private values - KP command parameters
  Private pTank1FillLevel As Long
  Private pTank1DesiredTemp As Long
  Private pTank1MixTime As Long
  Private PTank1MixingOn As Long
  Private PTank1Calloff As Long
  Private PDispenseTank As Long
  
  Private pIsManualCalloff As Boolean
  Private pTimeToTransfer As Long
  Private pTimeToTransferTimer As New acTimer
  
  Private pProgramHasAnotherTank1 As Boolean
  Private pLookAheadStopFound As Boolean
  Private pLookAheadRepeatFound As Boolean
  Private pInsertedProgram As Long
  Public LARepeat As String
  
  Public Property Get Job() As String
    Job = pDyelot
    If pRedye > 0 Then Job = pDyelot & "@" & pRedye
  End Property
  
  Public Property Get TimeToTransfer() As Long
    TimeToTransfer = pTimeToTransfer
  End Property
  Public Property Let TimeToTransfer(Value As Long)
    TimeToTransfer = Value
  End Property
  Public Property Get TimeToTransferTimeRemaining() As Long
    TimeToTransferTimeRemaining = pTimeToTransferTimer.TimeRemaining
  End Property
  Public Property Get CurrentRecipe() As Boolean
    CurrentRecipe = pCurrentRecipe
  End Property

Public Property Get RecipeStepCount() As Integer
On Error GoTo Err_RecipeStepCount

   RecipeStepCount = 0
   If pRecipe Is Nothing Then Exit Property
   
   RecipeStepCount = pRecipe.RecordCount

'Avoid Error Handler
  Exit Property

Err_RecipeStepCount:
   RecipeStepCount = 0
End Property
  
'***********************************************************************************************
'******                          Clear all relavent values                                ******
'***********************************************************************************************
Public Sub ResetAll()
  Set pRecipe = Nothing
  pDyelot = ""
  pRedye = 0
  pProgram = ""
  PDispenseTank = 0
  
  
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

'***********************************************************************************************
'******            Get RecordSet of Recipe associated with current dyelot                 ******
'***********************************************************************************************
Public Sub LoadCurrentRecipe()
On Error GoTo Err_LoadCurrentRecipe

'Default Return Value
  pCurrentRecipe = False
  
'Load the local database values for current dyelot & redye
  GetDyelotRedye
  
'Check that Dyelot & Redye are set. If dyelot is blank there is a empty space " " not ""
  If pDyelot = " " Then Exit Sub
  
  
'We need to either create an OBDC link to the central database on each controller, and reference
'   that link in the centalConnection string, or hardcode (parameter/xml?) the connection as
'   below <needs testing on remote database>

'Required: Create an OBDC link on each controller to reference the BatchDyeingCentral database
'   Administrative Tools > ODBC Data Source Administrator > System DSN (tab)
'     Add: "BatchDyeingCentral: as SQL Server
'     Server: "Adaptive-Server" or "10.1.21.200"
'     SQL Server Authentication: LoginID: Adaptive  Password: Control
'     Change The Default Database To: "BatchDyeingCentral
'     Test Data Source (ensure working connection) Select OK

'Create an ADO connection to the central database
  Dim centalConnection As ADODB.Connection15
  Set centalConnection = New ADODB.Connection
  centalConnection.Open "BatchDyeingCentral", "Adaptive", "Control"
  
  Dim rsRecipe As ADODB.Recordset15
  Set rsRecipe = New ADODB.Recordset
  Dim sql As String
  sql = "SELECT * FROM DyelotsBulkedRecipe WHERE BATCH_NUMBER='" & pDyelot & "' AND Redye=" & pRedye
  rsRecipe.Open sql, centalConnection, adOpenStatic, adLockReadOnly
  If rsRecipe.RecordCount > 0 Then
   rsRecipe.MoveLast: rsRecipe.MoveFirst
  End If
'  Dim mikecount As String
'  mikecount = rsRecipe.RecordCount
  If Not rsRecipe.EOF Then
    Set pRecipe = rsRecipe
    pCurrentRecipe = True
  Else
    pCurrentRecipe = False
  
  End If
  
'Tidy up
  centalConnection.Close
  Set centalConnection = Nothing


'Avoid Error Handler
  Exit Sub
  
Err_LoadCurrentRecipe:
  MsgBox Err.Description

  Set localConnection = Nothing
  Set centalConnection = Nothing
  Set rsRecipe = Nothing
  pCurrentRecipe = False
End Sub

'***********************************************************************************************
'******    Get RecordSet of Recipe associated with next scheduled/unblocked dyelot        ******
'***********************************************************************************************
Public Function LoadNextRecipe() As Boolean
On Error GoTo Err_LoadNextRecipe

'Default Return Value
  pNextRecipe = False
  
'Create an ADO connection to the local database
  Dim localConnection As ADODB.Connection15
  Set localConnection = New ADODB.Connection
  localConnection.Open "BatchDyeing"
  
  Dim rsDyelot As ADODB.Recordset15
  Set rsDyelot = localConnection.Execute("SELECT * FROM Dyelots WHERE State Is Null ORDER BY StartTime")
  rsDyelot.MoveFirst

  If Not rsDyelot.EOF Then
    'Set Dyelot and Redye - these must be non-null because they are the primary key
    pDyelot = CStr(rsDyelot!Dyelot)
    pRedye = CLng(rsDyelot!Redye)
  End If

'Check that Dyelot & Redye are set
  If pDyelot = "" Then Exit Function
  
'Create an ADO connection to the central database
  Dim centalConnection As ADODB.Connection15
  Set centalConnection = New ADODB.Connection
  centalConnection.Open "BatchDyeingCentral", "Adaptive", "Control"
  
  Dim rsRecipe As ADODB.Recordset15
  Dim sql As String
  sql = "SELECT * FROM DyelotsBulkedRecipe WHERE BATCH_NUMBER='" & pDyelot & "' AND " & "Redye=" & pRedye
  Set rsRecipe = centalConnection.Execute(sql)
  If Not rsRecipe.EOF Then
    Set pRecipe = rsRecipe
  End If

'Tidy up
  localConnection.Close
  Set localConnection = Nothing

  centralConnection.Close
  Set centralConnection = Nothing
  
'everything ok
  pNextRecipe = True

'Avoid Error Handler
  Exit Function
  
Err_LoadNextRecipe:
  Set localConnection = Nothing
  Set centalConnection = Nothing
  Set rsRecipe = Nothing
  pNextRecipe = False
End Function

'***********************************************************************************************
'****           This function returns a True if the next Calloff has a manual product       ****
'****                   "Dispensed" Field value for product row is set to "M"               ****
'***********************************************************************************************
Public Function IsManualProduct(ByVal Calloff As Long) As Boolean
On Error GoTo Err_IsManualProduct

'Default Return Value
  pIsManualCalloff = True
  IsManualProduct = True
  
'Make sure we've got something to check
  If Not pRecipe.EOF Then
  
    pRecipe.MoveFirst
    Do While Not pRecipe.EOF
      If CLng(pRecipe!ADD_NUMBER) = Calloff Then
        If CStr(pRecipe!Dispensed) = "M" Then
          pIsManualCalloff = True
          IsManualProduct = True
          Exit Function
        End If
      End If
    Loop
    
  End If
  

'No manual products found
  pIsManualCalloff = False
  IsManualProduct = False
  
'Avoid Error Handler
  Exit Function

Err_IsManualProduct:
  pIsManualCalloff = True
End Function

'***********************************************************************************************
'****             This function returns the destination tank for the next Calloff           ****
'****     Uses IsNextDyelot to read from current program steps or from next scheduled job   ****
'***********************************************************************************************
Public Function GetNextDestinationTank(ByVal ControlObject As Object) As Long
On Error GoTo Err_GetNextDestinationTank

'Default Return Value
  GetNextDestinationTank = 0
  
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
                PDispenseTank = 1
                Exit For
                
              Case "KP"
                'only find first kp after la
                KPFound = True
                KP1 = Steps(5)
                Exit For
                     
              Case "RP"
                PDispenseTank = 2
                Exit For
                
              End Select
          
          End If
        End If
      End If
    Next i
    
    'Set KP1 Parameters - next prep for tank 1
    pTank1Enabled = False
    pProgramHasAnotherTank1 = False
    pTank1FillLevel = 0
    pTank1DesiredTemp = 0
    pTank1MixTime = 0
    PTank1MixingOn = 0
    PDispenseTank = 0
    If Len(KP1) > 0 Then
      pProgramHasAnotherTank1 = True
      pTank1Enabled = True
      checkingfornotes = Split(KP1, "'")
      parameters = Split(checkingfornotes(0), ",")  '0 based array
      If UBound(parameters) >= 5 Then
        '0 = command '1 = Level, 2 = desired temp, 3 = mixing time, 4 = mixer on or off, 5 = DispenseTank
        If IsNumeric(parameters(1)) And IsNumeric(parameters(2)) And IsNumeric(parameters(3)) And _
                                        IsNumeric(parameters(4)) And IsNumeric(parameters(5)) Then
          pTank1FillLevel = CLng(parameters(1)) * 10
          pTank1DesiredTemp = CLng(parameters(2)) * 10
          pTank1MixTime = CLng(parameters(3)) * 60
          PTank1MixingOn = CLng(parameters(4))
          PDispenseTank = CLng(parameters(5))
        End If
      End If
    End If
    
  End With
  
  GetNextDestinationTank = PDispenseTank
  
  
'Avoid Error Handler
  Exit Function
  
Err_GetNextDestinationTank:
  GetNextDestinationTank = 0
End Function



Public Function CheckExistingProgram(ByVal ControlObject As Object) As Long

End Function



'***********************************************************************************************
'****                   This function returns the time before next transfer                 ****
'****                   based on the destination tank declared in KP command                ****
'***********************************************************************************************
Public Sub GetTimeBeforeTransfer(ByVal ControlObject As Object)
On Error GoTo err_gettimebeforetransfer
  Dim TimeBeforeTransfer As Long: TimeBeforeTransfer = 0 'Return this value
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
    If .DestTank = 1 Then
      'Look for the next AT or AC command
      For i = LBound(ProgramSteps) To UBound(ProgramSteps)
        Step = Split(ProgramSteps(i), Chr(255))
        If UBound(Step) >= 1 Then
          If (Step(0) = ProgNum) And (Step(1) = StepNum - 1) Then StartChecking = True
          If StartChecking Then
            Command = Left(Trim(Step(5)), 2)
            If Not ((Command = "AT") Or (Command = "AC")) Then _
              TimeBeforeTransfer = TimeBeforeTransfer + CLng(Step(3))
            If (Command = "AT") Or (Command = "AC") Then Exit For
          End If
        End If
        If i = UBound(ProgramSteps) Then TimeBeforeTransfer = 0
      Next i
    ElseIf .DestTank = 2 Then
      'Look for the next RT command
      For i = LBound(ProgramSteps) To UBound(ProgramSteps)
        Step = Split(ProgramSteps(i), Chr(255))
        If UBound(Step) >= 1 Then
          If (Step(0) = ProgNum) And (Step(1) = StepNum - 1) Then StartChecking = True
          If StartChecking Then
            Command = Left(Trim(Step(5)), 2)
            If Not (Command = "RT") Then TimeBeforeTransfer = TimeBeforeTransfer + CLng(Step(3))
            If (Command = "RT") Then Exit For
          End If
        End If
        If i = UBound(ProgramSteps) Then TimeBeforeTransfer = 0
      Next i
    End If
    
  End With
  
  pTimeToTransfer = TimeBeforeTransfer
  pTimeToTransferTimer.TimeRemaining = pTimeToTransfer * 60

'Avoid Error Handler
  Exit Sub
  
err_gettimebeforetransfer:
  pTimeToTransfer = 0
End Sub

#End If
End Class
