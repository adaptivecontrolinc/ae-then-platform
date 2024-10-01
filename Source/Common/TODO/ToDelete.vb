Public Class ToDelete


  'acAddRecipe
#If 0 Then
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acDrugroomPreviewold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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




  ' acAddRecipeOld:
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



  ' **************************
  ' acBlendControl.vb
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acBlendControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'Blend Control Code

Option Explicit

  Private BlendFactor As Long
  Private BlendDeadBand As Long
  Private ColdWaterTemp As Long
  Private HotWaterTemp As Long
  Private BlendMaxAdjustment As Long
  Private BlendMinAdjustment As Long
  
  Private BlendTargetTemp As Long
  Private BlendError As Long
  Private BlendAdjustment As Long
  Private BlendOutput As Long
  
  Private BlendTimer As New acTimer


Private Sub Class_Initialize()

'Default parameter values
  BlendFactor = 20
  BlendDeadBand = 10
  ColdWaterTemp = 500
  HotWaterTemp = 1200
  BlendMinAdjustment = 20
  BlendMaxAdjustment = 200
  
End Sub

Public Sub Params(ByVal FactorInTenthsPercent As Long, _
                  ByVal DeadBandInTenths As Long, _
                  ByVal MinAdjustmentInTenths As Long, _
                  ByVal MaxAdjustmentInTenths As Long, _
                  ByVal ColdWaterTempInTenths As Long, _
                  ByVal HotWaterTempInTenths As Long)

  BlendFactor = FactorInTenthsPercent
  If BlendFactor < 0 Then BlendFactor = 0
  If BlendFactor > 1000 Then BlendFactor = 1000
  
  BlendDeadBand = DeadBandInTenths
  If BlendDeadBand < 0 Then BlendDeadBand = 0
  If BlendDeadBand > 1000 Then BlendDeadBand = 1000

  ColdWaterTemp = ColdWaterTempInTenths
  If ColdWaterTemp < 320 Then ColdWaterTemp = 320
  If ColdWaterTemp > 1000 Then ColdWaterTemp = 1000
  
  HotWaterTemp = HotWaterTempInTenths
  If HotWaterTemp < ColdWaterTemp Then HotWaterTemp = ColdWaterTemp
  If HotWaterTemp > 1800 Then HotWaterTemp = 1800
  
  BlendMinAdjustment = MinAdjustmentInTenths
  If BlendMinAdjustment < 0 Then BlendMinAdjustment = 0
  If BlendMinAdjustment > 250 Then BlendMinAdjustment = 250
  
  BlendMaxAdjustment = MaxAdjustmentInTenths
  If BlendMaxAdjustment < 1 Then BlendMaxAdjustment = 1
  If BlendMaxAdjustment > 1000 Then BlendMaxAdjustment = 1000
  
End Sub

Public Sub Start(ByVal TargetTempInTenths As Long)

'Start timer and set target temp
  BlendTimer = 1
  BlendTargetTemp = TargetTempInTenths
  
'Calculate initial output
  Dim BlendOffset As Long, BlendRange As Long
  BlendOffset = TargetTempInTenths - ColdWaterTemp
  BlendRange = HotWaterTemp - ColdWaterTemp
  If BlendRange = 0 Then
    BlendOutput = 500
  Else
    BlendOutput = (BlendOffset * 1000) / BlendRange
  End If
  If BlendOutput < 0 Then BlendOutput = 0
  If BlendOutput > 1000 Then BlendOutput = 1000
  
End Sub

Public Sub Run(ByVal WaterTempInTenths As Long)

'Calculate Error
  BlendError = BlendTargetTemp - WaterTempInTenths

'Adjust output (once a second)
  If BlendTimer.Finished Then
    BlendTimer = 1
    If Abs(BlendError) > BlendDeadBand Then
      BlendAdjustment = (BlendError * BlendFactor) / 1000
      If BlendAdjustment >= 0 Then
        If BlendAdjustment < BlendMinAdjustment Then BlendAdjustment = BlendMinAdjustment
        If BlendAdjustment > BlendMaxAdjustment Then BlendAdjustment = BlendMaxAdjustment
      Else
        If -BlendAdjustment < BlendMinAdjustment Then BlendAdjustment = -BlendMinAdjustment
        If -BlendAdjustment > BlendMaxAdjustment Then BlendAdjustment = -BlendMaxAdjustment
      End If
      BlendOutput = BlendOutput + BlendAdjustment
    End If
  End If
  
'Limit output
  If BlendOutput < 0 Then BlendOutput = 0
  If BlendOutput > 1000 Then BlendOutput = 1000
  
End Sub

Public Property Get Output() As Long

  Output = BlendOutput

'Limit output
  If Output < 0 Then Output = 0
  If Output > 1000 Then Output = 1000

End Property







#End If


  'acDrugroomPreview:
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


  'acLidControl
#If 0 Then
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acLidControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum LidControlState
  Off
  ReleasePin
  OpenBand
  RaiseLid
  LidOpen
  LowerLid
  CloseBand
  EngagePin
End Enum
Public State As LidControlState
Public StateString As String
Public Timer As New acTimer

Private pLidClosed As Boolean
Private pLidOpen As Boolean

'===========================================================================================
Public Sub Run(MachineSafe As Boolean, LevelOkToOpen As Boolean, MainPump As Boolean, _
              LidClosedSwitch, LidLimitSwitch As Boolean, _
              CloseLidPB As Boolean, OpenLidPB As Boolean, _
              EmergencyStop As Boolean, _
              LockingBandOpenTime As Long, _
              LockingBandClosingTime As Long, _
              LidRaisingTime As Long)

  pLidClosed = LidClosedSwitch And LidLimitSwitch
  
  Select Case State
    
    Case Off
      StateString = ""
      'start open the lid if the machine is safe
      If OpenLidPB And MachineSafe And LevelOkToOpen And (Not MainPump) And (Not EmergencyStop) Then
         State = ReleasePin
      End If
      'lower the lid if its not already lower or locked.
      If CloseLidPB And (Not EmergencyStop) Then
        If LidClosedSwitch Then
          If Not LidLimitSwitch Then
            State = CloseBand
            Timer = LockingBandClosingTime
          Else
            State = EngagePin
          End If
        Else
          State = LowerLid
        End If
      End If

    Case ReleasePin
      StateString = "Release Pin " & TimerString(Timer.TimeRemaining)
      If LidLimitSwitch Then Timer = 5
      If Timer.Finished Then
         State = OpenBand
         Timer = LockingBandOpenTime
      End If
    
    Case OpenBand
      StateString = "Open Band " & TimerString(Timer.TimeRemaining)
      If Timer.Finished Then
         State = RaiseLid
         Timer = LidRaisingTime
      End If
    
    Case RaiseLid
      StateString = "Raise Lid " & TimerString(Timer.TimeRemaining)
      If Not OpenLidPB Then
        State = LidOpen
      End If
      If LidClosedSwitch Then Timer = LidRaisingTime
      If Timer.Finished Then
         State = LidOpen
      End If
    
    Case LidOpen
      StateString = "Lid Open "
      If OpenLidPB Then
         State = RaiseLid
         Timer = LidRaisingTime
      End If
      If CloseLidPB Then
        State = LowerLid
      End If
      
    Case LowerLid
      StateString = "Lower Lid " & TimerString(Timer.TimeRemaining)
      If (Not LidClosedSwitch) Then Timer = 3
      If Not CloseLidPB Then
        State = LidOpen
      End If
      If Timer.Finished Then
         State = CloseBand
         Timer = LockingBandClosingTime
      End If
          
    Case CloseBand
      StateString = "Close Band " & TimerString(Timer.TimeRemaining)
      If Timer.Finished Then
         State = EngagePin
      End If
          
     Case EngagePin
      StateString = "Raise Pin " & TimerString(Timer.TimeRemaining)
      If (Not LidLimitSwitch) Then Timer = 3
      If Timer.Finished Then
         State = Off
      End If
          
  End Select
End Sub
'===========================================================================================

Friend Property Get IsActive() As Boolean
  IsActive = (State <> Off)
End Property
Public Property Get IsReleasePin() As Boolean
  IsReleasePin = (State = ReleasePin) Or (State = OpenBand) Or (State = RaiseLid) Or (State = LidOpen) Or (State = LowerLid) Or (State = CloseBand)
End Property
Public Property Get IsOpenLockingBand() As Boolean
  IsOpenLockingBand = (State = OpenBand) Or (State = RaiseLid) Or (State = LidOpen) Or (State = LowerLid)
End Property
Public Property Get IsRaiseLid() As Boolean
  IsRaiseLid = (State = RaiseLid)
End Property
Public Property Get IsLowerLid() As Boolean
  IsLowerLid = (State = LowerLid)
End Property
Public Property Get IsCloseLockingBand() As Boolean
  IsCloseLockingBand = (State = CloseBand) Or (State = EngagePin) Or (State = Off)
End Property


#End If


  ' acManualAddDrain:
#If 0 Then
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acLidControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum LidControlState
  Off
  ReleasePin
  OpenBand
  RaiseLid
  LidOpen
  LowerLid
  CloseBand
  EngagePin
End Enum
Public State As LidControlState
Public StateString As String
Public Timer As New acTimer

Private pLidClosed As Boolean
Private pLidOpen As Boolean

'===========================================================================================
Public Sub Run(MachineSafe As Boolean, LevelOkToOpen As Boolean, MainPump As Boolean, _
              LidClosedSwitch, LidLimitSwitch As Boolean, _
              CloseLidPB As Boolean, OpenLidPB As Boolean, _
              EmergencyStop As Boolean, _
              LockingBandOpenTime As Long, _
              LockingBandClosingTime As Long, _
              LidRaisingTime As Long)

  pLidClosed = LidClosedSwitch And LidLimitSwitch
  
  Select Case State
    
    Case Off
      StateString = ""
      'start open the lid if the machine is safe
      If OpenLidPB And MachineSafe And LevelOkToOpen And (Not MainPump) And (Not EmergencyStop) Then
         State = ReleasePin
      End If
      'lower the lid if its not already lower or locked.
      If CloseLidPB And (Not EmergencyStop) Then
        If LidClosedSwitch Then
          If Not LidLimitSwitch Then
            State = CloseBand
            Timer = LockingBandClosingTime
          Else
            State = EngagePin
          End If
        Else
          State = LowerLid
        End If
      End If

    Case ReleasePin
      StateString = "Release Pin " & TimerString(Timer.TimeRemaining)
      If LidLimitSwitch Then Timer = 5
      If Timer.Finished Then
         State = OpenBand
         Timer = LockingBandOpenTime
      End If
    
    Case OpenBand
      StateString = "Open Band " & TimerString(Timer.TimeRemaining)
      If Timer.Finished Then
         State = RaiseLid
         Timer = LidRaisingTime
      End If
    
    Case RaiseLid
      StateString = "Raise Lid " & TimerString(Timer.TimeRemaining)
      If Not OpenLidPB Then
        State = LidOpen
      End If
      If LidClosedSwitch Then Timer = LidRaisingTime
      If Timer.Finished Then
         State = LidOpen
      End If
    
    Case LidOpen
      StateString = "Lid Open "
      If OpenLidPB Then
         State = RaiseLid
         Timer = LidRaisingTime
      End If
      If CloseLidPB Then
        State = LowerLid
      End If
      
    Case LowerLid
      StateString = "Lower Lid " & TimerString(Timer.TimeRemaining)
      If (Not LidClosedSwitch) Then Timer = 3
      If Not CloseLidPB Then
        State = LidOpen
      End If
      If Timer.Finished Then
         State = CloseBand
         Timer = LockingBandClosingTime
      End If
          
    Case CloseBand
      StateString = "Close Band " & TimerString(Timer.TimeRemaining)
      If Timer.Finished Then
         State = EngagePin
      End If
          
     Case EngagePin
      StateString = "Raise Pin " & TimerString(Timer.TimeRemaining)
      If (Not LidLimitSwitch) Then Timer = 3
      If Timer.Finished Then
         State = Off
      End If
          
  End Select
End Sub
'===========================================================================================

Friend Property Get IsActive() As Boolean
  IsActive = (State <> Off)
End Property
Public Property Get IsReleasePin() As Boolean
  IsReleasePin = (State = ReleasePin) Or (State = OpenBand) Or (State = RaiseLid) Or (State = LidOpen) Or (State = LowerLid) Or (State = CloseBand)
End Property
Public Property Get IsOpenLockingBand() As Boolean
  IsOpenLockingBand = (State = OpenBand) Or (State = RaiseLid) Or (State = LidOpen) Or (State = LowerLid)
End Property
Public Property Get IsRaiseLid() As Boolean
  IsRaiseLid = (State = RaiseLid)
End Property
Public Property Get IsLowerLid() As Boolean
  IsLowerLid = (State = LowerLid)
End Property
Public Property Get IsCloseLockingBand() As Boolean
  IsCloseLockingBand = (State = CloseBand) Or (State = EngagePin) Or (State = Off)
End Property


  ' ACManualAddFill

  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ACManualAddFill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===============================================================================================
'manual fill from add buttons
'===============================================================================================
'
  Option Explicit
  Public Enum ManualAddFillState
    Off
    Interlock
    FillingHot
    FillingCold
  End Enum
  Public State As ManualAddFillState
  

Public Sub Run(FillLevel As Long, AddLevel As Long, FillLevelDB As Long, FillHot As Boolean, _
              FillCold As Boolean)


    Select Case State
       
       Case Off
          If FillCold Or FillHot Then
             State = Interlock
          End If
          
       Case Interlock
          If FillHot Then
             State = FillingHot
          End If
          If FillCold Then
            State = FillingCold
          End If
          
        
       Case FillingHot
          If Not FillHot Then State = Off
          If (AddLevel >= ((FillLevel * 10) - FillLevelDB)) Then
             State = Off
             FillHot = False
          End If
         
        Case FillingCold
         If Not FillCold Then State = Off
          If (AddLevel >= ((FillLevel * 10) - FillLevelDB)) Then
             State = Off
             FillCold = False
          End If
          
    End Select
End Sub
Public Property Get IsFillingHot() As Boolean
  If (State = FillingHot) Then IsFillingHot = True
End Property
Public Property Get IsFillingCold() As Boolean
  If (State = FillingCold) Then IsFillingCold = True

End Property



  ' ACManualAddTransfer:
'===============================================================================================
'manual fill from add buttons
'===============================================================================================
'
  Option Explicit
  Public Enum ManualAddTransferState
    Off
    Interlock
    Transfer1
    Rinse
    Transfer2
    RinseToDrain
    Drain
    TurnOff
  End Enum
  Public State As ManualAddTransferState
  Public Timer As New acTimer
  

Public Sub Run(AddLevel As Long, AddTransferPB As Boolean, AddTransferFinished As Boolean, _
              TransferTimeBeforeRinse As Long, RinseTime As Long, TransferTimeAfterRinse As Long, RinseToDrainTime As Long, _
              DrainTime As Long)


    Select Case State
       
       Case Off
          If AddTransferPB Then
            State = Interlock
            AddTransferFinished = False
          End If
          
       Case Interlock
          If Not AddTransferPB Then State = Off
          State = Transfer1
          Timer = TransferTimeBeforeRinse
        
        Case Transfer1
          If Not AddTransferPB Then State = Off
          If AddLevel > 10 Then Timer = TransferTimeBeforeRinse
          If Timer.Finished Then
             State = Rinse
             Timer = RinseTime
          End If
          
        Case Rinse
          If Not AddTransferPB Then State = Off
          If Timer.Finished Then
             State = Transfer2
             Timer = TransferTimeAfterRinse
          End If
          
        Case Transfer2
          If Not AddTransferPB Then State = Off
          If AddLevel > 10 Then Timer = TransferTimeAfterRinse
          If Timer.Finished Then
             State = RinseToDrain
             Timer = RinseToDrainTime
          End If
        
        Case RinseToDrain
          If Not AddTransferPB Then State = Off
          If Timer.Finished Then
             State = Drain
             Timer = DrainTime
          End If
          
        Case Drain
          If Not AddTransferPB Then State = Off
          If AddLevel > 10 Then Timer = DrainTime
          If Timer.Finished Then
             State = TurnOff
             AddTransferFinished = True
          End If
          
        Case TurnOff
           If AddTransferPB = False Then
             AddTransferFinished = False
             State = Off
           End If
    
    End Select
End Sub
Public Property Get IsTransferring() As Boolean
 If (State = Transfer1) Or (State = Rinse) Or (State = Transfer2) Then IsTransferring = True
End Property
Public Property Get IsRinsing() As Boolean
  If (State = Rinse) Or (State = RinseToDrain) Then IsRinsing = True
End Property
Public Property Get IsDraining() As Boolean
  If (State = Drain) Or (State = RinseToDrain) Then IsDraining = True
End Property

End Class


#End If

End Class







