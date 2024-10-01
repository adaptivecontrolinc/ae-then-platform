Public Class LA_VB6_Copy

#If 0 Then
  ' LA Version 2.4 20220214
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===============================================================================================
'LA - Look ahead command
'Version 2013-07-22
'===============================================================================================
Option Explicit
Implements ACCommand

'Main LA State Machine
Public Enum LAState
  LAOff
  LACheckExistingDyelot
  LACheckNextDyelot
  LAStart
  LARun
  LADone
End Enum
Public State As LAState
Public StatePrevious As LAState

'These command states will be transfered to the KP commands in the next program
Public KP1 As New acAddPrepare
Public KP1Calloff As Long

'Use this to check for next dyelot every 60 seconds
Public CheckNextDyelotTimer As New acTimer

Public IPKP1 As String
Public LARepeat As String

'Private storage for properties
Private pDyelot As String
Private pRedye As Long
Private pJob As String
Private pBlocked As Boolean
Private pProgram As String
Private pTank1Enabled As Boolean
Private pTank1FillLevel As Long
Private pTank1DesiredTemp As Long
Private pTank1MixTime As Long
Private PTank1MixingOn As Long
Private PTank1Calloff As Long
Private PDispenseTank As Long
Private pProgramHasAnotherTank1 As Boolean
Private pLookAheadStopFound As Boolean
Private pLookAheadRepeatFound As Boolean
Private pInsertedProgram As Long

Public Property Get Dyelot() As String
  Dyelot = pDyelot
End Property
Public Property Get Redye() As Long
  Redye = pRedye
End Property
Public Property Get Job() As String
  Job = Dyelot
  If Redye > 0 Then Job = Dyelot & "@" & Redye
End Property
Public Property Get Blocked() As Boolean
  Blocked = pBlocked
End Property
Public Property Get Program() As String
  Program = pProgram
End Property
Public Property Get Tank1Enabled() As Boolean
  Tank1Enabled = pTank1Enabled
End Property
Public Property Get Tank1FillLevel() As Long
  Tank1FillLevel = pTank1FillLevel
End Property
Public Property Get Tank1DesiredTemp() As Long
  Tank1DesiredTemp = pTank1DesiredTemp
End Property
Public Property Get Tank1MixTime() As Long
  Tank1MixTime = pTank1MixTime
End Property
Public Property Get Tank1MixingOn() As Long
  Tank1MixingOn = PTank1MixingOn
End Property
Public Property Get Tank1Calloff() As Long
  Tank1Calloff = PTank1Calloff
End Property
Public Property Get DispenseTank() As Long
  DispenseTank = PDispenseTank
End Property
Public Property Get ProgramHasAnotherTank1() As Boolean
  ProgramHasAnotherTank1 = pProgramHasAnotherTank1
End Property
Public Property Get LookAheadStopFound() As Boolean
  LookAheadStopFound = pLookAheadStopFound
End Property
Public Property Get LookAheadRepeatFound() As Boolean
  LookAheadRepeatFound = pLookAheadRepeatFound
End Property
Public Property Get InsertedProg() As Long
  InsertedProg = pInsertedProgram
End Property

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Name=Look Ahead\r\nMinutes=0\r\nHelp=Signals the dispenser to start dispensing products for the next lot."

'Reset state & variables
  ResetAll
  
'Early bind ControlCode for speed and convenience
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject

'Initialize
  State = LAOff
  StatePrevious = LAOff
  pTank1Enabled = False
  
  With ControlCode
    'Make sure the command is enabled
    If .Parameters_LookAheadEnabled = 1 Then
       State = LACheckExistingDyelot
       KP1Calloff = GetKPCalloff(ControlObject)
    End If
  End With
  
'Carry on my wayward son...
  StepOn = True

End Sub

Private Sub ResetAll()
'Reset command state & variables
  IPKP1 = ""
  LARepeat = ""
  pJob = ""
  pDyelot = ""
  pRedye = 0
  pBlocked = False
  pProgram = ""
  pLookAheadStopFound = False
  pLookAheadRepeatFound = False
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)

'If the command is not active then don't run this code
  If Not IsOn Then Exit Sub
  
'Early bind ControlCode for speed and convenience
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  
  With ControlCode
  
    Do
      'Stay in this loop until state changes have completed
      ' - stops odd effects from momentary state transitions
      StatePrevious = State
      Select Case State
        Case LAOff
                
        Case LACheckExistingDyelot
         'Check to see if there is another kp in this program.
          If CheckExistingProgram(ControlCode) Then
            If ProgramHasAnotherTank1 = True Then
              .LARequest = False
              State = LAStart
            Else
              .LARequest = True
              State = LACheckNextDyelot
            End If
          End If
          'stop trying to check existing program if we've found a LA repeat
          If pLookAheadRepeatFound Then Cancel
       
        'Check to see if a dyelot is schedule behind this one
        Case LACheckNextDyelot
          'Check every 60 seconds
          If CheckNextDyelotTimer.Finished Then
            CheckNextDyelotTimer = 60
            'Was blocked and ignoreblocked disabled - program may change due to blocked, so reset and recheck
            If Blocked Then
              pLookAheadRepeatFound = False
              LARepeat = ""
            End If
            If CheckNextDyelot Then
              'If the next lot is not blocked then start preparing the tanks
              If (Not Blocked) Or (.Parameters_LookAheadIgnoreBlocked = 1) Then
                If pLookAheadRepeatFound Then
                  'stop trying to check next dyelot if we've found a LA repeat
                  If pLookAheadRepeatFound Then Cancel
                Else
                  .LARequest = False
                  State = LAStart
                  KP1Calloff = Tank1Calloff
                  .DrugroomPreview.LoadNextRecipeSteps
                  .DrugroomDisplayJob1 = Job
                  .LASetSqlString = True
                  .LASqlString = "update dyelots set committed=1 where dyelot='" & Dyelot & "'"
                End If
              End If
            Else
              If pLookAheadRepeatFound Then
                If (Not Blocked) Or (.Parameters_LookAheadIgnoreBlocked = 1) Then
                  .LARequest = False
                  Cancel
                End If
              End If
            End If
          End If
          
        Case LAStart
          If Tank1Enabled Then KP1.Start (Tank1FillLevel - .Parameters_Tank1FillLevelDeadBand), _
                                         (Tank1DesiredTemp - .Parameters_Tank1HeatDeadband), _
                                         Tank1MixTime, _
                                         Tank1MixingOn, _
                                         KP1Calloff, _
                                         300, _
                                         DispenseTank
          State = LARun
    
        Case LARun
          If Tank1Enabled Then KP1.Run ControlCode
          If Not (KP1.IsOn) Then State = LADone
          
        Case LADone
          Cancel
            
      End Select
    Loop Until (State = StatePrevious)
    
    'Cancel this command if parameter is changed
    If .Parameters_LookAheadEnabled <> 1 Then
      Cancel
      Exit Sub
    End If
  
  End With

End Sub

Private Sub ACCommand_Cancel()
'Looks a bit odd but this function is called when a lot finishes
' - LA command must remain active from lot to lot so we only cancel when both tanks are inactive
  If Not (KP1.IsOn) Then Cancel
End Sub

Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property

Public Sub Cancel()
  KP1.Cancel
  State = LAOff
End Sub

Public Property Get IsOn()
  IsOn = (State <> LAOff)
End Property

Public Property Get IsActive()
  IsActive = Not ((State = LAOff) Or (State = LACheckNextDyelot))
End Property

Public Property Get Tank1DispenseError() As Long
 Tank1DispenseError = KP1.AddDispenseError
End Property

'***********************************************************************************************
'******                  Check current Program for the next KP/LA/LS                      ******
'***********************************************************************************************
Public Function CheckExistingProgram(ByVal ControlObject As Object) As Boolean
On Error GoTo Err_CheckExistingProgram

'Added 8/5/2009 - gets current job (dyelot and redye) from the database instead of from parent.job
'Create an ADO connection to the local database
  Dim con As ADODB.Connection15
  Set con = New ADODB.Connection
  con.Open "BatchDyeing"
  
  Dim rsDyelot As ADODB.Recordset15
  Set rsDyelot = con.Execute("SELECT * FROM Dyelots WHERE State=2 order by StartTime")
  rsDyelot.MoveFirst

  If Not rsDyelot.EOF Then
    'Set Dyelot and Redye - these must be non-null because they are the primary key
    pDyelot = CStr(rsDyelot!Dyelot)
    pRedye = CLng(rsDyelot!Redye)
  End If

'Set default return value
  CheckExistingProgram = False
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
    Dim LSFound As Boolean, LAFound As Boolean
    StartChecking = False
    KPFound = False
    LAFound = False
    Dim KP1 As String
    For i = LBound(ProgramSteps) To UBound(ProgramSteps)
      'Ignore step 0
      If i > 0 Then
        Steps = Split(ProgramSteps(i), Chr(255))
        If UBound(Steps) >= 1 Then
          If (Steps(0) = ProgNum) And (Steps(1) = StepNum - 1) Then StartChecking = True
          If StartChecking And (Not KPFound) Then
             Select Case UCase(Left(Trim(Steps(5)), 2))

               Case "KP"
                 'only find first kp after la
                 KPFound = True
                 KP1 = Steps(5)
                 
               Case "LA"
                 'if we find another LA before then next KP, then do nothing 'New 8/24/2009
                 If LAFound Then
                   pLookAheadRepeatFound = True
                   LARepeat = "Program: " & Pad(ProgNum, "0", 4) & " Step: " & Pad(StepNum, "0", 3)
                 End If
                 LAFound = True 'pick up the current LA active
                 
               Case "LS" 'NEW
                 pLookAheadStopFound = True
                 Exit For
                 
              End Select
          End If
        End If
      End If
    Next i
  
'Do Nothing if a second LA is found before next KP
  If pLookAheadRepeatFound Then Exit Function
  
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
      If IsNumeric(parameters(1)) And IsNumeric(parameters(2)) And IsNumeric(parameters(3)) And IsNumeric(parameters(4)) And IsNumeric(parameters(5)) Then
        pTank1FillLevel = CLng(parameters(1)) * 10
        pTank1DesiredTemp = CLng(parameters(2)) * 10
        pTank1MixTime = CLng(parameters(3)) * 60
        PTank1MixingOn = CLng(parameters(4))
        PDispenseTank = CLng(parameters(5))
      End If
    End If
  End If
 
  End With

'everything ok
  CheckExistingProgram = True

'Avoid Error Handler
  Exit Function
  
Err_CheckExistingProgram:
  CheckExistingProgram = False
End Function

'***********************************************************************************************
'******       this function gets the next scheduled program from local database           ******
'***********************************************************************************************
Private Function CheckNextDyelot() As Boolean
On Error GoTo Err_GetNextDyelot
'Always use Connection15 and Recordset15 to avoid ADO version nonsense

'Default return value
  CheckNextDyelot = False

'Create an ADO connection to the local database
  Dim con As ADODB.Connection15
  Set con = New ADODB.Connection
  con.Open "BatchDyeing"
  
  Dim rsDyelot As ADODB.Recordset15
  Set rsDyelot = con.Execute("SELECT * FROM Dyelots WHERE State Is Null ORDER BY StartTime")
  rsDyelot.MoveFirst

  If Not rsDyelot.EOF Then
    'Set Dyelot and Redye - these must be non-null because they are the primary key
    pDyelot = CStr(rsDyelot!Dyelot)
    pRedye = CLng(rsDyelot!Redye)
    'Checked blocked status - SmallInt (Null = unBlocked, any other value = Blocked)
    If IsNull(rsDyelot!Blocked) Then
      pBlocked = False
    Else
      pBlocked = True
    End If
    'Set program - can't schedule dyelot without a program so assume there is one
    '  seems to be a weird bug with ADO (only tried 2.5)
    '  - you have to assign the value before testing for null and if you try to test after assigning the
    '    value the test will fail... very odd
    pProgram = rsDyelot!Program
    'TODO there could be more than one program...
    If IsNumeric(Program) Then
      'Get program steps for this program
      Dim rsProgram As ADODB.Recordset15
      Set rsProgram = con.Execute("SELECT * FROM Programs WHERE ProgramNumber=" & Program)
      rsDyelot.MoveFirst
      'Now look through the program steps
      Dim ProgramSteps As String
      ProgramSteps = rsProgram!Steps
      CheckNextDyelot = CheckNextProgram(ProgramSteps)
    End If
  End If

'Tidy up
  con.Close
  Set con = Nothing
  Set rsDyelot = Nothing
  Set rsProgram = Nothing

'Avoid Error Handler
  Exit Function
  
Err_GetNextDyelot:
  Set con = Nothing
  Set rsDyelot = Nothing
  Set rsProgram = Nothing
  CheckNextDyelot = False

End Function

'***********************************************************************************************
'******                   Check next scheduled program for IP/KP/LA/LS                    ******
'***********************************************************************************************
Private Function CheckNextProgram(ProgramSteps As String) As Boolean
On Error GoTo Err_CheckNextProgram

'Set default return value
  CheckNextProgram = False
  
'Make sure we've got something to check
  If Len(ProgramSteps) <= 0 Then Exit Function
  
'Use this to split out steps in this program
  Dim Steps() As String

'Use this to split out command parameters for a program step
  Dim parameters() As String
  Dim checkingfornotes() As String
'Split the program string into an array of program steps - one step per array element
  Steps = Split(ProgramSteps, vbCrLf)  '0 based array

'Make sure we've got something to check
  If UBound(Steps) <= 0 Then Exit Function
  
'Loop through the program steps to find the first KP command for each tank
  Dim KP1 As String, KP2 As String, IP As Long
  Dim KP1Found As Boolean
  KP1Found = False
  Dim i As Long
  For i = LBound(Steps) To UBound(Steps)
  
    If Not KP1Found Then 'Only look for the first KP1
      Select Case UCase(Left(Steps(i), 2))
      
        Case "IP"
          'get the program number
          parameters = Split(Steps(i), ",")  '0 based array
          IP = CLng(parameters(1))
          If CheckIPProgram(IP) Then
            KP1 = IPKP1
            If KP1 <> "" Then KP1Found = True
          ElseIf pLookAheadRepeatFound Then
            'LA repeat was found: do nothing
            LARepeat = "Program: " & Pad(IP, "0", 4)
            Exit For
          End If
        
        Case "KP"
          parameters = Split(Steps(i), ",")  '0 based array
          If UBound(parameters) >= 1 Then
            If IsNumeric(parameters(1)) Then
              If Len(KP1) <= 0 Then
                KP1Found = True
                KP1 = Steps(i)
              End If
            End If
          End If
            
        Case "LA"
          'if we find another LA before then next KP, then do nothing 'New 8/24/2009
          pLookAheadRepeatFound = True
          LARepeat = "Program: " & Pad(Program, "0", 4) & " Step: " & Pad(i, "0", 3)
          Exit For
                   
        Case "LS"
          pLookAheadStopFound = True
          Exit For
          
      End Select
    End If
    
  Next i

'Do Nothing if a second LA is found before next KP
  If pLookAheadRepeatFound Then Exit Function
  
'Set KP1 Parameters - next prep for tank 1
  pTank1Enabled = False
  pTank1FillLevel = 0
  pTank1DesiredTemp = 0
  pTank1MixTime = 0
  PTank1MixingOn = 0
  PTank1Calloff = 0
  PDispenseTank = 0
  If Len(KP1) > 0 Then
    pTank1Enabled = True
    checkingfornotes = Split(KP1, "'")
    parameters = Split(checkingfornotes(0), ",")  '0 based array
    If UBound(parameters) >= 5 Then
      '1 = Level, 2 = desired temp, 3 = mixing time, 4 = mixer on or off, 5= dispense tank
      If IsNumeric(parameters(1)) And IsNumeric(parameters(2)) And IsNumeric(parameters(3)) And IsNumeric(parameters(4)) And IsNumeric(parameters(5)) Then
        pTank1FillLevel = CLng(parameters(1)) * 10
        pTank1DesiredTemp = CLng(parameters(2)) * 10
        pTank1MixTime = CLng(parameters(3)) * 60
        PTank1MixingOn = CLng(parameters(4))
        PDispenseTank = CLng(parameters(5))
        PTank1Calloff = 1
     End If
    End If
  End If
  
'Everything completed okay
  CheckNextProgram = True
  
'Avoid Error Handler
  Exit Function

Err_CheckNextProgram:
  'Some Code
  CheckNextProgram = False

End Function

'***********************************************************************************************
'******                   1st IP found within next program - Get Steps                    ******
'***********************************************************************************************
Private Function CheckIPProgram(ProgramNumber As Long) As Boolean
On Error GoTo Err_checkipprogram
'Always use Connection15 and Recordset15 to avoid ADO version nonsense
'Default return value
  CheckIPProgram = False

'Create an ADO connection to the local database
  Dim con As ADODB.Connection15
  Set con = New ADODB.Connection
  con.Open "BatchDyeing"
  
  'TODO there could be more than one program...
  If IsNumeric(ProgramNumber) Then
    'Get program steps for this program
    Dim rsProgram As ADODB.Recordset15
    Set rsProgram = con.Execute("SELECT * FROM Programs WHERE ProgramNumber=" & ProgramNumber)
    'Now look through the program steps
    Dim ProgramSteps As String
    ProgramSteps = rsProgram!Steps
    CheckIPProgram = CheckIPProgramSteps(ProgramSteps)
  End If
  
'Tidy up
  con.Close
  Set con = Nothing
  Set rsProgram = Nothing
  
'Avoid Error Handler
  Exit Function
  
Err_checkipprogram:
  Set con = Nothing
  Set rsProgram = Nothing
  CheckIPProgram = False

End Function

'***********************************************************************************************
'******                         1st IP - check for IP/KP/LA/LS                            ******
'***********************************************************************************************
Private Function CheckIPProgramSteps(ProgramSteps As String) As Boolean
On Error GoTo Err_CheckIPProgramSteps

'Set default return value
  CheckIPProgramSteps = False
  
'Make sure we've got something to check
  If Len(ProgramSteps) <= 0 Then Exit Function
  
'Use this to split out steps in this program
  Dim Steps() As String

'Use this to split out command parameters for a program step
  Dim parameters() As String

'Split the program string into an array of program steps - one step per array element
  Steps = Split(ProgramSteps, vbCrLf)  '0 based array

'Make sure we've got something to check
  If UBound(Steps) <= 0 Then Exit Function

'Loop through the program steps to find the first KP command for each tank
  Dim IP2 As Long
  Dim KP1Found As Boolean
  KP1Found = False
  Dim i As Long
  For i = LBound(Steps) To UBound(Steps)
  
    If Not KP1Found Then 'Only look for the first KP
      Select Case UCase(Left(Steps(i), 2))
        
        Case "IP"
          'get the program number
          parameters = Split(Steps(i), ",")  '0 based array
          IP2 = CLng(parameters(1))
          CheckIP2Program (IP2)
           
        Case "KP"
          'Look to see if this is KP1 or KP2
          parameters = Split(Steps(i), ",")  '0 based array
          If UBound(parameters) >= 1 Then
            If IsNumeric(parameters(1)) Then
              If Len(IPKP1) <= 0 Then
                KP1Found = True
                IPKP1 = Steps(i)
              End If
            End If
          End If
            
        Case "LA"
          'if we find another LA before then next KP, then do nothing 'New 8/24/2009
          pLookAheadRepeatFound = True
          Exit For
                   
        Case "LS"
          pLookAheadStopFound = True
          Exit For
         
      End Select
    End If
  Next i

'Do Nothing if a second LA is found before next KP
  If pLookAheadRepeatFound Then Exit Function
  
'Everything completed okay
  CheckIPProgramSteps = True
  
'Avoid Error Handler
  Exit Function

Err_CheckIPProgramSteps:
  'Some Code
  CheckIPProgramSteps = False

End Function

'  Remainder is designed to Check Sub IP's within IP -
'    IP contains IP which contains IP which contains IP
'***********************************************************************************************
'******              IP found within 1st IP program steps - Get it's steps                ******
'***********************************************************************************************
'this is for a IP inside the first IP
Private Function CheckIP2Program(ProgramNumber As Long) As Boolean
On Error GoTo Err_checkipprogram
'Always use Connection15 and Recordset15 to avoid ADO version nonsense
'Default return value
  CheckIP2Program = False

'Create an ADO connection to the local database
  Dim con As ADODB.Connection15
  Set con = New ADODB.Connection
  con.Open "BatchDyeing"
  
  'TODO there could be more than one program...
  If IsNumeric(ProgramNumber) Then
    'Get program steps for this program
    Dim rsProgram As ADODB.Recordset15
    Set rsProgram = con.Execute("SELECT * FROM Programs WHERE ProgramNumber=" & ProgramNumber)
    'Now look through the program steps
    Dim ProgramSteps As String
    ProgramSteps = rsProgram!Steps
    CheckIP2Program = CheckIP2ProgramSteps(ProgramSteps)
  End If
 
'Tidy up
  con.Close
  Set con = Nothing
  Set rsProgram = Nothing
  
'Avoid Error Handler
  Exit Function
  
Err_checkipprogram:
  Set con = Nothing
  Set rsProgram = Nothing
  CheckIP2Program = False

End Function

'***********************************************************************************************
'******                      Check IP within IP's Steps - KP/LA/LS                        ******
'***********************************************************************************************
Private Function CheckIP2ProgramSteps(ProgramSteps As String) As Boolean
On Error GoTo Err_CheckIP2ProgramSteps

'Set default return value
  CheckIP2ProgramSteps = False
  
'Make sure we've got something to check
  If Len(ProgramSteps) <= 0 Then Exit Function
  
'Use this to split out steps in this program
  Dim Steps() As String

'Use this to split out command parameters for a program step
  Dim parameters() As String

'Split the program string into an array of program steps - one step per array element
  Steps = Split(ProgramSteps, vbCrLf)  '0 based array

'Make sure we've got something to check
  If UBound(Steps) <= 0 Then Exit Function

'Loop through the program steps to find the first KP command for each tank
  Dim IP3 As Long
  Dim KP1Found As Boolean
  KP1Found = False
  Dim i As Long
  For i = LBound(Steps) To UBound(Steps)
  
    If Not KP1Found Then 'Only look for the first KP1
      Select Case UCase(Left(Steps(i), 2))
        
        Case "IP"
          'get the program number
          parameters = Split(Steps(i), ",")  '0 based array
          IP3 = CLng(parameters(1))
          CheckIP3Program (IP3)
            
        Case "KP"
          'Look to see if this is KP1 or KP2
          parameters = Split(Steps(i), ",")  '0 based array
          If UBound(parameters) >= 1 Then
            If IsNumeric(parameters(1)) Then
              If Len(IPKP1) <= 0 Then
                KP1Found = True
                IPKP1 = Steps(i)
              End If
            End If
          End If
          
        Case "LA"
          'if we find another LA before then next KP, then do nothing 'New 8/24/2009
          pLookAheadRepeatFound = True
          LARepeat = "Program: " & Pad(Program, "0", 4) & " Step: " & Pad(i, "0", 3)
          Exit For
          
        Case "LS"
          pLookAheadStopFound = True
          Exit For
        
      End Select
    End If
  Next i

'Do Nothing if a second LA is found before next KP
  If pLookAheadRepeatFound Then Exit Function
  
'Everything completed okay
  CheckIP2ProgramSteps = True

'Avoid Error Handler
  Exit Function

Err_CheckIP2ProgramSteps:
  'Some Code
  CheckIP2ProgramSteps = False

End Function

'***********************************************************************************************
'******                           IP found within Sub IP - Get Steps                      ******
'***********************************************************************************************
Private Function CheckIP3Program(ProgramNumber As Long) As Boolean
On Error GoTo Err_checkip3program
'Always use Connection15 and Recordset15 to avoid ADO version nonsense
'Default return value
  CheckIP3Program = False

'Create an ADO connection to the local database
  Dim con As ADODB.Connection15
  Set con = New ADODB.Connection
  con.Open "BatchDyeing"
  
  'TODO there could be more than one program...
  If IsNumeric(ProgramNumber) Then
    'Get program steps for this program
    Dim rsProgram As ADODB.Recordset15
    Set rsProgram = con.Execute("SELECT * FROM Programs WHERE ProgramNumber=" & ProgramNumber)
    'Now look through the program steps
    Dim ProgramSteps As String
    ProgramSteps = rsProgram!Steps
    CheckIP3Program = CheckIP3ProgramSteps(ProgramSteps)
  End If

'Tidy up
  con.Close
  Set con = Nothing
  Set rsProgram = Nothing
  
'Avoid Error Handler
  Exit Function
  
Err_checkip3program:
  Set con = Nothing
  Set rsProgram = Nothing
  CheckIP3Program = False

End Function

'***********************************************************************************************
'******                          Check Sub IP's steps for IP/LA/LS                        ******
'***********************************************************************************************
Private Function CheckIP3ProgramSteps(ProgramSteps As String) As Boolean
On Error GoTo Err_CheckIP3ProgramSteps

'Set default return value
  CheckIP3ProgramSteps = False
  
'Make sure we've got something to check
  If Len(ProgramSteps) <= 0 Then Exit Function
  
'Use this to split out steps in this program
  Dim Steps() As String

'Use this to split out command parameters for a program step
  Dim parameters() As String

'Split the program string into an array of program steps - one step per array element
  Steps = Split(ProgramSteps, vbCrLf)  '0 based array

'Make sure we've got something to check
  If UBound(Steps) <= 0 Then Exit Function
 
'Loop through the program steps to find the first KP command for each tank
  Dim IP4 As Long
  Dim KP1Found As Boolean
  KP1Found = False
  Dim i As Long
  For i = LBound(Steps) To UBound(Steps)
    If Not KP1Found Then 'Only look for the first KP1
    
      Select Case UCase(Left(Steps(i), 2))
        
        Case "IP"
          'get the program number
          parameters = Split(Steps(i), ",")  '0 based array
          IP4 = CLng(parameters(1))
          CheckIP4Program (IP4)
            
        Case "KP"
          'Look to see if this is KP1 or KP2
          parameters = Split(Steps(i), ",")  '0 based array
          If UBound(parameters) >= 1 Then
            If IsNumeric(parameters(1)) Then
              If Len(IPKP1) <= 0 Then
                KP1Found = True
                IPKP1 = Steps(i)
              End If
            End If
          End If
          
        Case "LA"
          'if we find another LA before then next KP, then do nothing 'New 8/24/2009
          pLookAheadRepeatFound = True
          LARepeat = "Program: " & Pad(Program, "0", 4) & " Step: " & Pad(i, "0", 3)
          Exit For
          
        Case "LS"
          pLookAheadStopFound = True
          Exit For
       
      End Select
    End If
  Next i
  
'Do Nothing if a second LA is found before next KP
  If pLookAheadRepeatFound Then Exit Function
   
'Everything completed okay
  CheckIP3ProgramSteps = True

'Avoid Error Handler
  Exit Function

Err_CheckIP3ProgramSteps:
  CheckIP3ProgramSteps = False

End Function

'***********************************************************************************************
'******                    Sub IP found within this IP - Get Program Steps                ******
'***********************************************************************************************
'this is for a IP inside the third IP
Private Function CheckIP4Program(ProgramNumber As Long) As Boolean
On Error GoTo Err_checkip4program
'Always use Connection15 and Recordset15 to avoid ADO version nonsense
'Default return value
  CheckIP4Program = False

'Create an ADO connection to the local database
  Dim con As ADODB.Connection15
  Set con = New ADODB.Connection
  con.Open "BatchDyeing"
  
  'TODO there could be more than one program...
  If IsNumeric(ProgramNumber) Then
    'Get program steps for this program
    Dim rsProgram As ADODB.Recordset15
    Set rsProgram = con.Execute("SELECT * FROM Programs WHERE ProgramNumber=" & ProgramNumber)
    'Now look through the program steps
    Dim ProgramSteps As String
    ProgramSteps = rsProgram!Steps
    CheckIP4Program = CheckIP4ProgramSteps(ProgramSteps)
  End If
  
'Tidy up
  con.Close
  Set con = Nothing
  Set rsProgram = Nothing
  
'Avoid Error Handler
  Exit Function
  
Err_checkip4program:
  Set con = Nothing
  Set rsProgram = Nothing
  CheckIP4Program = False

End Function

'***********************************************************************************************
'******                   Check this Sub IP's steps - KP/LA/LS (No More IP's)             ******
'***********************************************************************************************
Private Function CheckIP4ProgramSteps(ProgramSteps As String) As Boolean
On Error GoTo Err_CheckIP4ProgramSteps

'Set default return value
  CheckIP4ProgramSteps = False
  
'Make sure we've got something to check
  If Len(ProgramSteps) <= 0 Then Exit Function
  
'Use this to split out steps in this program
  Dim Steps() As String

'Use this to split out command parameters for a program step
  Dim parameters() As String

'Split the program string into an array of program steps - one step per array element
  Steps = Split(ProgramSteps, vbCrLf)  '0 based array

'Make sure we've got something to check
  If UBound(Steps) <= 0 Then Exit Function
  
'Loop through the program steps to find the first KP command for each tank
  Dim KP1Found As Boolean
  KP1Found = False
  Dim i As Long
  For i = LBound(Steps) To UBound(Steps)
    
    If Not KP1Found Then 'Only look for the first KP1
      Select Case UCase(Left(Steps(i), 2))
                  
        'not looking for anymore ip
            
        Case "KP"
          'Look to see if this is KP1 or KP2
          parameters = Split(Steps(i), ",")  '0 based array
          If UBound(parameters) >= 1 Then
            If IsNumeric(parameters(1)) Then
              If Len(IPKP1) <= 0 Then
                KP1Found = True
                IPKP1 = Steps(i)
              End If
            End If
          End If
                   
        Case "LA"
          'if we find another LA before then next KP, then do nothing 'New 8/24/2009
          pLookAheadRepeatFound = True
          LARepeat = "Program: " & Pad(Program, "0", 4) & " Step: " & Pad(i, "0", 3)
          Exit For
          
        Case "LS"
          pLookAheadStopFound = True
          Exit For
       
      End Select
    End If
  Next i
    
'Do Nothing if a second LA is found before next KP
  If pLookAheadRepeatFound Then Exit Function
      
'Everything completed okay
  CheckIP4ProgramSteps = True
  
'Avoid Error Handler
  Exit Function

Err_CheckIP4ProgramSteps:
  CheckIP4ProgramSteps = False

End Function



#End If

#If 0 Then
  '===============================================================================================
'LA - Look ahead command
'Version 2013-07-22
'===============================================================================================
Option Explicit
Implements ACCommand

'Main LA State Machine
Public Enum LAState
  LAOff
  LACheckExistingDyelot
  LACheckNextDyelot
  LAStart
  LARun
  LADone
End Enum
Public State As LAState
Public StatePrevious As LAState

'These command states will be transfered to the KP commands in the next program
Public KP1 As New acAddPrepare
Public KP1Calloff As Long

'Use this to check for next dyelot every 60 seconds
Public CheckNextDyelotTimer As New acTimer

Public IPKP1 As String
Public LARepeat As String

'Private storage for properties
Private pDyelot As String
Private pRedye As Long
Private pJob As String
Private pBlocked As Boolean
Private pProgram As String
Private pTank1Enabled As Boolean
Private pTank1FillLevel As Long
Private pTank1DesiredTemp As Long
Private pTank1MixTime As Long
Private PTank1MixingOn As Long
Private PTank1Calloff As Long
Private PDispenseTank As Long
Private pProgramHasAnotherTank1 As Boolean
Private pLookAheadStopFound As Boolean
Private pLookAheadRepeatFound As Boolean
Private pInsertedProgram As Long

Public Property Get Dyelot() As String
  Dyelot = pDyelot
End Property
Public Property Get Redye() As Long
  Redye = pRedye
End Property
Public Property Get Job() As String
  Job = Dyelot
  If Redye > 0 Then Job = Dyelot & "@" & Redye
End Property
Public Property Get Blocked() As Boolean
  Blocked = pBlocked
End Property
Public Property Get Program() As String
  Program = pProgram
End Property
Public Property Get Tank1Enabled() As Boolean
  Tank1Enabled = pTank1Enabled
End Property
Public Property Get Tank1FillLevel() As Long
  Tank1FillLevel = pTank1FillLevel
End Property
Public Property Get Tank1DesiredTemp() As Long
  Tank1DesiredTemp = pTank1DesiredTemp
End Property
Public Property Get Tank1MixTime() As Long
  Tank1MixTime = pTank1MixTime
End Property
Public Property Get Tank1MixingOn() As Long
  Tank1MixingOn = PTank1MixingOn
End Property
Public Property Get Tank1Calloff() As Long
  Tank1Calloff = PTank1Calloff
End Property
Public Property Get DispenseTank() As Long
  DispenseTank = PDispenseTank
End Property
Public Property Get ProgramHasAnotherTank1() As Boolean
  ProgramHasAnotherTank1 = pProgramHasAnotherTank1
End Property
Public Property Get LookAheadStopFound() As Boolean
  LookAheadStopFound = pLookAheadStopFound
End Property
Public Property Get LookAheadRepeatFound() As Boolean
  LookAheadRepeatFound = pLookAheadRepeatFound
End Property
Public Property Get InsertedProg() As Long
  InsertedProg = pInsertedProgram
End Property

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)

'Reset state & variables
  ResetAll

'Early bind ControlCode for speed and convenience
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject

'Initialize
  State = LAOff
  StatePrevious = LAOff
  pTank1Enabled = False

  With ControlCode
    'Make sure the command is enabled
    If .Parameters_LookAheadEnabled = 1 Then
       State = LACheckExistingDyelot
       KP1Calloff = GetKPCalloff(ControlObject)
    End If
  End With

'Carry on my wayward son...
  StepOn = True

End Sub

Private Sub ResetAll()
'Reset command state & variables
  IPKP1 = ""
  LARepeat = ""
  pJob = ""
  pDyelot = ""
  pRedye = 0
  pBlocked = False
  pProgram = ""
  pLookAheadStopFound = False
  pLookAheadRepeatFound = False
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)

'If the command is not active then don't run this code
  If Not IsOn Then Exit Sub

'Early bind ControlCode for speed and convenience
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject

  With ControlCode

    Do
      'Stay in this loop until state changes have completed
      ' - stops odd effects from momentary state transitions
      StatePrevious = State
      Select Case State
        Case LAOff

        Case LACheckExistingDyelot
         'Check to see if there is another kp in this program.
          If CheckExistingProgram(ControlCode) Then
            If ProgramHasAnotherTank1 = True Then
              .LARequest = False
              State = LAStart
            Else
              .LARequest = True
              State = LACheckNextDyelot
            End If
          End If
          'stop trying to check existing program if we've found a LA repeat
          If pLookAheadRepeatFound Then Cancel

        'Check to see if a dyelot is schedule behind this one
        Case LACheckNextDyelot
          'Check every 60 seconds
          If CheckNextDyelotTimer.Finished Then
            CheckNextDyelotTimer = 60
            'Was blocked and ignoreblocked disabled - program may change due to blocked, so reset and recheck
            If Blocked Then
              pLookAheadRepeatFound = False
              LARepeat = ""
            End If
            If CheckNextDyelot Then
              'If the next lot is not blocked then start preparing the tanks
              If (Not Blocked) Or (.Parameters_LookAheadIgnoreBlocked = 1) Then
                If pLookAheadRepeatFound Then
                  'stop trying to check next dyelot if we've found a LA repeat
                  If pLookAheadRepeatFound Then Cancel
                Else
                  .LARequest = False
                  State = LAStart
                  KP1Calloff = Tank1Calloff
                  .DrugroomPreview.LoadNextRecipeSteps
                  .DrugroomDisplayJob1 = Job
                  .LASetSqlString = True
                  .LASqlString = "update dyelots set committed=1 where dyelot='" & Dyelot & "'"
                End If
              End If
            Else
              If pLookAheadRepeatFound Then
                If (Not Blocked) Or (.Parameters_LookAheadIgnoreBlocked = 1) Then
                  .LARequest = False
                  Cancel
                End If
              End If
            End If
          End If

        Case LAStart
          If Tank1Enabled Then KP1.Start (Tank1FillLevel - .Parameters_Tank1FillLevelDeadBand), _
                                         (Tank1DesiredTemp - .Parameters_Tank1HeatDeadband), _
                                         Tank1MixTime, _
                                         Tank1MixingOn, _
                                         KP1Calloff, _
                                         300, _
                                         DispenseTank
          State = LARun

        Case LARun
          If Tank1Enabled Then KP1.Run ControlCode
          If Not (KP1.IsOn) Then State = LADone

        Case LADone
          Cancel

      End Select
    Loop Until (State = StatePrevious)

    'Cancel this command if parameter is changed
    If .Parameters_LookAheadEnabled <> 1 Then
      Cancel
      Exit Sub
    End If

  End With

End Sub

Private Sub ACCommand_Cancel()
'Looks a bit odd but this function is called when a lot finishes
' - LA command must remain active from lot to lot so we only cancel when both tanks are inactive
  If Not (KP1.IsOn) Then Cancel
End Sub

Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property

Public Sub Cancel()
  KP1.Cancel
  State = LAOff
End Sub

Public Property Get IsOn()
  IsOn = (State <> LAOff)
End Property

Public Property Get IsActive()
  IsActive = Not ((State = LAOff) Or (State = LACheckNextDyelot))
End Property

Public Property Get Tank1DispenseError() As Long
 Tank1DispenseError = KP1.AddDispenseError
End Property

'***********************************************************************************************
'******                  Check current Program for the next KP/LA/LS                      ******
'***********************************************************************************************
Public Function CheckExistingProgram(ByVal ControlObject As Object) As Boolean
On Error GoTo Err_CheckExistingProgram

'Added 8/5/2009 - gets current job (dyelot and redye) from the database instead of from parent.job
'Create an ADO connection to the local database
  Dim con As ADODB.Connection15
  Set con = New ADODB.Connection
  con.Open "BatchDyeing"

  Dim rsDyelot As ADODB.Recordset15
  Set rsDyelot = con.Execute("SELECT * FROM Dyelots WHERE State=2 order by StartTime")
  rsDyelot.MoveFirst

  If Not rsDyelot.EOF Then
    'Set Dyelot and Redye - these must be non-null because they are the primary key
    pDyelot = CStr(rsDyelot!Dyelot)
    pRedye = CLng(rsDyelot!Redye)
  End If

'Set default return value
  CheckExistingProgram = False
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
    Dim LSFound As Boolean, LAFound As Boolean
    StartChecking = False
    KPFound = False
    LAFound = False
    Dim KP1 As String
    For i = LBound(ProgramSteps) To UBound(ProgramSteps)
      'Ignore step 0
      If i > 0 Then
        Steps = Split(ProgramSteps(i), Chr(255))
        If UBound(Steps) >= 1 Then
          If (Steps(0) = ProgNum) And (Steps(1) = StepNum - 1) Then StartChecking = True
          If StartChecking And (Not KPFound) Then
             Select Case UCase(Left(Trim(Steps(5)), 2))

               Case "KP"
                 'only find first kp after la
                 KPFound = True
                 KP1 = Steps(5)

               Case "LA"
                 'if we find another LA before then next KP, then do nothing 'New 8/24/2009
                 If LAFound Then
                   pLookAheadRepeatFound = True
                   LARepeat = "Program: " & Pad(ProgNum, "0", 4) & " Step: " & Pad(StepNum, "0", 3)
                 End If
                 LAFound = True 'pick up the current LA active

               Case "LS" 'NEW
                 pLookAheadStopFound = True
                 Exit For

              End Select
          End If
        End If
      End If
    Next i

'Do Nothing if a second LA is found before next KP
  If pLookAheadRepeatFound Then Exit Function

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
      If IsNumeric(parameters(1)) And IsNumeric(parameters(2)) And IsNumeric(parameters(3)) And IsNumeric(parameters(4)) And IsNumeric(parameters(5)) Then
        pTank1FillLevel = CLng(parameters(1)) * 10
        pTank1DesiredTemp = CLng(parameters(2)) * 10
        pTank1MixTime = CLng(parameters(3)) * 60
        PTank1MixingOn = CLng(parameters(4))
        PDispenseTank = CLng(parameters(5))
      End If
    End If
  End If

  End With

'everything ok
  CheckExistingProgram = True

'Avoid Error Handler
  Exit Function

Err_CheckExistingProgram:
  CheckExistingProgram = False
End Function

'***********************************************************************************************
'******       this function gets the next scheduled program from local database           ******
'***********************************************************************************************
Private Function CheckNextDyelot() As Boolean
On Error GoTo Err_GetNextDyelot
'Always use Connection15 and Recordset15 to avoid ADO version nonsense

'Default return value
  CheckNextDyelot = False

'Create an ADO connection to the local database
  Dim con As ADODB.Connection15
  Set con = New ADODB.Connection
  con.Open "BatchDyeing"

  Dim rsDyelot As ADODB.Recordset15
  Set rsDyelot = con.Execute("SELECT * FROM Dyelots WHERE State Is Null ORDER BY StartTime")
  rsDyelot.MoveFirst

  If Not rsDyelot.EOF Then
    'Set Dyelot and Redye - these must be non-null because they are the primary key
    pDyelot = CStr(rsDyelot!Dyelot)
    pRedye = CLng(rsDyelot!Redye)
    'Checked blocked status - SmallInt (Null = unBlocked, any other value = Blocked)
    If IsNull(rsDyelot!Blocked) Then
      pBlocked = False
    Else
      pBlocked = True
    End If
    'Set program - can't schedule dyelot without a program so assume there is one
    '  seems to be a weird bug with ADO (only tried 2.5)
    '  - you have to assign the value before testing for null and if you try to test after assigning the
    '    value the test will fail... very odd
    pProgram = rsDyelot!Program
    'TODO there could be more than one program...
    If IsNumeric(Program) Then
      'Get program steps for this program
      Dim rsProgram As ADODB.Recordset15
      Set rsProgram = con.Execute("SELECT * FROM Programs WHERE ProgramNumber=" & Program)
      rsDyelot.MoveFirst
      'Now look through the program steps
      Dim ProgramSteps As String
      ProgramSteps = rsProgram!Steps
      CheckNextDyelot = CheckNextProgram(ProgramSteps)
    End If
  End If

'Tidy up
  con.Close
  Set con = Nothing
  Set rsDyelot = Nothing
  Set rsProgram = Nothing

'Avoid Error Handler
  Exit Function

Err_GetNextDyelot:
  Set con = Nothing
  Set rsDyelot = Nothing
  Set rsProgram = Nothing
  CheckNextDyelot = False

End Function

'***********************************************************************************************
'******                   Check next scheduled program for IP/KP/LA/LS                    ******
'***********************************************************************************************
Private Function CheckNextProgram(ProgramSteps As String) As Boolean
On Error GoTo Err_CheckNextProgram

'Set default return value
  CheckNextProgram = False

'Make sure we've got something to check
  If Len(ProgramSteps) <= 0 Then Exit Function

'Use this to split out steps in this program
  Dim Steps() As String

'Use this to split out command parameters for a program step
  Dim parameters() As String
  Dim checkingfornotes() As String
'Split the program string into an array of program steps - one step per array element
  Steps = Split(ProgramSteps, vbCrLf)  '0 based array

'Make sure we've got something to check
  If UBound(Steps) <= 0 Then Exit Function

'Loop through the program steps to find the first KP command for each tank
  Dim KP1 As String, KP2 As String, IP As Long
  Dim KP1Found As Boolean
  KP1Found = False
  Dim i As Long
  For i = LBound(Steps) To UBound(Steps)

    If Not KP1Found Then 'Only look for the first KP1
      Select Case UCase(Left(Steps(i), 2))

        Case "IP"
          'get the program number
          parameters = Split(Steps(i), ",")  '0 based array
          IP = CLng(parameters(1))
          If CheckIPProgram(IP) Then
            KP1 = IPKP1
            If KP1 <> "" Then KP1Found = True
          ElseIf pLookAheadRepeatFound Then
            'LA repeat was found: do nothing
            LARepeat = "Program: " & Pad(IP, "0", 4)
            Exit For
          End If

        Case "KP"
          parameters = Split(Steps(i), ",")  '0 based array
          If UBound(parameters) >= 1 Then
            If IsNumeric(parameters(1)) Then
              If Len(KP1) <= 0 Then
                KP1Found = True
                KP1 = Steps(i)
              End If
            End If
          End If

        Case "LA"
          'if we find another LA before then next KP, then do nothing 'New 8/24/2009
          pLookAheadRepeatFound = True
          LARepeat = "Program: " & Pad(Program, "0", 4) & " Step: " & Pad(i, "0", 3)
          Exit For

        Case "LS"
          pLookAheadStopFound = True
          Exit For

      End Select
    End If

  Next i

'Do Nothing if a second LA is found before next KP
  If pLookAheadRepeatFound Then Exit Function

'Set KP1 Parameters - next prep for tank 1
  pTank1Enabled = False
  pTank1FillLevel = 0
  pTank1DesiredTemp = 0
  pTank1MixTime = 0
  PTank1MixingOn = 0
  PTank1Calloff = 0
  PDispenseTank = 0
  If Len(KP1) > 0 Then
    pTank1Enabled = True
    checkingfornotes = Split(KP1, "'")
    parameters = Split(checkingfornotes(0), ",")  '0 based array
    If UBound(parameters) >= 5 Then
      '1 = Level, 2 = desired temp, 3 = mixing time, 4 = mixer on or off, 5= dispense tank
      If IsNumeric(parameters(1)) And IsNumeric(parameters(2)) And IsNumeric(parameters(3)) And IsNumeric(parameters(4)) And IsNumeric(parameters(5)) Then
        pTank1FillLevel = CLng(parameters(1)) * 10
        pTank1DesiredTemp = CLng(parameters(2)) * 10
        pTank1MixTime = CLng(parameters(3)) * 60
        PTank1MixingOn = CLng(parameters(4))
        PDispenseTank = CLng(parameters(5))
        PTank1Calloff = 1
     End If
    End If
  End If

'Everything completed okay
  CheckNextProgram = True

'Avoid Error Handler
  Exit Function

Err_CheckNextProgram:
  'Some Code
  CheckNextProgram = False

End Function

'***********************************************************************************************
'******                   1st IP found within next program - Get Steps                    ******
'***********************************************************************************************
Private Function CheckIPProgram(ProgramNumber As Long) As Boolean
On Error GoTo Err_checkipprogram
'Always use Connection15 and Recordset15 to avoid ADO version nonsense
'Default return value
  CheckIPProgram = False

'Create an ADO connection to the local database
  Dim con As ADODB.Connection15
  Set con = New ADODB.Connection
  con.Open "BatchDyeing"

  'TODO there could be more than one program...
  If IsNumeric(ProgramNumber) Then
    'Get program steps for this program
    Dim rsProgram As ADODB.Recordset15
    Set rsProgram = con.Execute("SELECT * FROM Programs WHERE ProgramNumber=" & ProgramNumber)
    'Now look through the program steps
    Dim ProgramSteps As String
    ProgramSteps = rsProgram!Steps
    CheckIPProgram = CheckIPProgramSteps(ProgramSteps)
  End If

'Tidy up
  con.Close
  Set con = Nothing
  Set rsProgram = Nothing

'Avoid Error Handler
  Exit Function

Err_checkipprogram:
  Set con = Nothing
  Set rsProgram = Nothing
  CheckIPProgram = False

End Function

'***********************************************************************************************
'******                         1st IP - check for IP/KP/LA/LS                            ******
'***********************************************************************************************
Private Function CheckIPProgramSteps(ProgramSteps As String) As Boolean
On Error GoTo Err_CheckIPProgramSteps

'Set default return value
  CheckIPProgramSteps = False

'Make sure we've got something to check
  If Len(ProgramSteps) <= 0 Then Exit Function

'Use this to split out steps in this program
  Dim Steps() As String

'Use this to split out command parameters for a program step
  Dim parameters() As String

'Split the program string into an array of program steps - one step per array element
  Steps = Split(ProgramSteps, vbCrLf)  '0 based array

'Make sure we've got something to check
  If UBound(Steps) <= 0 Then Exit Function

'Loop through the program steps to find the first KP command for each tank
  Dim IP2 As Long
  Dim KP1Found As Boolean
  KP1Found = False
  Dim i As Long
  For i = LBound(Steps) To UBound(Steps)

    If Not KP1Found Then 'Only look for the first KP
      Select Case UCase(Left(Steps(i), 2))

        Case "IP"
          'get the program number
          parameters = Split(Steps(i), ",")  '0 based array
          IP2 = CLng(parameters(1))
          CheckIP2Program (IP2)

        Case "KP"
          'Look to see if this is KP1 or KP2
          parameters = Split(Steps(i), ",")  '0 based array
          If UBound(parameters) >= 1 Then
            If IsNumeric(parameters(1)) Then
              If Len(IPKP1) <= 0 Then
                KP1Found = True
                IPKP1 = Steps(i)
              End If
            End If
          End If

        Case "LA"
          'if we find another LA before then next KP, then do nothing 'New 8/24/2009
          pLookAheadRepeatFound = True
          Exit For

        Case "LS"
          pLookAheadStopFound = True
          Exit For

      End Select
    End If
  Next i

'Do Nothing if a second LA is found before next KP
  If pLookAheadRepeatFound Then Exit Function

'Everything completed okay
  CheckIPProgramSteps = True

'Avoid Error Handler
  Exit Function

Err_CheckIPProgramSteps:
  'Some Code
  CheckIPProgramSteps = False

End Function

'  Remainder is designed to Check Sub IP's within IP -
'    IP contains IP which contains IP which contains IP
'***********************************************************************************************
'******              IP found within 1st IP program steps - Get it's steps                ******
'***********************************************************************************************
'this is for a IP inside the first IP
Private Function CheckIP2Program(ProgramNumber As Long) As Boolean
On Error GoTo Err_checkipprogram
'Always use Connection15 and Recordset15 to avoid ADO version nonsense
'Default return value
  CheckIP2Program = False

'Create an ADO connection to the local database
  Dim con As ADODB.Connection15
  Set con = New ADODB.Connection
  con.Open "BatchDyeing"

  'TODO there could be more than one program...
  If IsNumeric(ProgramNumber) Then
    'Get program steps for this program
    Dim rsProgram As ADODB.Recordset15
    Set rsProgram = con.Execute("SELECT * FROM Programs WHERE ProgramNumber=" & ProgramNumber)
    'Now look through the program steps
    Dim ProgramSteps As String
    ProgramSteps = rsProgram!Steps
    CheckIP2Program = CheckIP2ProgramSteps(ProgramSteps)
  End If

'Tidy up
  con.Close
  Set con = Nothing
  Set rsProgram = Nothing

'Avoid Error Handler
  Exit Function

Err_checkipprogram:
  Set con = Nothing
  Set rsProgram = Nothing
  CheckIP2Program = False

End Function

'***********************************************************************************************
'******                      Check IP within IP's Steps - KP/LA/LS                        ******
'***********************************************************************************************
Private Function CheckIP2ProgramSteps(ProgramSteps As String) As Boolean
On Error GoTo Err_CheckIP2ProgramSteps

'Set default return value
  CheckIP2ProgramSteps = False

'Make sure we've got something to check
  If Len(ProgramSteps) <= 0 Then Exit Function

'Use this to split out steps in this program
  Dim Steps() As String

'Use this to split out command parameters for a program step
  Dim parameters() As String

'Split the program string into an array of program steps - one step per array element
  Steps = Split(ProgramSteps, vbCrLf)  '0 based array

'Make sure we've got something to check
  If UBound(Steps) <= 0 Then Exit Function

'Loop through the program steps to find the first KP command for each tank
  Dim IP3 As Long
  Dim KP1Found As Boolean
  KP1Found = False
  Dim i As Long
  For i = LBound(Steps) To UBound(Steps)

    If Not KP1Found Then 'Only look for the first KP1
      Select Case UCase(Left(Steps(i), 2))

        Case "IP"
          'get the program number
          parameters = Split(Steps(i), ",")  '0 based array
          IP3 = CLng(parameters(1))
          CheckIP3Program (IP3)

        Case "KP"
          'Look to see if this is KP1 or KP2
          parameters = Split(Steps(i), ",")  '0 based array
          If UBound(parameters) >= 1 Then
            If IsNumeric(parameters(1)) Then
              If Len(IPKP1) <= 0 Then
                KP1Found = True
                IPKP1 = Steps(i)
              End If
            End If
          End If

        Case "LA"
          'if we find another LA before then next KP, then do nothing 'New 8/24/2009
          pLookAheadRepeatFound = True
          LARepeat = "Program: " & Pad(Program, "0", 4) & " Step: " & Pad(i, "0", 3)
          Exit For

        Case "LS"
          pLookAheadStopFound = True
          Exit For

      End Select
    End If
  Next i

'Do Nothing if a second LA is found before next KP
  If pLookAheadRepeatFound Then Exit Function

'Everything completed okay
  CheckIP2ProgramSteps = True

'Avoid Error Handler
  Exit Function

Err_CheckIP2ProgramSteps:
  'Some Code
  CheckIP2ProgramSteps = False

End Function

'***********************************************************************************************
'******                           IP found within Sub IP - Get Steps                      ******
'***********************************************************************************************
Private Function CheckIP3Program(ProgramNumber As Long) As Boolean
On Error GoTo Err_checkip3program
'Always use Connection15 and Recordset15 to avoid ADO version nonsense
'Default return value
  CheckIP3Program = False

'Create an ADO connection to the local database
  Dim con As ADODB.Connection15
  Set con = New ADODB.Connection
  con.Open "BatchDyeing"

  'TODO there could be more than one program...
  If IsNumeric(ProgramNumber) Then
    'Get program steps for this program
    Dim rsProgram As ADODB.Recordset15
    Set rsProgram = con.Execute("SELECT * FROM Programs WHERE ProgramNumber=" & ProgramNumber)
    'Now look through the program steps
    Dim ProgramSteps As String
    ProgramSteps = rsProgram!Steps
    CheckIP3Program = CheckIP3ProgramSteps(ProgramSteps)
  End If

'Tidy up
  con.Close
  Set con = Nothing
  Set rsProgram = Nothing

'Avoid Error Handler
  Exit Function

Err_checkip3program:
  Set con = Nothing
  Set rsProgram = Nothing
  CheckIP3Program = False

End Function

'***********************************************************************************************
'******                          Check Sub IP's steps for IP/LA/LS                        ******
'***********************************************************************************************
Private Function CheckIP3ProgramSteps(ProgramSteps As String) As Boolean
On Error GoTo Err_CheckIP3ProgramSteps

'Set default return value
  CheckIP3ProgramSteps = False

'Make sure we've got something to check
  If Len(ProgramSteps) <= 0 Then Exit Function

'Use this to split out steps in this program
  Dim Steps() As String

'Use this to split out command parameters for a program step
  Dim parameters() As String

'Split the program string into an array of program steps - one step per array element
  Steps = Split(ProgramSteps, vbCrLf)  '0 based array

'Make sure we've got something to check
  If UBound(Steps) <= 0 Then Exit Function

'Loop through the program steps to find the first KP command for each tank
  Dim IP4 As Long
  Dim KP1Found As Boolean
  KP1Found = False
  Dim i As Long
  For i = LBound(Steps) To UBound(Steps)
    If Not KP1Found Then 'Only look for the first KP1

      Select Case UCase(Left(Steps(i), 2))

        Case "IP"
          'get the program number
          parameters = Split(Steps(i), ",")  '0 based array
          IP4 = CLng(parameters(1))
          CheckIP4Program (IP4)

        Case "KP"
          'Look to see if this is KP1 or KP2
          parameters = Split(Steps(i), ",")  '0 based array
          If UBound(parameters) >= 1 Then
            If IsNumeric(parameters(1)) Then
              If Len(IPKP1) <= 0 Then
                KP1Found = True
                IPKP1 = Steps(i)
              End If
            End If
          End If

        Case "LA"
          'if we find another LA before then next KP, then do nothing 'New 8/24/2009
          pLookAheadRepeatFound = True
          LARepeat = "Program: " & Pad(Program, "0", 4) & " Step: " & Pad(i, "0", 3)
          Exit For

        Case "LS"
          pLookAheadStopFound = True
          Exit For

      End Select
    End If
  Next i

'Do Nothing if a second LA is found before next KP
  If pLookAheadRepeatFound Then Exit Function

'Everything completed okay
  CheckIP3ProgramSteps = True

'Avoid Error Handler
  Exit Function

Err_CheckIP3ProgramSteps:
  CheckIP3ProgramSteps = False

End Function

'***********************************************************************************************
'******                    Sub IP found within this IP - Get Program Steps                ******
'***********************************************************************************************
'this is for a IP inside the third IP
Private Function CheckIP4Program(ProgramNumber As Long) As Boolean
On Error GoTo Err_checkip4program
'Always use Connection15 and Recordset15 to avoid ADO version nonsense
'Default return value
  CheckIP4Program = False

'Create an ADO connection to the local database
  Dim con As ADODB.Connection15
  Set con = New ADODB.Connection
  con.Open "BatchDyeing"

  'TODO there could be more than one program...
  If IsNumeric(ProgramNumber) Then
    'Get program steps for this program
    Dim rsProgram As ADODB.Recordset15
    Set rsProgram = con.Execute("SELECT * FROM Programs WHERE ProgramNumber=" & ProgramNumber)
    'Now look through the program steps
    Dim ProgramSteps As String
    ProgramSteps = rsProgram!Steps
    CheckIP4Program = CheckIP4ProgramSteps(ProgramSteps)
  End If

'Tidy up
  con.Close
  Set con = Nothing
  Set rsProgram = Nothing

'Avoid Error Handler
  Exit Function

Err_checkip4program:
  Set con = Nothing
  Set rsProgram = Nothing
  CheckIP4Program = False

End Function

'***********************************************************************************************
'******                   Check this Sub IP's steps - KP/LA/LS (No More IP's)             ******
'***********************************************************************************************
Private Function CheckIP4ProgramSteps(ProgramSteps As String) As Boolean
On Error GoTo Err_CheckIP4ProgramSteps

'Set default return value
  CheckIP4ProgramSteps = False

'Make sure we've got something to check
  If Len(ProgramSteps) <= 0 Then Exit Function

'Use this to split out steps in this program
  Dim Steps() As String

'Use this to split out command parameters for a program step
  Dim parameters() As String

'Split the program string into an array of program steps - one step per array element
  Steps = Split(ProgramSteps, vbCrLf)  '0 based array

'Make sure we've got something to check
  If UBound(Steps) <= 0 Then Exit Function

'Loop through the program steps to find the first KP command for each tank
  Dim KP1Found As Boolean
  KP1Found = False
  Dim i As Long
  For i = LBound(Steps) To UBound(Steps)

    If Not KP1Found Then 'Only look for the first KP1
      Select Case UCase(Left(Steps(i), 2))

        'not looking for anymore ip

        Case "KP"
          'Look to see if this is KP1 or KP2
          parameters = Split(Steps(i), ",")  '0 based array
          If UBound(parameters) >= 1 Then
            If IsNumeric(parameters(1)) Then
              If Len(IPKP1) <= 0 Then
                KP1Found = True
                IPKP1 = Steps(i)
              End If
            End If
          End If

        Case "LA"
          'if we find another LA before then next KP, then do nothing 'New 8/24/2009
          pLookAheadRepeatFound = True
          LARepeat = "Program: " & Pad(Program, "0", 4) & " Step: " & Pad(i, "0", 3)
          Exit For

        Case "LS"
          pLookAheadStopFound = True
          Exit For

      End Select
    End If
  Next i

'Do Nothing if a second LA is found before next KP
  If pLookAheadRepeatFound Then Exit Function

'Everything completed okay
  CheckIP4ProgramSteps = True

'Avoid Error Handler
  Exit Function

Err_CheckIP4ProgramSteps:
  CheckIP4ProgramSteps = False

End Function

#End If















#If 0 Then
  '===============================================================================================
'General purpose add tank prepare command
'===============================================================================================
Option Explicit
Public Enum AddPrepareState
  Off
  WaitIdle
  PreFill
  DispenseWaitTurn
  DispenseWaitReady
  DispenseWaitProducts
  DispenseWaitResponse
  Fill
  Heat
  Slow
  Fast
  MixForTime
  Ready
  KAInterlock
  KATransfer1
  KARinse
  KATransfer2
  KARinseToDrain
  KATransferToDrain
  InManual
End Enum

Public State As AddPrepareState
Public StatePrev As AddPrepareState
Public StateString As String

Public AddPreFillLevel As Long
Public AddFillLevel As Long
Public DesiredTemperature As Long
Public AddMixTime As Long
Public AddMixing As Boolean
Public OverrunTime As Long
Public OverrunTimer As New acTimer
Public MixTimer As New acTimer
Public HeatOn As Boolean, FillOn As Boolean
Public AddCallOff As Long
Public Timer As New acTimer
Public AdvanceTimer As New acTimer
Public HeatPrepTimer As New acTimer
Public DispenseTimer As New acTimer
Private NumberOfRinses As Long
Public WaitReadyTimer As New acTimerUp

'variables for alarms on manual dispense or error. does nothing
Public ManualAdd As Boolean
Public AddDispenseError As Boolean
Public AlarmRedyeIssue As Boolean

'For display only
Public AddRecipeStep As String
Private RecipeSteps(1 To 64, 1 To 8) As String

'For Display in drugroom preview
Public DrugroomDisplay As String

'Dispense States
Public DispenseCalloff As Long            'This filled in by us
Public DispenseTank As Long               'This filled in by us
Public Dispensestate As Long              'This filled in by AutoDispenser
Public Dispenseproducts As String         'This filled in by Autodispenser

Public DispenseDyesOnly As Boolean        'Used to determine dispenser delay type
Public DispenseChemsOnly As Boolean
Public DispenseDyesChems As Boolean
Private DispenseDyes As Boolean
Private DispenseChems As Boolean

'Dispense States
Private Const DispenseReady As Long = 101
Private Const DispenseBusy As Long = 102
Private Const DispenseAuto As Long = 201
Private Const DispenseScheduled As Long = 202
Private Const DispenseComplete As Long = 301
Private Const DispenseManual As Long = 302
Private Const DispenseError As Long = 309

Public Sub Start(FillLevel As Long, DesiredTemp As Long, MixTime As Long, MixingOn As Long, Calloff As Long, StandardTime As Long, DTank As Long)
  AddFillLevel = FillLevel
  DesiredTemperature = DesiredTemp
  If DesiredTemperature > 1800 Then DesiredTemperature = 1800
  AddMixTime = MixTime
  AddMixing = MixingOn
  AddCallOff = Calloff
  OverrunTime = StandardTime
  DispenseTank = DTank

  State = WaitIdle
  WaitReadyTimer.Pause
  Timer = 5
  AddDispenseError = False

  DispenseDyes = False
  DispenseChems = False
  DispenseDyesOnly = False
  DispenseChemsOnly = False
  DispenseDyesChems = False

End Sub

Public Sub Run(ByVal ControlObject As Object)
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode
    Dispensestate = .Dispensestate
    Dispenseproducts = .Dispenseproducts

    'Run an advance timer to reset alarms, where necessary
    If Not .IO_RemoteRun Then AdvanceTimer.TimeRemaining = 2
    If AdvanceTimer.Finished Then AlarmRedyeIssue = False

  Do
    'Remember state and loop until state does not change
    'This makes sure we go through all the state changes before proceeding and setting IO
     StatePrev = State

  Select Case State

    Case Off
      StateString = ""
      DrugroomDisplay = "Tank 1 Idle"

    Case WaitIdle    'Wait for destination tank to be idle - no active drains...
      DrugroomDisplay = "Wait For Destination Idle "
      If AlarmRedyeIssue And (.Parameters_EnableRedyeIssueAlarm = 1) Then
        Timer = 5
        StateString = "Check Tank (Redye Issue), Hold Run to reset " & TimerString(AdvanceTimer.TimeRemaining)
        DrugroomDisplay = "Check Tank (Redye Issue), Hold Run to Reset " & TimerString(AdvanceTimer.TimeRemaining)
      ElseIf (DispenseTank = 1) And (.AD.IsOn Or .ManualAddDrain.IsActive) Then
        Timer = 5
        StateString = "Wait to Add - " & .AD.StateString
      ElseIf (DispenseTank = 2) And (.RD.IsOn Or .ManualReserveDrain.IsActive) Then
        Timer = 5
        StateString = "Wait to Reserve - " & .RD.StateString
      Else
        StateString = "Wait for Destination Idle "
      End If
      If Timer.Finished Then State = PreFill

    Case PreFill
      StateString = "Tank 1 PreFill to " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters_Tank1PreFillLevel / 10, "0", 3) & "%"
      DrugroomDisplay = "Tank 1 PreFill to " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters_Tank1PreFillLevel / 10, "0", 3) & "%"
      If (.Tank1Level >= .Parameters_Tank1PreFillLevel) Then
         State = DispenseWaitTurn
      End If

    Case DispenseWaitTurn    'Wait for other tank to be finished with dispensing before new dispense
      StateString = "Wait for Turn "
      DrugroomDisplay = "Wait for Turn"
      If (AddCallOff = 0) Or (.Parameters_DispenseEnabled <> 1) Then
        HeatPrepTimer = .Parameters_Tank1HeatPrepTimer
        State = Fill
        ManualAdd = True
      Else
        DispenseTimer = .Parameters_DispenseReadyDelayTime * 60
        State = DispenseWaitReady
        DispenseCalloff = 0
      End If

    Case DispenseWaitReady    'Wait to make sure dispenser has completed previous dispense and is ready
      StateString = "Wait for Dispenser Ready "
      DrugroomDisplay = "Wait for Dispenser Ready "
      If Dispensestate = DispenseReady Then                                   '101
         State = DispenseWaitProducts
         DispenseTimer = .Parameters_DispenseResponseDelayTime * 60
         DispenseCalloff = AddCallOff
         .DispenseTank = DispenseTank
      End If
      If (AddCallOff = 0) Or (.Parameters_DispenseEnabled <> 1) Then
        HeatPrepTimer = .Parameters_Tank1HeatPrepTimer
        State = Fill
        DispenseCalloff = 0
        ManualAdd = True
      End If

    Case DispenseWaitProducts     'Wait for Dispensestate & DispenseProducts String to be set to determine how to signal delays
      StateString = "Wait for Dispense Products "
      DrugroomDisplay = "Wait for Dispense Products "
      DispenseTimer = .Parameters_DispenseResponseDelayTime * 60
      If Dispensestate <> DispenseReady Then                '(DispenseBusy = 102) but if these no recipe, we'll get a (DispenseManual = 302)
        If Dispenseproducts <> "" Then
          'split the products
          'products() = "Step <calloff>| <Ingredient_id> : <Amount> <Units> <DResult> | <Ingredient_Desc> "
          Dim ProductsArray() As String, i As Long
          ProductsArray = Split(Dispenseproducts, "|")
          For i = 1 To UBound(ProductsArray)                'Disregard the 0 row "Step <calloff>
            Dim position As Long
            position = InStr(ProductsArray(i), ":")         'Need to verify ":" is in the current row due to second split "|"
            If position > 0 Then
              Dim test As String
              test = Mid(ProductsArray(i), 6, 1)
              If test = ":" Then                            'If ":" is at 5th position, we have a chemical => "| 1004: ..."
                DispenseChems = True
              Else
                DispenseDyes = True
              End If
            End If
          Next i
          If DispenseChems And DispenseDyes Then
            DispenseDyesChems = True
          Else
            If DispenseChems Then
              DispenseChemsOnly = True
            Else: DispenseDyesOnly = True
            End If
          End If
        End If
        'Proceed to the next state
        State = DispenseWaitResponse
        DispenseCalloff = AddCallOff
        .DispenseTank = DispenseTank
      End If
      If (AddCallOff = 0) Or (.Parameters_DispenseEnabled <> 1) Then
        HeatPrepTimer = .Parameters_Tank1HeatPrepTimer
        State = Fill
        DispenseCalloff = 0
        ManualAdd = True
      End If

    Case DispenseWaitResponse     'Wait for response from dispenser
      StateString = "Wait For Reponse From Dispenser "
      DrugroomDisplay = "Wait for Response From Dispenser "
      AddRecipeStep = Dispenseproducts
      Select Case Dispensestate
         Case DispenseComplete                                                '301
          If Not .WK.IsOn Then
            If DispenseTank = 2 Then
             .ReserveReady = True
            ElseIf DispenseTank = 1 Then
             .AddReady = True
            End If
          End If
          .KR.ACCommand_Cancel
          If .LA.KP1.IsOn Then .LAActive = True
          State = Off
          Cancel
          .DispenseTank = 0
          DispenseTank = 0
          DispenseCalloff = 0

        Case DispenseManual                                                  '302
          HeatPrepTimer = .Parameters_Tank1HeatPrepTimer
          State = Fill
          DispenseCalloff = 0
          ManualAdd = True

        Case DispenseError                                                   '309
          AddDispenseError = True
          HeatPrepTimer = .Parameters_Tank1HeatPrepTimer
          State = Fill
          DispenseCalloff = 0

      End Select

    Case Fill
      StateString = "Tank 1 Fill to " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(AddFillLevel / 10, "0", 3) & "%"
      DrugroomDisplay = "Fill to " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(AddFillLevel / 10, "0", 3) & "%"
      If .IO_Tank1Manual_SW Then State = InManual
      If AddFillLevel = 0 Then
        WaitReadyTimer.Start
        State = Slow
      End If
      If .Tank1Level > AddFillLevel Then State = Heat

    Case Heat
      If Not .DrugroomMixerOn Then
        StateString = "Tank 1 Level Too Low to Heat " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters_Tank1MixerOnLevel / 10, "0", 3) & "%"
        DrugroomDisplay = "Level Too Low to Heat " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters_Tank1MixerOnLevel / 10, "0", 3) & "%"
      Else
        StateString = "Tank 1 Heat to " & Pad(.IO_Tank1Temp / 10, "0", 3) & "F / " & Pad(DesiredTemperature / 10, "0", 3)
        DrugroomDisplay = "Heat to " & Pad(.IO_Tank1Temp / 10, "0", 3) & "F / " & Pad(DesiredTemperature / 10, "0", 3)
      End If
      AddMixing = True
      If .IO_Tank1Manual_SW Then State = InManual
      If AddFillLevel = 0 Then State = Slow
      If .IO_Tank1Temp < DesiredTemperature Then HeatOn = True
      If .IO_Tank1Temp > DesiredTemperature Then
        HeatOn = False
        State = Slow
        WaitReadyTimer.Start
        OverrunTimer = OverrunTime
      End If

    Case Slow
      StateString = "Waiting on Tank 1 " & TimerString(WaitReadyTimer.TimeElapsed)
      DrugroomDisplay = "Prepare Tank 1"
      If .IO_Tank1Manual_SW Then State = InManual
      If .IO_Tank1Temp >= DesiredTemperature Then HeatOn = False
      If .IO_Tank1Temp < DesiredTemperature Then HeatOn = True
      If DispenseTank = 2 Then
        'Reserve Tank
        If .RT.IsWaitReady Then State = Fast
      ElseIf DispenseTank = 1 Then
        'Add Tank
        If .AC.IsWaitReady Then State = Fast
        If .AT.IsWaitReady Then State = Fast
      End If
      If .Tank1Ready Then
        WaitReadyTimer.Pause
        State = MixForTime
        MixTimer = AddMixTime
      End If
      If WaitReadyTimer.IsPaused Then WaitReadyTimer.Restart 'should cover LA copy to KP function
      If OverrunTimer.Finished Then
         State = Fast
      End If

    Case Fast
      StateString = "Waiting on Tank 1 " & TimerString(WaitReadyTimer.TimeElapsed)
      DrugroomDisplay = "Prepare Tank 1"
      If .IO_Tank1Manual_SW Then State = InManual
      If .IO_Tank1Temp >= DesiredTemperature Then HeatOn = False
      If .IO_Tank1Temp < DesiredTemperature Then HeatOn = True
      If WaitReadyTimer.IsPaused Then WaitReadyTimer.Restart 'should cover LA copy to KP function
      If (.Tank1Level And .Tank1Ready) Then
        WaitReadyTimer.Pause
        State = MixForTime
        MixTimer = AddMixTime
      End If

    Case MixForTime
      If Not .DrugroomMixerOn Then
        StateString = "Tank 1 Level Too Low to Mix " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters_Tank1MixerOnLevel / 10, "0", 3) & "%"
        DrugroomDisplay = "Level Too Low to Mix " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters_Tank1MixerOnLevel / 10, "0", 3) & "%"
      Else
        StateString = "Tank 1 Mix For Time " & TimerString(MixTimer.TimeRemaining)
        DrugroomDisplay = "Mix For Time " & TimerString(MixTimer.TimeRemaining)
      End If
      If .IO_Tank1Manual_SW Then State = InManual
      If AddFillLevel = 0 Then State = Ready
      If .IO_Tank1Temp >= DesiredTemperature Then HeatOn = False
      If .IO_Tank1Temp < DesiredTemperature Then HeatOn = True
      If MixTimer.Finished Then State = Ready
      If Not .Tank1Ready Then
        State = Fast
        WaitReadyTimer.Restart
      End If

    Case Ready
      StateString = "Tank 1 Ready "
      DrugroomDisplay = "Ready "
      If .IO_Tank1Manual_SW Then State = InManual
      If (DispenseTank = 1) And (.AdditionLevel > .Parameters_AdditionMaxTransferLevel) Then
         State = KAInterlock
      Else
        State = KATransfer1
      End If
      FillOn = False
      HeatOn = False
      AddMixing = False
      AddDispenseError = False
      ManualAdd = False

    Case KAInterlock
      StateString = "Add Level High for Tank 1 Transfer " & Pad(.AdditionLevel / 10, "0", 3) & "%"
      DrugroomDisplay = "Add Level High for Tank 1 Transfer " & Pad(.AdditionLevel / 10, "0", 3) & "%"
      If .IO_Tank1Manual_SW Then State = InManual
      If (.AdditionLevel <= .Parameters_AdditionMaxTransferLevel) Then
        State = KATransfer1
      End If

    Case KATransfer1
      If (.Tank1Level > 10) Then
        Timer = .Parameters_Tank1TimeBeforeRinse
        StateString = "Tank 1 Transfer Empty " & Pad(.Tank1Level / 10, "0", 3) & "%"
        DrugroomDisplay = "Transfer Empty " & Pad(.Tank1Level / 10, "0", 3) & "%"
      Else
        StateString = "Tank 1 Transfer Empty " & TimerString(Timer.TimeRemaining)
        DrugroomDisplay = "Transfer Empty " & TimerString(Timer.TimeRemaining)
      End If
      If .IO_Tank1Manual_SW Then State = InManual
      If Timer.Finished Then
        NumberOfRinses = .Parameters_DrugroomRinses
        If .KR.IsOn Then
          If (.KR.RinseMachine = 0) And (.KR.RinseDrain = 0) Then
            If Not .WK.IsOn Then
              If (DispenseTank = 1) Then .AddReady = True
              If (DispenseTank = 2) Then .ReserveReady = True
            End If
            If .LA.KP1.IsOn Then .LAActive = True
            DispenseTank = 0
            .DispenseTank = 0
            .Tank1Ready = False
            .KR.ACCommand_Cancel
            Cancel
          ElseIf .KR.RinseMachine = 0 Then
            State = KARinseToDrain
            Timer = .Parameters_Tank1RinseToDrainTime
          End If
        Else
          State = KARinse
          Timer = .Parameters_Tank1RinseTime
        End If
      End If

    Case KARinse
      StateString = "Tank 1 Rinse " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters_Tank1RinseLevel / 10, "0", 3) & "%"
      DrugroomDisplay = "Rinse " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters_Tank1RinseLevel / 10, "0", 3) & "%"
      If .IO_Tank1Manual_SW Then State = InManual
      If (.Tank1Level > .Parameters_Tank1RinseLevel) Then
        State = KATransfer2
        Timer = .Parameters_Tank1TimeAfterRinse
      End If

    Case KATransfer2
      If (.Tank1Level > 10) Then
        Timer = .Parameters_Tank1TimeAfterRinse
        StateString = "Tank 1 Transfer Rinse " & Pad(.Tank1Level / 10, "0", 3) & "%"
        DrugroomDisplay = "Transfer Rinse " & Pad(.Tank1Level / 10, "0", 3) & "%"
      Else
        StateString = "Tank 1 Transfer Rinse " & TimerString(Timer.TimeRemaining)
        DrugroomDisplay = "Transfer Rinse " & TimerString(Timer.TimeRemaining)
      End If
      If .IO_Tank1Manual_SW Then State = InManual
      If Timer.Finished Then
        If Not .WK.IsOn Then
          If (DispenseTank = 1) Then .AddReady = True
          If (DispenseTank = 2) Then .ReserveReady = True
        End If
        If .LA.KP1.IsOn Then .LAActive = True
        DispenseTank = 0
        .DispenseTank = 0
        .Tank1Ready = False
        State = KARinseToDrain
        Timer = .Parameters_Tank1RinseToDrainTime
      End If

    Case KARinseToDrain
      StateString = "Tank 1 Rinse To Drain " & TimerString(Timer.TimeRemaining)
      DrugroomDisplay = "Rinse To Drain " & TimerString(Timer.TimeRemaining)
      If .IO_Tank1Manual_SW Then State = InManual
      If Timer.Finished Or (.Tank1Level >= 500) Then
        State = KATransferToDrain
        Timer = .Parameters_Tank1DrainTime
        NumberOfRinses = NumberOfRinses - 1
      End If

    Case KATransferToDrain
      If (.Tank1Level > 10) Then
        Timer = .Parameters_Tank1DrainTime
        StateString = "Tank 1 To Drain " & Pad(.Tank1Level / 10, "0", 3) & "%"
        DrugroomDisplay = "Transfer To Drain " & Pad(.Tank1Level / 10, "0", 3) & "%"
      Else
        StateString = "Tank 1 To Drain " & TimerString(Timer.TimeRemaining)
        DrugroomDisplay = "Transfer To Drain " & TimerString(Timer.TimeRemaining)
      End If
      If .IO_Tank1Manual_SW Then State = InManual
      If Timer.Finished Then
        If NumberOfRinses > 0 Then
          State = KARinseToDrain
          Timer = .Parameters_Tank1RinseToDrainTime
        Else
          If Not .WK.IsOn Then
            If (DispenseTank = 1) Then .AddReady = True
            If (DispenseTank = 2) Then .ReserveReady = True
          End If
          If .LA.KP1.IsOn Then .LAActive = True
          DispenseTank = 0
          .DispenseTank = 0
          .Tank1Ready = False
          .KR.ACCommand_Cancel
          Cancel
        End If
      End If

    Case InManual
      StateString = "Drugroom Switch In Manual "
      DrugroomDisplay = "Switch In Manual"
      .Tank1Ready = False
      HeatPrepTimer = .Parameters_Tank1HeatPrepTimer
      If Not .IO_Tank1Manual_SW Then
        HeatPrepTimer = .Parameters_Tank1HeatPrepTimer
        State = Fill
      End If

  End Select

  Loop Until (StatePrev = State)    'Loop until state does not change

  End With
End Sub

Public Sub Cancel()
  FillOn = False
  State = Off
  HeatOn = False
  AddMixing = False
  AddDispenseError = False
  WaitReadyTimer.Pause
  AlarmRedyeIssue = False
  ManualAdd = False

  'This is to clear prepare properties and hopefully resolve issues in LA where wrong calloff is used
  AddCallOff = 0
  AddFillLevel = 0
  AddMixTime = 0
  AddRecipeStep = ""
  DesiredTemperature = 0
  DispenseCalloff = 0
  DispenseTank = 0
  Dispenseproducts = ""
  Dispensestate = 0

  DispenseDyes = False
  DispenseChems = False
  DispenseDyesOnly = False
  DispenseChemsOnly = False
  DispenseDyesChems = False

End Sub

Friend Sub ProgramStart()
On Error Resume Next
  Dispenseproducts = ""
  DispenseCalloff = 0
 ' DispenseTank = 0
  AddCallOff = 0
  AddRecipeStep = ""
  Dim i1 As Long, i2 As Long
  For i1 = 1 To 64
    For i2 = 1 To 8
      RecipeSteps(i1, i2) = ""
    Next i2
  Next i1
End Sub

Public Sub CopyTo(Target As acAddPrepare)
  With Target
    .State = State
    .AddFillLevel = AddFillLevel
    .AddCallOff = AddCallOff
    .OverrunTimer.TimeRemaining = OverrunTimer.TimeRemaining
    .AddMixing = AddMixing
    .DesiredTemperature = DesiredTemperature
    .AddMixTime = AddMixTime
    .OverrunTime = OverrunTime
    .MixTimer.TimeRemaining = MixTimer.TimeRemaining
    .WaitReadyTimer.TimeElapsed = WaitReadyTimer.TimeElapsed

    .ManualAdd = ManualAdd
    .AlarmRedyeIssue = AlarmRedyeIssue
    .AddDispenseError = AddDispenseError
    .AddRecipeStep = AddRecipeStep
    .DispenseCalloff = DispenseCalloff
    .DispenseTank = DispenseTank
    .Dispensestate = Dispensestate
    .Dispenseproducts = Dispenseproducts
    .Timer.TimeRemaining = Timer.TimeRemaining

    .DrugroomDisplay = DrugroomDisplay
    .DispenseDyesChems = DispenseDyesChems
    .DispenseDyesOnly = DispenseDyesOnly
    .DispenseChemsOnly = DispenseChemsOnly
    .DispenseTimer.TimeRemaining = DispenseTimer.TimeRemaining
    .MixTimer.TimeRemaining = MixTimer.TimeRemaining

  End With
End Sub

Public Property Get IsOn() As Boolean
  IsOn = (State <> Off)
End Property
Public Property Get IsWaitingToDispense() As Boolean
  IsWaitingToDispense = (State = DispenseWaitTurn)
End Property
Public Property Get IsDispensing() As Boolean
  IsDispensing = (State = DispenseWaitReady) Or (State = DispenseWaitProducts) Or (State = DispenseWaitResponse)
End Property
Friend Property Get IsDispenseWaitReady() As Boolean
  IsDispenseWaitReady = (State = DispenseWaitReady)
End Property
Friend Property Get IsDispenseReadyOverrun() As Boolean
  IsDispenseReadyOverrun = IsDispenseWaitReady And DispenseTimer.Finished
End Property
Friend Property Get IsDispenseWaitResponse() As Boolean
  IsDispenseWaitResponse = (State = DispenseWaitResponse)
End Property
Friend Property Get IsDispenseResponseOverrun() As Boolean
  IsDispenseResponseOverrun = IsDispenseWaitResponse And DispenseTimer.Finished
End Property

Public Property Get IsFill() As Boolean
  IsFill = ((State = Fill) Or FillOn) And (Not State = InManual)
End Property
Public Property Get IsHeating() As Boolean
  IsHeating = ((State = Heat) Or HeatOn) And (Not State = InManual)
End Property
Friend Property Get IsHeatTankOverrun() As Boolean
  IsHeatTankOverrun = ((State = Fill) Or (State = Heat)) And HeatPrepTimer.Finished
End Property
Public Property Get IsSlow() As Boolean
  IsSlow = (State = Slow)
End Property
Public Property Get IsFast() As Boolean
  IsFast = (State = Fast)
End Property
Friend Property Get IsWaitReady() As Boolean
  IsWaitReady = (State = Slow) Or (State = Fast) Or (State = MixForTime)
End Property
Public Property Get IsMixing() As Boolean
  IsMixing = (State = MixForTime)
End Property
Public Property Get IsReady() As Boolean
  IsReady = (State = Ready)
End Property
Public Property Get IsMixerOn() As Boolean
  IsMixerOn = AddMixing And Not ((State = InManual) Or IsDispensing Or IsWaitingToDispense Or IsWaitIdle)
End Property
Public Property Get IsOverrun() As Boolean
  IsOverrun = IsWaitReady And OverrunTimer.Finished
End Property
Friend Property Get IsInterlocked() As Boolean
 IsInterlocked = (State = KAInterlock)
End Property
Friend Property Get IsTransfer() As Boolean
  IsTransfer = (State = KATransfer1) Or (State = KATransfer2) Or (State = KARinseToDrain) Or (State = KATransferToDrain)
End Property
Friend Property Get IsTransferToAddition() As Boolean
 IsTransferToAddition = ((State = KATransfer1) Or (State = KARinse) Or _
                       (State = KATransfer2)) And (DispenseTank = 1)
End Property
Friend Property Get IsTransferToReserve() As Boolean
 IsTransferToReserve = ((State = KATransfer1) Or (State = KARinse) Or _
                       (State = KATransfer2)) And (DispenseTank = 2)
End Property
Friend Property Get IsTransferToDrain() As Boolean
 IsTransferToDrain = IsRinseToDrain Or (State = KATransferToDrain)
End Property
Friend Property Get IsRinse() As Boolean
  If (State = KARinse) Then IsRinse = True
End Property
Friend Property Get IsRinseToDrain() As Boolean
  If (State = KARinseToDrain) Then IsRinseToDrain = True
End Property
Friend Property Get IsPaused() As Boolean
  IsPaused = (State = KAPause)
End Property
Friend Property Get IsWaitIdle() As Boolean
  IsWaitIdle = (State = WaitIdle)
End Property
Friend Property Get IsDelayed() As Boolean
  IsDelayed = (IsWaitReady And IsOverrun)
End Property
Friend Property Get IsInManual() As Boolean
  IsInManual = (State = InManual)
End Property

#End If




#If 0 Then


'MW's Scholl1415 .NET code


<Command("Look Ahead", "", "", "", "", CommandType.ParallelCommand), _
Description("Signals the dispenser to start dispensing products for the next lot."), _
Category("Look Ahead Functions")> _
Public Class Command810 : Inherits MarshalByRefObject : Implements ACCommand

#Region "Enumeration"

  Public Enum EState
    Off
    CheckNextDyelot
    Start
    Run
    Done
  End Enum
#End Region

  Private ReadOnly ControlCode As ControlCode
  Public Sub New(ByVal controlCode As ControlCode)
    Me.ControlCode = controlCode
  End Sub
  Public Sub ParametersChanged(ByVal ParamArray param() As Integer) Implements ACCommand.ParametersChanged

  End Sub

  Public Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With ControlCode
      'Reset state & variables
      ResetAll()
      'Initialize
      State = EState.Off
      StatePrevious = EState.Off
      TankDEnabled = False
      LookAheadStopFound = False

      If Parameters_LookAheadEnabled = 1 Then State = EState.CheckNextDyelot

      Return True

    End With
  End Function

  Public Function Run() As Boolean Implements ACCommand.Run
    'If the command is not active then don't run this code
    If Not IsOn Then Exit Function

    With ControlCode

      Do
        'Stay in this loop until state changes have completed
        ' - stops odd effects from momentary state transitions
        StatePrevious = State
        Select Case State
          Case EState.Off


            'Check to see if a dyelot is schedule behind this one
          Case EState.CheckNextDyelot
            'Check every 60 seconds
            If CheckNextDyelotTimer.Finished Then
              CheckNextDyelotTimer.Seconds = 60
              If CheckNextDyelot() Then
                'If the next lot is not blocked then start preparing the tanks
                'If the next lot is not blocked then start preparing the tanks
                If (Not Blocked) OrElse (Parameters_LookAheadIgnoreBlocked = 1) Then
                  State = EState.Start

                End If
              End If
            End If

          Case EState.Start
            If TankDEnabled Then
              TankDPrepare.Start(TankDCallOff, .Command901.Parameters_StandardDispensePrepareTime, .Command901.Parameters_StandardManualPrepareTime, .Command901.Parameters_TankDPrefillLevel)
              .DispenseDyelot = Dyelot
              .DispenseRedye = Redye
            End If

            State = EState.Run

          Case EState.Run
            If .Command785.IsWaitReady Then
              WaitReady1 = True
            Else
              WaitReady1 = False
            End If


            If TankDEnabled Then TankDPrepare.Run(.AddTankDLevel, .AddTankDReady, WaitReady1, .Command785.IsAddTankTransferValve, Me.ControlCode, TankDCallOff)

            If Not (TankDPrepare.IsOn) Then State = EState.Done

          Case EState.Done
            Cancel()

        End Select
      Loop Until (State = StatePrevious)

      'Cancel this command if parameter is changed
      If Parameters_LookAheadEnabled <> 1 Then
        Cancel()
        Exit Function
      End If

    End With

  End Function


#Region "cancels"
  Public Sub Cancel() Implements ACCommand.Cancel
    If Not (TankDPrepare.IsOn) Then CancelAll()
  End Sub
  Public Sub CancelAll()
    TankDPrepare.Cancel()
    State = EState.Off
  End Sub
#End Region

#Region "Public properties"
  Public ReadOnly Property IsOn() As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property
  Public ReadOnly Property IsActive() As Boolean
    Get
      Return Not ((State = EState.Off) OrElse (State = EState.CheckNextDyelot))
    End Get
  End Property

#End Region

#Region "functions and subs"

  Private Sub ResetAll()
    'Reset command state & variables
    Dyelot = ""
    Redye = 0
    Blocked = False
    Program = ""
    IPKP1 = ""
    TankDCallOff = 0
    TankDEnabled = False
  End Sub

  Private Function CheckNextDyelot() As Boolean
    Dim sql As String = Nothing
    Try
      With ControlCode

        'Get dyelot 
        sql = "select top(1) dyelot,redye,program,blocked,batched from dyelots where state is null order by starttime"
        Dim dtdyelot As System.Data.DataTable = .Parent.DbGetDataTable(sql)

        'Make sure we have one row
        If dtdyelot Is Nothing OrElse dtdyelot.Rows.Count <> 1 Then Return False

        Dim drdyelot As System.Data.DataRow = dtdyelot.Rows(0)

        Dyelot = Null.NullToEmptyString(drdyelot("Dyelot"))
        Redye = Null.NullToZeroInteger(drdyelot("Redye"))
        Blocked = Null.NullToFalse(drdyelot("Blocked"))
        Batched = Not drdyelot.IsNull("Batched")
        Program = Null.NullToEmptyString(drdyelot("Program"))

        ' The first batch ready to run on this machine (if any) must be not blocked and must have the Batched time set to be considered
        If Blocked Or Not Batched Then Return False

        ' Dim num As Integer
        'If Not Integer.TryParse(Program, num) Then Return False

        'get program steps
        sql = "SELECT Steps FROM Programs WHERE ProgramNumber='" & Program & "'"
        Dim ProgramSteps As String
        Dim dtprogram As System.Data.DataTable = .Parent.DbGetDataTable(sql)
        If dtprogram Is Nothing OrElse dtprogram.Rows.Count <> 1 Then Return False

        Dim drprogram As System.Data.DataRow = dtprogram.Rows(0)

        ProgramSteps = Null.NullToEmptyString(drprogram("Steps"))

        Return CheckNextProgram(ProgramSteps)

      End With


    Catch ex As Exception
      Utilities.Log.LogError(ex, sql)

    End Try

    Return False
  End Function

  Private Function CheckNextProgram(ByRef ProgramSteps As String) As Boolean

    Try
      With ControlCode


        'Make sure we've got something to check
        If ProgramSteps.Length <= 0 Then Return False

        'These separators are used by programs and prefixed steps
        Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
        Dim separator1 As String = Convert.ToChar(255)
        Dim separator2 As String = ","

        ' Use this to split out steps in this program
        Dim Steps() As String
        Dim Parameters() As String
        Dim checkingfornotes() As String

        'Split the program string into an array of program steps - one step per array element
        Steps = ProgramSteps.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

        'Make sure we've have something to check
        If Steps.GetUpperBound(0) <= 0 Then Return False

        'figure out if kp1 found only need command 901
        Dim IP As Integer
        Dim prompt As Integer
        Dim CommandString As String, CommandCode() As String

        For i As Integer = 0 To Steps.GetUpperBound(0)
          If LookAheadStopFound = True Then Exit For
          If TankDEnabled = True Then Exit For
          CommandString = Steps(i)
          CommandCode = CommandString.Split(separator2.ToCharArray)
          'look for the commands we want
          Select Case CommandCode(0)

            'Obviously we would normally be looking at more commands so I left the Case statement in
            Case "IP"
              'get the program number
              checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
              Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array
              If Parameters.GetUpperBound(0) >= 2 Then
                IP = CInt(Parameters(2)) * 256 + CInt(Parameters(1))
              Else
                IP = CInt(Parameters(1))
              End If

              CheckIPProgram(IP)

            Case "702" 'brine fill
              'get the program number
              checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
              Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array

              If Parameters.GetUpperBound(0) >= 1 Then
                prompt = CInt(Parameters(1))
              End If

              If prompt <> 99 Then 'if its not 99 then change it.
                TankDCallOff = TankDCallOff + 1
              End If

            Case "805" 'prepare ed
              'get the program number
              checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
              Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array

              If Parameters.GetUpperBound(0) >= 1 Then
                prompt = CInt(Parameters(1))
              End If

              If prompt <> 99 Then 'if its not 99 then change it.
                TankDCallOff = TankDCallOff + 1
              End If

            Case "821" 'prepare ed
              'get the program number
              checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
              Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array

              If Parameters.GetUpperBound(0) >= 1 Then
                prompt = CInt(Parameters(1))
              End If

              If prompt <> 99 Then 'if its not 99 then change it.
                TankDCallOff = TankDCallOff + 1
              End If

            Case "901" 'd tank prepare
              'get the program number
              checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
              Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array

              If Parameters.GetUpperBound(0) >= 1 Then
                prompt = CInt(Parameters(1))
              End If

              If prompt <> 99 Then 'if its not 99 then change it.
                TankDCallOff = TankDCallOff + 1
              Else
                TankDCallOff = 99
              End If
              TankDEnabled = True
              Exit For 'we found it we are done.

            Case "811"
              LookAheadStopFound = True
              Exit For

          End Select
        Next i

      End With

      Return True

    Catch ex As Exception
      Utilities.Log.LogError(ex, "error in check program")

    End Try

    Return False

  End Function

  'this is for first IP
  Private Function CheckIPProgram(ByRef ProgramNumber As Integer) As Boolean
    Dim sql As String = Nothing

    Try
      With ControlCode


        'get program steps
        sql = "SELECT Steps FROM Programs WHERE ProgramNumber='" & ProgramNumber & "'"

        Dim ProgramSteps As String
        Dim dtprogram As System.Data.DataTable = .Parent.DbGetDataTable(sql)
        If dtprogram Is Nothing OrElse dtprogram.Rows.Count <> 1 Then Return False

        Dim drprogram As System.Data.DataRow = dtprogram.Rows(0)

        ProgramSteps = Null.NullToEmptyString(drprogram("Steps"))

        Return CheckIPProgramSteps(ProgramSteps)

      End With

    Catch ex As Exception
      Utilities.Log.LogError(ex, "error with checking the IP")

    End Try

    Return False
  End Function

  'check program steps for the first ip
  Private Function CheckIPProgramSteps(ByRef ProgramSteps As String) As Boolean

    Try

      'Make sure we've got something to check
      If ProgramSteps.Length <= 0 Then Return False

      'These separators are used by programs and prefixed steps
      Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
      Dim separator1 As String = Convert.ToChar(255)
      Dim separator2 As String = ","

      ' Use this to split out steps in this program
      Dim Steps() As String
      Dim Parameters() As String
      Dim checkingfornotes() As String

      'Split the program string into an array of program steps - one step per array element
      Steps = ProgramSteps.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

      'Make sure we've have something to check
      If Steps.GetUpperBound(0) <= 0 Then Return False

      Dim IP2 As Integer
      Dim CommandString As String, CommandCode() As String
      Dim prompt As Integer

      'Loop through the program steps to find the first KP command for each tank
      For i As Integer = 0 To Steps.GetUpperBound(0)

        CommandString = Steps(i)
        CommandCode = CommandString.Split(separator2.ToCharArray)

        Select Case CommandCode(0)

          Case "IP"
            'get the program number
            checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array
            If Parameters.GetUpperBound(0) >= 2 Then
              IP2 = CInt(Parameters(2)) * 256 + CInt(Parameters(1))
            Else
              IP2 = CInt(Parameters(1))
            End If
            CheckIP2Program(IP2)


          Case "702" 'brine fill
            'get the program number
            checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array

            If Parameters.GetUpperBound(0) >= 1 Then
              prompt = CInt(Parameters(1))
            End If

            If prompt <> 99 Then 'if its not 99 then change it.
              TankDCallOff = TankDCallOff + 1
            End If

          Case "805" 'prepare ed
            'get the program number
            checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array

            If Parameters.GetUpperBound(0) >= 1 Then
              prompt = CInt(Parameters(1))
            End If

            If prompt <> 99 Then 'if its not 99 then change it.
              TankDCallOff = TankDCallOff + 1
            End If

          Case "821" 'prepare ed
            'get the program number
            checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array

            If Parameters.GetUpperBound(0) >= 1 Then
              prompt = CInt(Parameters(1))
            End If

            If prompt <> 99 Then 'if its not 99 then change it.
              TankDCallOff = TankDCallOff + 1
            End If

          Case "901" 'd tank prepare
            'get the program number
            checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array

            If Parameters.GetUpperBound(0) >= 1 Then
              prompt = CInt(Parameters(1))
            End If

            If prompt <> 99 Then 'if its not 99 then change it.
              TankDCallOff = TankDCallOff + 1
            Else
              TankDCallOff = 99
            End If

            TankDEnabled = True
            Exit For 'we found it we are done.

          Case "811"
            LookAheadStopFound = True
            Exit For

        End Select
      Next i

      'Everything completed okay
      Return True

    Catch ex As Exception
      Utilities.Log.LogError(ex, "error with checking IP program steps")

    End Try

    Return False
  End Function

  'this is for a IP inside the first IP
  Private Function CheckIP2Program(ByRef ProgramNumber As Integer) As Boolean
    Try

      With ControlCode

        'get program steps
        Dim sql As String = "SELECT Steps FROM Programs WHERE ProgramNumber='" & ProgramNumber & "'"

        Dim dtprogram As System.Data.DataTable = .Parent.DbGetDataTable(sql)
        If dtprogram Is Nothing OrElse dtprogram.Rows.Count <> 1 Then Return False

        Dim drprogram As System.Data.DataRow = dtprogram.Rows(0)

        Dim ProgramSteps As String = Null.NullToEmptyString(drprogram("Steps"))

        Return CheckIP2ProgramSteps(ProgramSteps)

      End With


    Catch ex As Exception
      Utilities.Log.LogError(ex, "error with checking IP program steps")

    End Try

    Return False

  End Function

  Private Function CheckIP2ProgramSteps(ByRef ProgramSteps As String) As Boolean
    Try

      'Make sure we've got something to check
      If ProgramSteps.Length <= 0 Then Return False

      'These separators are used by programs and prefixed steps
      Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
      Dim separator1 As String = Convert.ToChar(255)
      Dim separator2 As String = ","

      ' Use this to split out steps in this program
      Dim Steps() As String
      Dim Parameters() As String
      Dim checkingfornotes() As String

      'Split the program string into an array of program steps - one step per array element
      Steps = ProgramSteps.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

      'Make sure we've have something to check
      If Steps.GetUpperBound(0) <= 0 Then Return False

      'figure out if kp1 found,kp2 found, batch weight found, etc....
      Dim IP3 As Integer
      Dim CommandString As String, CommandCode() As String
      Dim prompt As Integer

      'Loop through the program steps to find the first KP command for each tank
      For i As Integer = 0 To Steps.GetUpperBound(0)

        CommandString = Steps(i)
        CommandCode = CommandString.Split(separator2.ToCharArray)

        Select Case CommandCode(0)
          'Obviously we would normally be looking at more commands so I left the Case statement in
          Case "IP"
            'get the program number
            checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array
            If Parameters.GetUpperBound(0) >= 2 Then
              IP3 = CInt(Parameters(2)) * 256 + CInt(Parameters(1))
            Else
              IP3 = CInt(Parameters(1))
            End If
            CheckIP3Program(IP3)


          Case "702" 'brine fill
            'get the program number
            checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array

            If Parameters.GetUpperBound(0) >= 1 Then
              prompt = CInt(Parameters(1))
            End If

            If prompt <> 99 Then 'if its not 99 then change it.
              TankDCallOff = TankDCallOff + 1
            End If

          Case "805" 'prepare ed
            'get the program number
            checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array

            If Parameters.GetUpperBound(0) >= 1 Then
              prompt = CInt(Parameters(1))
            End If

            If prompt <> 99 Then 'if its not 99 then change it.
              TankDCallOff = TankDCallOff + 1
            End If

          Case "821" 'prepare ed
            'get the program number
            checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array

            If Parameters.GetUpperBound(0) >= 1 Then
              prompt = CInt(Parameters(1))
            End If

            If prompt <> 99 Then 'if its not 99 then change it.
              TankDCallOff = TankDCallOff + 1
            End If


          Case "901" 'd tank prepare
            'get the program number
            checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array

            If Parameters.GetUpperBound(0) >= 1 Then
              prompt = CInt(Parameters(1))
            End If

            If prompt <> 99 Then 'if its not 99 then change it.
              TankDCallOff = TankDCallOff + 1
            Else
              TankDCallOff = 99
            End If
            TankDEnabled = True
            Exit For 'we found it we are done.

          Case "811"
            LookAheadStopFound = True
            Exit For

        End Select
      Next i



    Catch ex As Exception
      Utilities.Log.LogError(ex, "error with checking IP program steps")

    End Try

    Return False
  End Function


  'this is for a IP inside the second IP
  Private Function CheckIP3Program(ByRef ProgramNumber As Integer) As Boolean
    Try

      With ControlCode

        'get program steps
        Dim sql As String = "SELECT Steps FROM Programs WHERE ProgramNumber='" & ProgramNumber & "'"

        Dim dtprogram As System.Data.DataTable = .Parent.DbGetDataTable(sql)
        If dtprogram Is Nothing OrElse dtprogram.Rows.Count <> 1 Then Return False

        Dim drprogram As System.Data.DataRow = dtprogram.Rows(0)

        Dim ProgramSteps As String = Null.NullToEmptyString(drprogram("Steps"))

        Return CheckIP3ProgramSteps(ProgramSteps)

      End With


    Catch ex As Exception
      Utilities.Log.LogError(ex, "error with checking IP program steps")

    End Try

    Return False

  End Function
  Private Function CheckIP3ProgramSteps(ByRef ProgramSteps As String) As Boolean
    Try

      'Make sure we've got something to check
      If ProgramSteps.Length <= 0 Then Return False

      'These separators are used by programs and prefixed steps
      Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
      Dim separator1 As String = Convert.ToChar(255)
      Dim separator2 As String = ","

      ' Use this to split out steps in this program
      Dim Steps() As String
      Dim Parameters() As String
      Dim checkingfornotes() As String

      'Split the program string into an array of program steps - one step per array element
      Steps = ProgramSteps.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

      'Make sure we've have something to check
      If Steps.GetUpperBound(0) <= 0 Then Return False

      'figure out if kp1 found,kp2 found, batch weight found, etc....
      Dim IP4 As Integer
      Dim CommandString As String, CommandCode() As String
      Dim prompt As Integer

      'Loop through the program steps to find the first KP command for each tank
      For i As Integer = 0 To Steps.GetUpperBound(0)

        CommandString = Steps(i)
        CommandCode = CommandString.Split(separator2.ToCharArray)

        Select Case CommandCode(0)
          'Obviously we would normally be looking at more commands so I left the Case statement in
          Case "IP"
            'get the program number
            checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array
            If Parameters.GetUpperBound(0) >= 2 Then
              IP4 = CInt(Parameters(2)) * 256 + CInt(Parameters(1))
            Else
              IP4 = CInt(Parameters(1))
            End If
            CheckIP4Program(IP4)


          Case "702" 'brine fill
            'get the program number
            checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array

            If Parameters.GetUpperBound(0) >= 1 Then
              prompt = CInt(Parameters(1))
            End If

            If prompt <> 99 Then 'if its not 99 then change it.
              TankDCallOff = TankDCallOff + 1
            End If

          Case "805" 'prepare ed
            'get the program number
            checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array

            If Parameters.GetUpperBound(0) >= 1 Then
              prompt = CInt(Parameters(1))
            End If

            If prompt <> 99 Then 'if its not 99 then change it.
              TankDCallOff = TankDCallOff + 1
            End If

          Case "821" 'prepare ed
            'get the program number
            checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array

            If Parameters.GetUpperBound(0) >= 1 Then
              prompt = CInt(Parameters(1))
            End If

            If prompt <> 99 Then 'if its not 99 then change it.
              TankDCallOff = TankDCallOff + 1
            End If

          Case "901" 'd tank prepare
            'get the program number
            checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array

            If Parameters.GetUpperBound(0) >= 1 Then
              prompt = CInt(Parameters(1))
            End If

            If prompt <> 99 Then 'if its not 99 then change it.
              TankDCallOff = TankDCallOff + 1
            Else
              TankDCallOff = 99
            End If
            TankDEnabled = True
            Exit For 'we found it we are done.

          Case "811"
            LookAheadStopFound = True
            Exit For

        End Select
      Next i



    Catch ex As Exception
      Utilities.Log.LogError(ex, "error with checking IP program steps")

    End Try

    Return False
  End Function

  'this is for a IP inside the third IP
  Private Function CheckIP4Program(ByRef ProgramNumber As Integer) As Boolean
    Try

      With ControlCode

        'get program steps
        Dim sql As String = "SELECT Steps FROM Programs WHERE ProgramNumber='" & ProgramNumber & "'"

        Dim dtprogram As System.Data.DataTable = .Parent.DbGetDataTable(sql)
        If dtprogram Is Nothing OrElse dtprogram.Rows.Count <> 1 Then Return False

        Dim drprogram As System.Data.DataRow = dtprogram.Rows(0)

        Dim ProgramSteps As String = Null.NullToEmptyString(drprogram("Steps"))

        Return CheckIP4ProgramSteps(ProgramSteps)

      End With


    Catch ex As Exception
      Utilities.Log.LogError(ex, "error with checking IP program steps")

    End Try

    Return False

  End Function
  Private Function CheckIP4ProgramSteps(ByRef ProgramSteps As String) As Boolean
    Try

      'Make sure we've got something to check
      If ProgramSteps.Length <= 0 Then Return False

      'These separators are used by programs and prefixed steps
      Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
      Dim separator1 As String = Convert.ToChar(255)
      Dim separator2 As String = ","

      ' Use this to split out steps in this program
      Dim Steps() As String
      Dim Parameters() As String
      Dim checkingfornotes() As String

      'Split the program string into an array of program steps - one step per array element
      Steps = ProgramSteps.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

      'Make sure we've have something to check
      If Steps.GetUpperBound(0) <= 0 Then Return False

      'figure out if kp1 found,kp2 found, batch weight found, etc....
      Dim CommandString As String, CommandCode() As String
      Dim prompt As Integer
      'Loop through the program steps to find the first KP command for each tank
      For i As Integer = 0 To Steps.GetUpperBound(0)

        CommandString = Steps(i)
        CommandCode = CommandString.Split(separator2.ToCharArray)

        Select Case CommandCode(0)
          'Obviously we would normally be looking at more commands so I left the Case statement in


          Case "702" 'brine fill
            'get the program number
            checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array

            If Parameters.GetUpperBound(0) >= 1 Then
              prompt = CInt(Parameters(1))
            End If

            If prompt <> 99 Then 'if its not 99 then change it.
              TankDCallOff = TankDCallOff + 1
            End If

          Case "805" 'prepare ed
            'get the program number
            checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array

            If Parameters.GetUpperBound(0) >= 1 Then
              prompt = CInt(Parameters(1))
            End If

            If prompt <> 99 Then 'if its not 99 then change it.
              TankDCallOff = TankDCallOff + 1
            End If

          Case "821" 'prepare ed
            'get the program number
            checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array

            If Parameters.GetUpperBound(0) >= 1 Then
              prompt = CInt(Parameters(1))
            End If

            If prompt <> 99 Then 'if its not 99 then change it.
              TankDCallOff = TankDCallOff + 1
            End If

          Case "901" 'd tank prepare
            'get the program number
            checkingfornotes = Steps(i).Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            Parameters = checkingfornotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries) '0 based array

            If Parameters.GetUpperBound(0) >= 1 Then
              prompt = CInt(Parameters(1))
            End If

            If prompt <> 99 Then 'if its not 99 then change it.
              TankDCallOff = TankDCallOff + 1
            Else
              TankDCallOff = 99
            End If
            TankDEnabled = True
            Exit For 'we found it we are done.

          Case "811"
            LookAheadStopFound = True
            Exit For

        End Select
      Next i



    Catch ex As Exception
      Utilities.Log.LogError(ex, "error with checking IP program steps")

    End Try

    Return False
  End Function

#End Region

#Region "Instances"
  Public TankDPrepare As New AddPrepareDTank(Me.ControlCode)
#End Region


#Region "Variables"

  Property WaitReady1() As Boolean
  Property IPKP1() As String

  Private job_ As String
  Public Property Job() As String
    Get
      Return job_
    End Get
    Private Set(ByVal value As String)
      job_ = Dyelot
      If Redye > 0 Then job_ = Dyelot & "@" & Redye
    End Set
  End Property

  Property Dyelot As String
  Property Redye As Integer
  Property Blocked As Boolean
  Property Batched As Boolean
  Property Program As String
  Property TankDEnabled As Boolean
  Property TankDCallOff As Integer
  Property LookAheadStopFound() As Boolean

#End Region


#Region "parameters"
  <Parameter(0, 1), Category("Look Ahead"), _
    Description("Set to one to enable the look ahead function.")> _
  Public Parameters_LookAheadEnabled As Integer
  <Parameter(0, 1), Category("Look Ahead"), _
    Description("Set to one to enable the look ahead function for blocked dyelots.")> _
  Public Parameters_LookAheadIgnoreBlocked As Integer
#End Region

  Shared Function bcdtodecimal(bcd As Integer) As Integer
    Dim ret As Integer
    Dim mult = 1

    Do While bcd <> 0
      ret = ret + (bcd Mod 16) * mult
      mult = mult * 10
      bcd = bcd \ 16
    Loop
    Return ret
  End Function
End Class
#End If













End Class
