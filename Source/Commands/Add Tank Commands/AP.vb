'American & Efird - Mt. Holly Then Platform to GMX
' Version 2024-09-17

Imports Utilities.Translations

<Command("Add Prepare", "Msg:|0-99|", "", "", ""), _
TranslateCommand("es", "A�adir preparar", "Msg:|0-99|"), _
Description("Signal the operator to prepare the Addition tank."), _
TranslateDescription("es", "El operador para preparar el dep�sito adem�s de la se�al."),
Category("Add Tank Commands"), TranslateCategory("es", "Add Tank Commands")>
Public Class AP : Inherits MarshalByRefObject : Implements ACCommand

  'Command states
  Public Enum EState
    Off
    WaitIdle
    Slow
    Fast
    Ready
  End Enum
  Public State As EState
  Public Status As String
  Public Message As Integer
  Public Calloff As Integer
  Private flashSlow_ As Boolean
  Private flashFast_ As Boolean

  Friend ReadOnly Timer As New Timer
  Public ReadOnly Property TimeAdvance As Integer
    Get
      Return Timer.Seconds
    End Get
  End Property
  Friend Property TimerOverrun As New Timer
  Public ReadOnly Property TimeOverrun As Integer
    Get
      Return TimerOverrun.Seconds
    End Get
  End Property

  'Keep a local reference to the control code object for convenience
  Private ReadOnly controlCode As ControlCode

  Public Sub New(ByVal controlCode As ControlCode)
    Me.controlCode = controlCode
    Me.controlCode.Commands.Add(Me)
  End Sub

  Public Sub ParametersChanged(ByVal ParamArray param() As Integer) Implements ACCommand.ParametersChanged
    'If the command is not running then do nothing the changes will be picked up when the command starts
    If Not IsOn Then Exit Sub
    ' Do nothing
  End Sub

  Public Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With controlCode

      ' Cancel addition tank commands
      .AD.Cancel()

      'Command parameters
      Message = 0

      'Check array bounds just to be on the safe side
      If param.GetUpperBound(0) >= 1 Then Message = param(1)

      'Get CallOff Value
      Calloff = Utilities.Programs.GetCallOffNumber(.Parent.CurrentStep, .Parent.Programs)

      'Set default state
      State = EState.WaitIdle
      Timer.Seconds = 2

      'Set overrun timer
      TimerOverrun.Minutes = .Parameters.StandardTimeAddPrepare

      'Tell the control system to step on
      Return True
    End With
  End Function

  Public Function Run() As Boolean Implements ACCommand.Run
    With controlCode
      flashFast_ = .FlashFast
      flashSlow_ = .FlashSlow


      Select Case State
        Case EState.Off
          Status = ("")

        Case EState.WaitIdle
          Status = Translate("Wait Idle")
          If Not .IO.AdvancePb Then
            State = EState.Slow
            Timer.Seconds = 2
          End If

        Case EState.Slow
          Status = Translate("Wait Ready")
          If .AC.IsWaitReady Then State = EState.Fast
          If .AT.IsWaitReady Then State = EState.Fast
          If Not .AdvancePb Then
            Timer.Seconds = 2
          Else
            Status = Translate("Hold Advance") & Timer.ToString(1)
            If Timer.Finished Then .AddReady = True
          End If
          If .AddReady Then State = EState.Ready

        Case EState.Fast
          If Not .AdvancePb Then
            Timer.Seconds = 2
          Else
            Status = Translate("Hold Advance") & Timer.ToString(1)
            If Timer.Finished Then .AddReady = True
          End If
          If .AddReady Then State = EState.Ready

        Case EState.Ready
          Status = Translate("Add Ready")
          If Not .AddReady Then State = EState.Slow

      End Select
    End With
  End Function

  Public Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
    Timer.Cancel()
    TimerOverrun.Cancel()
    Status = ""
    Calloff = 0
  End Sub

  Public ReadOnly Property IsOn() As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property

  Public ReadOnly Property IsOverrun() As Boolean
    Get
      Return IsOn AndAlso IsWaitReady AndAlso TimerOverrun.Finished
    End Get
  End Property

  Public ReadOnly Property IsWaitReady() As Boolean
    Get
      Return (State = EState.Slow) OrElse (State = EState.Fast)
    End Get
  End Property

  Public ReadOnly Property IoAddLamp() As Boolean
    Get
      Return (State = EState.Ready) OrElse ((State = EState.Slow) And flashSlow_) OrElse ((State = EState.Fast) And flashFast_)
    End Get
  End Property



  ' Add Preapre
#If 0 Then
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acAddPrepare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
      If AlarmRedyeIssue And (.Parameters.EnableRedyeIssueAlarm = 1) Then
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
      StateString = "Tank 1 PreFill to " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters.Tank1PreFillLevel / 10, "0", 3) & "%"
      DrugroomDisplay = "Tank 1 PreFill to " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters.Tank1PreFillLevel / 10, "0", 3) & "%"
      If (.Tank1Level >= .Parameters.Tank1PreFillLevel) Then
         State = DispenseWaitTurn
      End If
      
    Case DispenseWaitTurn    'Wait for other tank to be finished with dispensing before new dispense
      StateString = "Wait for Turn "
      DrugroomDisplay = "Wait for Turn"
      If (AddCallOff = 0) Or (.Parameters.DispenseEnabled <> 1) Then
        HeatPrepTimer = .Parameters.Tank1HeatPrepTimer
        State = Fill
        ManualAdd = True
      Else
        DispenseTimer = .Parameters.DispenseReadyDelayTime * 60
        State = DispenseWaitReady
        DispenseCalloff = 0
      End If
      
    Case DispenseWaitReady    'Wait to make sure dispenser has completed previous dispense and is ready
      StateString = "Wait for Dispenser Ready "
      DrugroomDisplay = "Wait for Dispenser Ready "
      If Dispensestate = DispenseReady Then                                   '101
         State = DispenseWaitProducts
         DispenseTimer = .Parameters.DispenseResponseDelayTime * 60
         DispenseCalloff = AddCallOff
         .DispenseTank = DispenseTank
      End If
      If (AddCallOff = 0) Or (.Parameters.DispenseEnabled <> 1) Then
        HeatPrepTimer = .Parameters.Tank1HeatPrepTimer
        State = Fill
        DispenseCalloff = 0
        ManualAdd = True
      End If
          
    Case DispenseWaitProducts     'Wait for Dispensestate & DispenseProducts String to be set to determine how to signal delays
      StateString = "Wait for Dispense Products "
      DrugroomDisplay = "Wait for Dispense Products "
      DispenseTimer = .Parameters.DispenseResponseDelayTime * 60
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
      If (AddCallOff = 0) Or (.Parameters.DispenseEnabled <> 1) Then
        HeatPrepTimer = .Parameters.Tank1HeatPrepTimer
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
          HeatPrepTimer = .Parameters.Tank1HeatPrepTimer
          State = Fill
          DispenseCalloff = 0
          ManualAdd = True

        Case DispenseError                                                   '309
          AddDispenseError = True
          HeatPrepTimer = .Parameters.Tank1HeatPrepTimer
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
        StateString = "Tank 1 Level Too Low to Heat " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters.Tank1MixerOnLevel / 10, "0", 3) & "%"
        DrugroomDisplay = "Level Too Low to Heat " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters.Tank1MixerOnLevel / 10, "0", 3) & "%"
      Else
        StateString = "Tank 1 Heat to " & Pad(.IO_Tank1Temp / 10, "0", 3) & "F / " & Pad(DesiredTemperature / 10, "0", 3)
        DrugroomDisplay = "Heat to " & Pad(.IO_Tank1Temp / 10, "0", 3) & "F / " & Pad(DesiredTemperature / 10, "0", 3)
      End If
      If .IO_Tank1Manual_SW Then State = InManual
      If AddFillLevel = 0 Then State = Slow
      ' Heat Start/Stop
      If .IO_Tank1Temp < DesiredTemperature Then
        HeatOn = True
        AddMixing = True ' Must mix if heating
      End If
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
      ' Heat Start/Stop
      If .IO_Tank1Temp < DesiredTemperature Then
        HeatOn = True
        AddMixing = True ' Must mix if heating
      End If
      If .IO_Tank1Temp >= DesiredTemperature Then HeatOn = False
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
      ' Heat Start/Stop
      If .IO_Tank1Temp < DesiredTemperature Then
        HeatOn = True
        AddMixing = True ' Must mix if heating
      End If
      If .IO_Tank1Temp >= DesiredTemperature Then HeatOn = False
      If WaitReadyTimer.IsPaused Then WaitReadyTimer.Restart 'should cover LA copy to KP function
      If .Tank1Ready Then
        If .Tank1Level Then
          WaitReadyTimer.Pause
          State = MixForTime
          MixTimer = AddMixTime
        Else
          StateString = "Tank 1 Level low " & Pad(.Tank1Level / 10, "0", 3) & "%"
        End If
      End If

      
    Case MixForTime
      If Not .DrugroomMixerOn Then
        StateString = "Tank 1 Level Too Low to Mix " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters.Tank1MixerOnLevel / 10, "0", 3) & "%"
        DrugroomDisplay = "Level Too Low to Mix " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters.Tank1MixerOnLevel / 10, "0", 3) & "%"
      Else
        StateString = "Tank 1 Mix For Time " & TimerString(MixTimer.TimeRemaining)
        DrugroomDisplay = "Mix For Time " & TimerString(MixTimer.TimeRemaining)
      End If
      If .IO_Tank1Manual_SW Then State = InManual
      If AddFillLevel = 0 Then State = Ready
      ' Heat Start/Stop
      If .IO_Tank1Temp < DesiredTemperature Then
        HeatOn = True
        AddMixing = True ' Must mix if heating
      End If
      If .IO_Tank1Temp >= DesiredTemperature Then HeatOn = False
      If MixTimer.Finished Then State = Ready
      If Not .Tank1Ready Then
        State = Fast
        WaitReadyTimer.Restart
      End If
      
    Case Ready
      StateString = "Tank 1 Ready "
      DrugroomDisplay = "Ready "
      If .IO_Tank1Manual_SW Then State = InManual
      If (DispenseTank = 1) And (.AdditionLevel > .Parameters.AdditionMaxTransferLevel) Then
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
      If (.AdditionLevel <= .Parameters.AdditionMaxTransferLevel) Then
        State = KATransfer1
      End If
    
    Case KATransfer1
      If (.Tank1Level > 10) Then
        Timer = .Parameters.Tank1TimeBeforeRinse
        StateString = "Tank 1 Transfer Empty " & Pad(.Tank1Level / 10, "0", 3) & "%"
        DrugroomDisplay = "Transfer Empty " & Pad(.Tank1Level / 10, "0", 3) & "%"
      Else
        StateString = "Tank 1 Transfer Empty " & TimerString(Timer.TimeRemaining)
        DrugroomDisplay = "Transfer Empty " & TimerString(Timer.TimeRemaining)
      End If
      If .IO_Tank1Manual_SW Then State = InManual
      If Timer.Finished Then
        NumberOfRinses = .Parameters.DrugroomRinses
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
            Timer = .Parameters.Tank1RinseToDrainTime
          End If
        Else
          State = KARinse
          Timer = .Parameters.Tank1RinseTime
        End If
      End If
            
    Case KARinse
      StateString = "Tank 1 Rinse " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters.Tank1RinseLevel / 10, "0", 3) & "%"
      DrugroomDisplay = "Rinse " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters.Tank1RinseLevel / 10, "0", 3) & "%"
      If .IO_Tank1Manual_SW Then State = InManual
      If (.Tank1Level > .Parameters.Tank1RinseLevel) Then
        State = KATransfer2
        Timer = .Parameters.Tank1TimeAfterRinse
      End If
      
    Case KATransfer2
      If (.Tank1Level > 10) Then
        Timer = .Parameters.Tank1TimeAfterRinse
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
        Timer = .Parameters.Tank1RinseToDrainTime
      End If
      
    Case KARinseToDrain
      StateString = "Tank 1 Rinse To Drain " & TimerString(Timer.TimeRemaining)
      DrugroomDisplay = "Rinse To Drain " & TimerString(Timer.TimeRemaining)
      If .IO_Tank1Manual_SW Then State = InManual
      If Timer.Finished Or (.Tank1Level >= 500) Then
        State = KATransferToDrain
        Timer = .Parameters.Tank1DrainTime
        NumberOfRinses = NumberOfRinses - 1
      End If
    
    Case KATransferToDrain
      If (.Tank1Level > 10) Then
        Timer = .Parameters.Tank1DrainTime
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
          Timer = .Parameters.Tank1RinseToDrainTime
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
      HeatPrepTimer = .Parameters.Tank1HeatPrepTimer
      If Not .IO_Tank1Manual_SW Then
        HeatPrepTimer = .Parameters.Tank1HeatPrepTimer
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
  IsHeating = (HeatOn And ((State = Heat) Or (State = Slow) Or (State = Fast))) And _
                (Not State = InManual)
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



  ' AP.vb
#If 0 Then
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===============================================================================================
'AP - Addition tank prepare command
'===============================================================================================
Option Explicit
Implements ACCommand
Public Enum APState
  Off
  WaitIdle
  Slow
  Fast
  Ready
End Enum
Public State As APState
Public StateString As String
Public AddPreparePrompt As Long
Public OverrunTimer As New acTimer
Public AdvanceTimer As New acTimer
  
Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Parameters=Prompt |0-99|\r\nName=Add Prepare\r\nHelp=Signals the operator to prepare the addition tank."
Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  AddPreparePrompt = Param(1)
  With ControlCode
     .AddReadyFromPB = False
  End With
  
'Carry on my wayward son...
  StepOn = True
  State = WaitIdle
  AdvanceTimer.TimeRemaining = 2
  
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)

'Early bind ControlCode for speed and convenience
  Dim ControlCode As ControlCode
  Set ControlCode = ControlObject

  With ControlCode
    Select Case State
    
      Case Off
        StateString = ""
        
      Case WaitIdle
        StateString = "AP: Wait to lose Run Pushbutton "
        If Not .IO_RemoteRun Then
          OverrunTimer.TimeRemaining = .Parameters.StandardLocalPrepTime
          AdvanceTimer.TimeRemaining = 2
          State = Slow
        End If
       
      Case Slow
        StateString = "AP: Wait For Add Ready "
        If Not .IO_RemoteRun Then AdvanceTimer.TimeRemaining = 2
        If .AT.IsWaitReady Then State = Fast
        If .AddReadyFromPB Or AdvanceTimer.Finished Then
          .AddReady = True
          State = Ready
        End If
          
      Case Fast
        StateString = "AP: Wait For Add Ready "
        If .AddReadyFromPB Or AdvanceTimer.Finished Then
          .AddReady = True
          State = Ready
        End If
       
      Case Ready
        StateString = "AP: Add Ready "
        
    End Select
  End With
  
End Sub
Public Sub ACCommand_Cancel()
  State = Off
End Sub
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property
Friend Property Get IsOn() As Boolean
  If (State > Off) Then IsOn = True
End Property
Friend Property Get IsActive() As Boolean
  IsActive = IsOn
End Property
Friend Property Get IsSlow() As Boolean
  If (State = Slow) Then IsSlow = True
End Property
Friend Property Get IsFast() As Boolean
  If (State = Fast) Then IsFast = True
End Property
Friend Property Get IsWaitReady() As Boolean
  If IsSlow Or IsFast Then IsWaitReady = True
End Property
Friend Property Get IsReady() As Boolean
  If (State = Ready) Then IsReady = True
End Property
Public Property Get IsOverrun() As Boolean
  If OverrunTimer.Finished And (State = Fast) Then IsOverrun = True
End Property

#End If

#If 0 Then
  ' TODO REmove
  
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===============================================================================================
'AP - Addition tank prepare command
'===============================================================================================
Option Explicit
Implements ACCommand
Public Enum APState
  Off
  WaitIdle
  Slow
  Fast
  Ready
End Enum
Public State As APState
Public StateString As String
Public AddPreparePrompt As Long
Public OverrunTimer As New acTimer
Public AdvanceTimer As New acTimer
  
Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Parameters=Prompt |0-99|\r\nName=Add Prepare\r\nHelp=Signals the operator to prepare the addition tank."
Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  AddPreparePrompt = Param(1)
  With ControlCode
     .AddReadyFromPB = False
  End With
  
'Carry on my wayward son...
  StepOn = True
  State = WaitIdle
  AdvanceTimer.TimeRemaining = 2
  
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)

'Early bind ControlCode for speed and convenience
  Dim ControlCode As ControlCode
  Set ControlCode = ControlObject

  With ControlCode
    Select Case State
    
      Case Off
        StateString = ""
        
      Case WaitIdle
        StateString = "AP: Wait to lose Run Pushbutton "
        If Not .IO_RemoteRun Then
          OverrunTimer.TimeRemaining = .Parameters.StandardLocalPrepTime
          AdvanceTimer.TimeRemaining = 2
          State = Slow
        End If
       
      Case Slow
        StateString = "AP: Wait For Add Ready "
        If Not .IO_RemoteRun Then AdvanceTimer.TimeRemaining = 2
        If .AT.IsWaitReady Then State = Fast
        If .AddReadyFromPB Or AdvanceTimer.Finished Then
          .AddReady = True
          State = Ready
        End If
          
      Case Fast
        StateString = "AP: Wait For Add Ready "
        If .AddReadyFromPB Or AdvanceTimer.Finished Then
          .AddReady = True
          State = Ready
        End If
       
      Case Ready
        StateString = "AP: Add Ready "
        
    End Select
  End With
  
End Sub
Public Sub ACCommand_Cancel()
  State = Off
End Sub
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property
Friend Property Get IsOn() As Boolean
  If (State > Off) Then IsOn = True
End Property
Friend Property Get IsActive() As Boolean
  IsActive = IsOn
End Property
Friend Property Get IsSlow() As Boolean
  If (State = Slow) Then IsSlow = True
End Property
Friend Property Get IsFast() As Boolean
  If (State = Fast) Then IsFast = True
End Property
Friend Property Get IsWaitReady() As Boolean
  If IsSlow Or IsFast Then IsWaitReady = True
End Property
Friend Property Get IsReady() As Boolean
  If (State = Ready) Then IsReady = True
End Property
Public Property Get IsOverrun() As Boolean
  If OverrunTimer.Finished And (State = Fast) Then IsOverrun = True
End Property

#End If
End Class
