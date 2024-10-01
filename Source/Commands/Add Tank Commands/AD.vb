'American & Efird - MX Package
' Version 2024-08-22

Imports Utilities.Translations

<Command("Add Drain", "|0-9|", "", "", ""), _
  TranslateCommand("es", "Añadir desagüe", "|0-9|"), _
  Description("Drain and Rinse the add tank the specified number of times."), _
  TranslateDescription("es", "Escurra y enjuague el tanque de agregar el número especificado de veces."),
  Category("Add Tank Commands"), TranslateCategory("es", "Add Tank Commands")>
Public Class AD : Inherits MarshalByRefObject : Implements ACCommand
  Private Const commandName_ As String = ("AD: ")

  Public Enum EState
    Off
    WaitIdle
    DrainStart
    DrainActive
    Done
  End Enum
  Public State As EState
  Public StateString As String
  Public NumberOfDrains As Integer

  Public Timer As New Timer
  Friend TimerCancel As New Timer
  Friend TimerOverrun As New Timer
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

    'Enable command parameter adjustments, if parameter set
    If controlCode.Parameters.EnableCommandChg = 1 Then
      'Check array bounds just to be on the safe side
      If param.GetUpperBound(0) >= 1 Then NumberOfDrains = param(1)
    End If

  End Sub

  Public Function Start(ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With controlCode

      'Check array bounds just to be on the safe side
      If param.GetUpperBound(0) >= 1 Then NumberOfDrains = param(1)

      'Set default state
      State = EState.WaitIdle
      Timer.Seconds = 10

      ' Set an overrun timer to cancel other tank commands overlapping
      ' example: KP waits for AD waiting for AF...
      TimerCancel.Minutes = 2

      ' TODO - maybe use AddControl.IsOverrun
      TimerOverrun.Minutes = 5

      'This is a background command, but wait until after Sync State to step on
      Return False
    End With
  End Function

  Public Function Run() As Boolean Implements ACCommand.Run
    With controlCode

      Select Case State
        Case EState.Off
          StateString = (" ")

        Case EState.WaitIdle
          StateString = Translate("Wait Idle")
          ' Delay timer has finished - cancel add commands
          If TimerCancel.Finished Then
            If .AC.IsOn Then .AC.Cancel()
            If .AF.IsOn Then .AF.Cancel()
            If .AP.IsOn Then .AP.Cancel()
            If .AT.IsOn Then .AT.Cancel()
          End If
          ' Cancel active add commands - we've jumped to this command
          If .AF.IsOn Then StateString = .AF.Status
          If .AT.IsForeground Then .AT.Cancel()
          If .AT.IsBackground Then StateString = .AT.Status : Timer.Seconds = 1
          If .AC.IsForeground Then .AC.Cancel()
          If .AC.IsBackground Then StateString = ("AC: ") & .AC.Status : Timer.Seconds = 1
          ' Timer Delay
          If Timer.Finished Then
            ' Call the AddControl subroutine
            .AddControl.DrainAuto(NumberOfDrains)

            State = EState.DrainStart
          End If

        Case EState.DrainStart
          StateString = .AddControl.Status
          ' Wait for AddControl to begin the drain sequence before stepping on - could be a sequence issue
          If .AddControl.DrainIsActive Then
            State = EState.DrainActive
            Timer.Seconds = 2

            ' Step On
            Return True
          End If

        Case EState.DrainActive
          StateString = Translate("Active") & Timer.ToString(1)
          If .AddControl.DrainIsActive Then
            StateString = .AddControl.Status
            Timer.Seconds = 5
          End If
          If Timer.Finished Then
            .AddControl.Cancel()
            State = EState.Done
            Timer.Seconds = 5
          End If

        Case EState.Done
          StateString = Translate("Completing") & Timer.ToString(1)
          If Timer.Finished Then State = EState.Off

      End Select
    End With
  End Function

  Public Sub Cancel() Implements ACCommand.Cancel
    NumberOfDrains = 0
    State = EState.Off
    controlCode.AddControl.DrainCancel()
    Timer.Cancel()
    TimerOverrun.Cancel()
  End Sub

  Public ReadOnly Property IsOn As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property

  Public ReadOnly Property IsForeground As Boolean
    Get
      Return (State = EState.WaitIdle) OrElse (State = EState.DrainStart)
    End Get
  End Property

  Public ReadOnly Property IsOverrun() As Boolean
    Get
      Return IsOn AndAlso TimerOverrun.Finished
    End Get
  End Property


  ' AD.vb
#If 0 Then

'===============================================================================================
'AD - 
'===============================================================================================
Option Explicit
Implements ACCommand
Public Enum ADState
  ADOff
  ADWaitForAF
  ADDrainEmpty
  ADFillForMix
  ADMixForTime
  ADCloseMixForTime
  ADDrainRinse
End Enum
Public State As ADState
Public StateString As String
Public NumberOfRinsesDrains As Long
Public ADStateWas As ADState
Public Timer As New acTimer
Public MixTimer As New acTimer
  
Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Parameters= |0-9|\r\nName=Add Drain:\r\nHelp=drain and rinse the add tank for the specified number of times."

'Early bind ControlCode for speed and convenience
  Dim ControlCode As ControlCode
  Set ControlCode = ControlObject
  NumberOfRinsesDrains = Param(1)
  With ControlCode
    .AC.ACCommand_Cancel
    .AT.ACCommand_Cancel
    
    'start timer for 5 minutes: if RF takes more than 5 minutes, cancel the RF to prevent background funciton overlap
    'example: KP waits for AD waiting for AF....
    Timer = 300
    State = ADWaitForAF
    StepOn = True
  End With
  
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
Dim ControlCode As ControlCode: Set ControlCode = ControlObject
  With ControlCode
     Select Case State
     
      Case ADOff
        StateString = ""
        
      Case ADWaitForAF
        If Timer.Finished Then .AF.ACCommand_Cancel
        If .AF.IsOn Then
          StateString = "AD waiting " & .AF.StateString & " " & TimerString(Timer.TimeRemaining)
        Else
          .AF.AddMixing = False
          State = ADDrainEmpty
          Timer = .Parameters.AddTransferToDrainTime
        End If
                 
      Case ADDrainEmpty
        If (.AdditionLevel > 10) Then
          StateString = "AD: Drain Empty " & Pad(.AdditionLevel / 10, "0", 3) & "%"
          Timer = .Parameters.AddTransferToDrainTime
        Else: StateString = "AD: Drain Empty " & TimerString(Timer.TimeRemaining)
        End If
        If Timer.Finished Then State = ADFillForMix
       
      Case ADFillForMix
        StateString = "AD: Fill For Mix " & Pad(.AdditionLevel / 10, "0", 3) & "% / " & Pad(.Parameters.AddRinseFillLevel / 10, "0", 3) & "%"
        If .AdditionLevel > .Parameters.AddRinseFillLevel Then
          State = ADMixForTime
          Timer = .Parameters.AddRinseMixTime
          MixTimer = .Parameters.AddRinseMixPulseTime
        End If
       
      Case ADMixForTime
        StateString = "AD: Mixing " & TimerString(Timer.TimeRemaining)
        If MixTimer.Finished Then
          MixTimer = 1
          State = ADCloseMixForTime
        End If
        If Timer.Finished Then
          Timer = .Parameters.AddTransferToDrainTime
          State = ADDrainRinse
        End If
         
      Case ADCloseMixForTime
        StateString = "AD: Mixing " & TimerString(Timer.TimeRemaining)
        If MixTimer.Finished Then
           MixTimer = .Parameters.AddRinseMixPulseTime
           State = ADMixForTime
        End If
        If Timer.Finished Then
           Timer = .Parameters.AddTransferToDrainTime
           State = ADDrainRinse
        End If
                 
      Case ADDrainRinse
        If (.AdditionLevel > 10) Then
          StateString = "AD: Drain Empty " & Pad(.AdditionLevel / 10, "0", 3) & "%"
          Timer = .Parameters.AddTransferToDrainTime
        Else: StateString = "AD: Drain Empty " & TimerString(Timer.TimeRemaining)
        End If
        If Timer.Finished Then
           NumberOfRinsesDrains = NumberOfRinsesDrains - 1
           If NumberOfRinsesDrains > 0 Then
              State = ADDrainEmpty
              Timer = .Parameters.AddTransferRinseToDrainTime
           Else
              State = ADOff
           End If
        End If
    
    End Select
  End With
End Sub
Friend Sub ACCommand_Cancel()
  State = ADOff
End Sub
Friend Property Get IsOn() As Boolean
  IsOn = (State <> ADOff)
End Property
Friend Property Get IsActive() As Boolean
  IsActive = (State > ADOff) And (State <= ADDrainRinse)
End Property
Friend Property Get IsTransferToDrain() As Boolean
 IsTransferToDrain = (State = ADDrainRinse) Or (State = ADDrainEmpty)
End Property
Friend Property Get IsFillForMix() As Boolean
  If (State = ADFillForMix) Then IsFillForMix = True
End Property
Friend Property Get IsMixingForTime() As Boolean
  If (State = ADMixForTime) Then IsMixingForTime = True
End Property
Friend Property Get IsMixCloseMix() As Boolean
  If (State = ADCloseMixForTime) Then IsMixCloseMix = True
End Property
Friend Property Get IsAddPumping() As Boolean
  If IsMixingForTime Or IsMixCloseMix Then IsAddPumping = True
End Property
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property

#End If





#If 0 Then
  ' TODO Check 


  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===============================================================================================
'AD - Kitchen tank prepare command
'===============================================================================================
Option Explicit
Implements ACCommand
Public Enum ADState
  ADOff
  ADWaitForAF
  ADDrainEmpty
  ADFillForMix
  ADMixForTime
  ADCloseMixForTime
  ADDrainRinse
End Enum
Public State As ADState
Public StateString As String
Public NumberOfRinsesDrains As Long
Public ADStateWas As ADState
Public Timer As New acTimer
Public MixTimer As New acTimer
  
Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Parameters= |0-9|\r\nName=Add Drain:\r\nHelp=drain and rinse the add tank for the specified number of times."

'Early bind ControlCode for speed and convenience
  Dim ControlCode As ControlCode
  Set ControlCode = ControlObject
  NumberOfRinsesDrains = Param(1)
  With ControlCode
    .AC.ACCommand_Cancel
    .AT.ACCommand_Cancel
    
    'start timer for 5 minutes: if RF takes more than 5 minutes, cancel the RF to prevent background funciton overlap
    'example: KP waits for AD waiting for AF....
    Timer = 300
    State = ADWaitForAF
    StepOn = True
  End With
  
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
Dim ControlCode As ControlCode: Set ControlCode = ControlObject
  With ControlCode
     Select Case State
     
      Case ADOff
        StateString = ""
        
      Case ADWaitForAF
        If Timer.Finished Then .AF.ACCommand_Cancel
        If .AF.IsOn Then
          StateString = "AD waiting " & .AF.StateString & " " & TimerString(Timer.TimeRemaining)
        Else
          .AF.AddMixing = False
          State = ADDrainEmpty
          Timer = .Parameters.AddTransferToDrainTime
        End If
                 
      Case ADDrainEmpty
        If (.AdditionLevel > 10) Then
          StateString = "AD: Drain Empty " & Pad(.AdditionLevel / 10, "0", 3) & "%"
          Timer = .Parameters.AddTransferToDrainTime
        Else: StateString = "AD: Drain Empty " & TimerString(Timer.TimeRemaining)
        End If
        If Timer.Finished Then State = ADFillForMix
       
      Case ADFillForMix
        StateString = "AD: Fill For Mix " & Pad(.AdditionLevel / 10, "0", 3) & "% / " & Pad(.Parameters.AddRinseFillLevel / 10, "0", 3) & "%"
        If .AdditionLevel > .Parameters.AddRinseFillLevel Then
          State = ADMixForTime
          Timer = .Parameters.AddRinseMixTime
          MixTimer = .Parameters.AddRinseMixPulseTime
        End If
       
      Case ADMixForTime
        StateString = "AD: Mixing " & TimerString(Timer.TimeRemaining)
        If MixTimer.Finished Then
          MixTimer = 1
          State = ADCloseMixForTime
        End If
        If Timer.Finished Then
          Timer = .Parameters.AddTransferToDrainTime
          State = ADDrainRinse
        End If
         
      Case ADCloseMixForTime
        StateString = "AD: Mixing " & TimerString(Timer.TimeRemaining)
        If MixTimer.Finished Then
           MixTimer = .Parameters.AddRinseMixPulseTime
           State = ADMixForTime
        End If
        If Timer.Finished Then
           Timer = .Parameters.AddTransferToDrainTime
           State = ADDrainRinse
        End If
                 
      Case ADDrainRinse
        If (.AdditionLevel > 10) Then
          StateString = "AD: Drain Empty " & Pad(.AdditionLevel / 10, "0", 3) & "%"
          Timer = .Parameters.AddTransferToDrainTime
        Else: StateString = "AD: Drain Empty " & TimerString(Timer.TimeRemaining)
        End If
        If Timer.Finished Then
           NumberOfRinsesDrains = NumberOfRinsesDrains - 1
           If NumberOfRinsesDrains > 0 Then
              State = ADDrainEmpty
              Timer = .Parameters.AddTransferRinseToDrainTime
           Else
              State = ADOff
           End If
        End If
    
    End Select
  End With
End Sub
Friend Sub ACCommand_Cancel()
  State = ADOff
End Sub
Friend Property Get IsOn() As Boolean
  IsOn = (State <> ADOff)
End Property
Friend Property Get IsActive() As Boolean
  IsActive = (State > ADOff) And (State <= ADDrainRinse)
End Property
Friend Property Get IsTransferToDrain() As Boolean
 IsTransferToDrain = (State = ADDrainRinse) Or (State = ADDrainEmpty)
End Property
Friend Property Get IsFillForMix() As Boolean
  If (State = ADFillForMix) Then IsFillForMix = True
End Property
Friend Property Get IsMixingForTime() As Boolean
  If (State = ADMixForTime) Then IsMixingForTime = True
End Property
Friend Property Get IsMixCloseMix() As Boolean
  If (State = ADCloseMixForTime) Then IsMixCloseMix = True
End Property
Friend Property Get IsAddPumping() As Boolean
  If IsMixingForTime Or IsMixCloseMix Then IsAddPumping = True
End Property
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property


#End If
End Class
