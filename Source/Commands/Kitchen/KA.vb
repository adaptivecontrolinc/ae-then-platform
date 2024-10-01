'American & Efird - Package With Add & Reserve
' Version 2024-08-27

Imports Utilities.Translations

<Command("Kitchen Add", "Destination:|AR|", " ", "", "10"), 
Description("Transfers the kitchen tank to the specified destination."), 
TranslateDescription("es", "Transfiere el depósito de cocina al destino especificado."), 
Category("Kitchen Functions"), TranslateCategory("es", "Funciones de cocina")> 
Public Class KA : Inherits MarshalByRefObject : Implements ACCommand
  Private Const commandName_ As String = ("KA: ")

  'Command States
  Public Enum EState
    Off
    WaitReady
    Pause
    WaitTransfer
    Transfer1
    Transfer1Delay
    Rinse
    Transfer2
    Transfer2Delay
    RinseToDrainPause
    RinseToDrain
    TransferToDrain
    TransferToDrainDelay
    Complete
  End Enum
  Public State As EState
  Public StateRestart As EState
  Public Status As String
  Public Destination As EKitchenDestination
  Public NumberOfRinses As Integer

  Public Timer As New Timer
  Friend ReadOnly TimerLevelDisregard As New Timer
  Public ReadOnly Property TimeLevelDisregard As Integer
    Get
      Return TimerLevelDisregard.Seconds
    End Get
  End Property
  Friend Property TimerOverrun As New Timer
  Public ReadOnly Property TimeOverrun As Integer
    Get
      Return TimerOverrun.Seconds
    End Get
  End Property
  Private timerWaitReady_ As New TimerUp
  Public ReadOnly Property TimeWaitReady As Integer
    Get
      If (Not IsOn) OrElse (State > EState.WaitReady) Then
        timerWaitReady_.Stop()
      End If
      Return timerWaitReady_.Seconds
    End Get
  End Property

  'Uses these for flashy lamp stuff
  Private flashSlow As Boolean
  Private flashFast As Boolean

  'Keep a local reference to the control code object for convenience
  Private ReadOnly ControlCode As ControlCode
  Sub New(ByVal controlCode As ControlCode)
    Me.ControlCode = controlCode
    Me.ControlCode.Commands.Add(Me)
  End Sub

  Sub ParametersChanged(ByVal ParamArray param() As Integer) Implements ACCommand.ParametersChanged
    'If the command is not running then do nothing the changes will be picked up when the command starts
    If Not IsOn Then Exit Sub

    'Enable command parameter adjustments, if parameter set
    If ControlCode.Parameters.EnableCommandChg = 1 Then
      Start(param)
    End If
  End Sub

  Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With ControlCode

      ' Cancel Commands
      '.KA.Cancel()
      .CK.Cancel()
      .DR.Cancel() : .FI.Cancel() : .HD.Cancel()
      .RH.Cancel() : .RI.Cancel() : .TM.Cancel() ': .TU.Cancel()
      .LD.Cancel() : .PH.Cancel() : .SA.Cancel() : .UL.Cancel()
      .WT.Cancel()

      ' Cancel PressureOn command if active during Atmospheric Commands
      .PR.Cancel()

      ' Check - below was KA parameters for vb6
      '   R: If (Destination = 82) Then .ReserveReady = False
      '   A: If (Destination = 65) Then .AddReady = False



      ' Command parameters
      If param.GetUpperBound(0) >= 1 Then Destination = CType(param(1), EKitchenDestination)
      If (Destination = EKitchenDestination.Reserve) Then .ReserveReady = False
      If (Destination = EKitchenDestination.Add) Then .AddReady = False

      ' Initial State
      State = EState.WaitReady
      Timer.Seconds = 2
      TimerOverrun.Minutes = (.Parameters.StandardTimeAddPrepare) + 5

      'This is a foreground command so do not step on
      Return False
    End With
  End Function

  Function Run() As Boolean Implements ACCommand.Run
    With ControlCode
      ' Use control code flashers so everything stays in sync
      flashSlow = ControlCode.FlashSlow
      flashFast = ControlCode.FlashFast

      ' Pause the state machine under these conditions
      Dim pauseCommand As Boolean = .Parent.IsPaused OrElse .EStop
      Dim pauseTransfer As Boolean = .Parent.IsPaused OrElse .EStop

      Static StartState As EState
      Do
        ' Remember state and loop until state does not change
        ' This makes sure we go through all the state changes before proceeding and setting IO
        StartState = State

        ' If pausing command, remember where we were
        If pauseCommand AndAlso (State > EState.WaitTransfer) AndAlso (State < EState.RinseToDrainPause) Then
          ' Could get here while in WaitSafe (Estop Pressed, etc), if setting to WaitSafe, then will lock the command
          If (State > EState.WaitTransfer) Then StateRestart = State
          State = EState.Pause
          Timer.Pause()
        End If

        ' If pausing the transfer (not safe or pump not running), remember where we were
        If (pauseTransfer) AndAlso (State > EState.WaitTransfer) AndAlso (State < EState.RinseToDrainPause) Then
          ' Only update with a further state (prevent locking command)
          If (State > EState.WaitTransfer) Then StateRestart = State
          State = EState.WaitTransfer
          Timer.Pause()
        End If


        '********************************************************************************************
        '******   STATE LOGIC
        '********************************************************************************************
        Select Case State
          Case EState.Off
            Status = (" ")

          Case EState.WaitReady
            Status = Translate("Wait Ready") & Timer.ToString(1)
            If .KP.KP1.IsOn Then
              Status = .KP.KP1.DrugroomDisplay
              ' Must wait for KP to finish mixing before transferring tank to ET
              If Not .KP.KP1.IsReady Then Timer.Seconds = 5
            End If
            ' Restart Drugroom Preview
            If Not .Parent.IsPaused AndAlso timerWaitReady_.IsPaused Then timerWaitReady_.Restart()
            ' Reset Ready Delay Timer
            If Not .Tank1Ready Then Timer.Seconds = 5
            If .Tank1Ready AndAlso Timer.Finished Then
              State = EState.WaitTransfer
              timerWaitReady_.Pause()
              .KP.Cancel()
              Timer.Seconds = 2
            End If
            ' Update Restart State
            StateRestart = EState.WaitReady

          Case EState.Pause
            Status = Translate("Paused")
            If .EStop Then Status = Translate("EStop") : Timer.Seconds = 2
            If .Parent.IsPaused Then Status = Translate("Paused") : Timer.Seconds = 2
            If Timer.Finished Then
              ' Begin or restart
              If StateRestart > EState.Rinse Then
                State = StateRestart
                Timer.Restart()
              Else
                State = EState.WaitTransfer
                Timer.Seconds = 2
              End If
            End If

          Case EState.WaitTransfer
            Status = Translate("Wait to start") & Timer.ToString(1)
            If .EStop Then Status = Translate("EStop") : Timer.Seconds = 2
            If .Parent.IsPaused Then Status = Translate("Paused") : Timer.Seconds = 2
            If Destination = EKitchenDestination.Add Then
              If (.AddLevel > .Parameters.AdditionMaxTransferLevel) Then Status = Translate("Add Level High") : Timer.Seconds = 2
            End If
            If Timer.Finished Then
              ' Begin or restart
              If StateRestart > EState.Rinse Then
                State = StateRestart
                Timer.Restart()
              Else
                State = EState.Transfer1
                Timer.Seconds = .Parameters.Tank1TimeBeforeRinse
              End If
            End If
            ' Update Restart State
            StateRestart = EState.WaitTransfer


            '********************************************************************************************
            '******   TRANSFER TANK TO DESTINATION
            '********************************************************************************************
          Case EState.Transfer1
            Status = Translate("Transferring")
            If (.Tank1Level > 10) Then
              Status = Status & (" ") & (.Tank1Level / 10).ToString("#0.0") & "%"
              Timer.Seconds = MinMax(.Parameters.Tank1TimeBeforeRinse, 10, 600)
            Else
              Status = Status & Timer.ToString(1)
            End If
            If (.Parameters.Tank1LevelDisregardTimeTransfer > 0) Then
              ' Disregard tank level input due to signal bouncing - use delay timer
              Status = Translate("Transfer") & TimerLevelDisregard.ToString(1)
              If TimerLevelDisregard.Finished Then
                State = EState.Transfer1Delay
                Timer.Seconds = MinMax(.Parameters.Tank1TimeBeforeRinse, 10, 600)
              End If
            Else
              ' Use level timer
              If Timer.Finished Then
                State = EState.Transfer1Delay
                Timer.Seconds = MinMax(.Parameters.Tank1TimeBeforeRinse, 10, 600)
              End If
            End If
            ' Update Restart State
            StateRestart = EState.Transfer1

          Case EState.Transfer1Delay
            ' Once level drops below 1.0% for 2 seconds, disregard level entireley 
            '   UltraSonic kitchen tank level transmitters bounce around when tank is empty, this is a workaround
            Status = Translate("Transferring") & Timer.ToString(1)
            If Timer.Finished Then
              .Tank1Ready = False
              NumberOfRinses = .Parameters.Tank1RinseNumber
              ' Kitchen Rinse command active
              If .KR.IsRinseMachineOff AndAlso .KR.IsRinseDrainOff Then
                .KR.Cancel()
                Cancel()
              ElseIf .KR.IsRinseMachineOff Then
                State = EState.Transfer2
                Timer.Seconds = MinMax(.Parameters.Tank1RinseToDrainTime, 10, 300)
              Else
                State = EState.Rinse
                Timer.Seconds = MinMax(.Parameters.Tank1RinseTime, 1, 60)
              End If
            End If
            ' Update Restart State
            StateRestart = EState.Transfer1

          Case EState.Rinse
            Status = Translate("Rinsing") & Timer.ToString(1)
            If Timer.Finished Then
              NumberOfRinses -= 1
              State = EState.Transfer2
              Timer.Seconds = .Parameters.Tank1TimeAfterRinse
              TimerLevelDisregard.Seconds = MinMax(.Parameters.Tank1LevelDisregardTimeTransfer, 10, 100)
            End If
            ' Update Restart State
            StateRestart = EState.Transfer1

          Case EState.Transfer2
            Status = Translate("Transferring")
            If (.Tank1Level > 10) Then
              Status = Status & (" ") & (.Tank1Level / 10).ToString("#0.0") & "%"
              Timer.Seconds = .Parameters.Tank1TimeAfterRinse
            Else
              Status = Status & Timer.ToString(1)
            End If
            If (.Parameters.Tank1LevelDisregardTimeTransfer > 0) Then
              ' Disregard tank level input due to signal bouncing - use delay timer
              Status = Translate("Transfer") & TimerLevelDisregard.ToString(1)
              If TimerLevelDisregard.Finished Then
                State = EState.Transfer2Delay
                Timer.Seconds = .Parameters.Tank1TimeAfterRinse
              End If
            Else
              ' Use level timer
              If Timer.Finished Then
                State = EState.Transfer2Delay
                Timer.Seconds = .Parameters.Tank1TimeAfterRinse
              End If
            End If
            ' Update Restart State
            StateRestart = EState.Transfer2

          Case EState.Transfer2Delay
            ' Once level drops below 1.0% for 2 seconds, disregard level entireley 
            '   UltraSonic kitchen tank level transmitters bounce around when tank is empty, this is a workaround
            Status = Translate("Transferring") & Timer.ToString(1)
            If Timer.Finished Then
              If (NumberOfRinses > 0) Then
                State = EState.Rinse
                Timer.Seconds = MinMax(.Parameters.Tank1RinseTime, 1, 60)
              Else
                ' Finished transferring, rinse to drain
                State = EState.RinseToDrain
                Timer.Seconds = .Parameters.Tank1RinseToDrainTime
              End If
            End If
            ' Update Restart State
            StateRestart = EState.Transfer2


            '********************************************************************************************
            '******   TRANSFER TANK TO DRAIN
            '********************************************************************************************
          Case EState.RinseToDrainPause
            Status = Translate("Paused")
            If .EStop Then Status = Translate("EStop") : Timer.Seconds = 2
            If .Parent.IsPaused Then Status = Translate("Paused") : Timer.Seconds = 2
            If Timer.Finished Then
              ' Begin or restart
              If StateRestart > EState.RinseToDrain Then
                State = StateRestart
                Timer.Restart()
              Else
                ' Restart Rinsing
                State = EState.RinseToDrain
                Timer.Seconds = .Parameters.Tank1RinseToDrainTime
              End If
            End If

          Case EState.RinseToDrain
            Status = Translate("Rinsing To Drain") & Timer.ToString(1)
            ' Use Rinse Time
            If Timer.Finished Then
              State = EState.TransferToDrain
                Timer.Seconds = MinMax(.Parameters.Tank1DrainTime, 10, 600)
              TimerLevelDisregard.Seconds = MinMax(.Parameters.Tank1LevelDisregardTimeTransfer, 10, 100)
            End If
            ' Pause Transfer to Drain
            If .Parent.IsPaused OrElse .EStop Then
              ' Set pause state (even if it's the next state, to prevent re-rinsing)
              StateRestart = State
              State = EState.RinseToDrainPause
              Timer.Pause()
            End If

          Case EState.TransferToDrain
            Status = Translate("Transferring To Drain")
            If (.Tank1Level > 10) Then
              Status = Status & (" ") & (.Tank1Level / 10).ToString("#0.0") & "%"
              Timer.Seconds = .Parameters.Tank1DrainTime
            Else : Status = Status & Timer.ToString(1)
            End If
            If (.Parameters.Tank1LevelDisregardTimeTransfer > 0) Then
              ' Disregard tank level input due to signal bouncing - use delay timer
              Status = Translate("Transfer") & (" ") & TimerLevelDisregard.ToString
              If TimerLevelDisregard.Finished Then
                State = EState.TransferToDrainDelay
                Timer.Seconds = 5
              End If
            Else
              ' Use level timer
              If Timer.Finished Then
                State = EState.TransferToDrainDelay
                Timer.Seconds = 5
              End If
            End If
            ' Pause Transfer to Drain
            If .Parent.IsPaused OrElse .EStop Then
              ' Set pause state (even if it's the next state, to prevent re-rinsing)
              StateRestart = State
              State = EState.RinseToDrainPause
              Timer.Seconds = MinMax(.Parameters.Tank1DrainTime, 10, 600)
              Timer.Pause()
            End If

          Case EState.TransferToDrainDelay
            Status = Translate("Transferring To Drain") & Timer.ToString(1)
            If Timer.Finished Then
              If Destination = EKitchenDestination.Add Then .AddReady = True
              If Destination = EKitchenDestination.Reserve Then .ReserveReady = True
              State = EState.Complete
              Timer.Seconds = 5
            End If

          Case EState.Complete
            Status = Translate("Completing") & Timer.ToString(1)
            If Timer.Finished Then
              ' Cancel kitchen commands 
              .KP.Cancel() : .KR.Cancel()
              Destination = 0
              .Tank1Ready = False
              State = EState.Off
            End If

        End Select

      Loop Until (StartState = State)

    End With
  End Function

  Public Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
    Timer.Cancel()
    TimerOverrun.Cancel()
    TimerLevelDisregard.Cancel()
    timerWaitReady_.Stop()
  End Sub

#Region " Public Properties"

  ReadOnly Property IsOn As Boolean Implements ACCommand.IsOn
    Get
      IsOn = State <> EState.Off
    End Get
  End Property
  ReadOnly Property IsOverrun As Boolean
    Get
      Return IsOn AndAlso TimerOverrun.Finished
    End Get
  End Property
  ReadOnly Property IsWaitReady As Boolean
    Get
      Return (State = EState.WaitReady)
    End Get
  End Property
  ReadOnly Property IsActive As Boolean
    Get
      Return (State > EState.Pause) AndAlso (State < EState.Complete)
    End Get
  End Property
  ReadOnly Property IsForeground As Boolean
    Get
      Return (State >= EState.WaitReady) AndAlso (State <= EState.Transfer2Delay)
    End Get
  End Property
  Public ReadOnly Property IsBackgroundRinsing As Boolean
    Get
      Return (State >= EState.RinseToDrainPause) AndAlso (State <= EState.Complete)
    End Get
  End Property

#End Region

#Region " I/O PROPERTIES "

  ReadOnly Property IoTank1Transfer As Boolean
    Get
      Return ((State >= EState.Transfer1) AndAlso (State <= EState.Transfer2Delay)) OrElse
             ((State >= EState.RinseToDrain) AndAlso (State <= EState.TransferToDrainDelay))
    End Get
  End Property

  ReadOnly Property IoTank1Rinse As Boolean
    Get
      Return (State = EState.Rinse) OrElse (State = EState.RinseToDrain)
    End Get
  End Property

  ReadOnly Property IoTank1Mixer As Boolean
    Get
      Return False '
    End Get
  End Property

  ReadOnly Property IoTank1ToAdd As Boolean
    Get
      Return Destination = EKitchenDestination.Add AndAlso _
             ((State >= EState.Transfer1) AndAlso (State <= EState.Transfer2Delay))
    End Get
  End Property

  ReadOnly Property IoTank1ToReserve As Boolean
    Get
      Return Destination = EKitchenDestination.Reserve AndAlso _
             ((State >= EState.Transfer1) AndAlso (State <= EState.Transfer2Delay))
    End Get
  End Property

  ReadOnly Property IoTank1ToDrain As Boolean
    Get
      Return (State >= EState.RinseToDrain) AndAlso (State <= EState.TransferToDrainDelay)
    End Get
  End Property

#End Region



#If 0 Then
' TODO

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
'===============================================================================================
'KA - Kitchen tank prepare command
'===============================================================================================
Option Explicit
Implements ACCommand
Public Enum KAState
  KAOff
  KAOn
  KAPause
  KAWaitReady
  KAInterlock
  KATransfer1
  KARinse
  KATransfer2
  KARinseToDrain
  KATransferToDrain
End Enum
Public StateString As String
Public State As KAState, Tank1State As KAState
Public KAStateWas As KAState
Private NumberOfRinses As Long, NumberOfFlushes As Long
Public Timer As New acTimer
Public OverrunTimer As New acTimer
Public Destination As Variant
Public WaitReadyTimer As New acTimerUp

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Name=Kitchen Add\r\nParameters=Dest. |AR| \r\nHelp=Transfers tank 1 to the addition tank (A) or the reserve tank (R)"
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode
    Destination = Param(1)
    If (Destination = 82) Then .ReserveReady = False
    If (Destination = 65) Then .AddReady = False
    State = KAWaitReady
    WaitReadyTimer.Start
    OverrunTimer = 360
  End With
  StepOn = True
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
Dim ControlCode As ControlCode: Set ControlCode = ControlObject
  With ControlCode
  
  Tank1State = State
  
  'pause the transfer and remember where we were
  If .IO_EStop_PB Then
    If State > KAPause Then
      KAStateWas = State
      State = KAPause
      Timer.Pause
    End If
  End If

  'Tank 1 Control
  Select Case State
    
    Case KAOff
      StateString = ""
      
    Case KAWaitReady
      If .KP.IsOn And Not .KP.KP1.IsReady Then
        StateString = "KA: " & .KP.KP1.StateString
        Timer = 1
      Else
        StateString = "KA: Wait Ready " & TimerString(Timer.TimeRemaining)
      End If
      OverrunTimer = 360
      If Timer.Finished Then
        WaitReadyTimer.Pause
        'make sure level not to high in add tank
        If (Destination = 65) Then
          State = KAInterlock
        Else
          State = KATransfer1
         .KP.ACCommand_Cancel
        End If
      End If
      
    Case KAInterlock
      StateString = "KA: Add Level To High to Transfer "
      If (.AdditionLevel <= .Parameters.AdditionMaxTransferLevel) Then
        State = KATransfer1
        .KP.ACCommand_Cancel
      End If
      
    Case KAPause
      StateString = "KA: Paused "
      If Not (.IO_EStop_PB) Then
        State = KAStateWas
        Timer.Restart
      End If
   
   Case KATransfer1
      If (.Tank1Level > 10) Then
        StateString = "KA: Tank 1 Transfer " & Pad(.Tank1Level / 10, "0.0", 3) & "%"
        Timer = .Parameters.Tank1TimeBeforeRinse
      Else: StateString = "KA: Tank 1 Transfer " & TimerString(Timer.TimeRemaining)
      End If
      If Timer.Finished Then
        NumberOfRinses = .Parameters.DrugroomRinses
        State = KARinse
        Timer = .Parameters.Tank1RinseTime
      End If
    
    Case KARinse
      StateString = "KA: Tank 1 Rinse " & Pad(.Tank1Level / 10, "0.0", 3) & " / " & Pad(.Parameters.Tank1RinseLevel / 10, "0.0", 3) & " %"
      If (.Tank1Level > .Parameters.Tank1RinseLevel) Then
        State = KATransfer2
        Timer = .Parameters.Tank1TimeAfterRinse
        NumberOfRinses = NumberOfRinses - 1
      End If
    
    Case KATransfer2
      If (.Tank1Level > 10) Then
        StateString = "KA: Tank 1 Transfer " & Pad(.Tank1Level / 10, "0.0", 3) & "%"
        Timer = .Parameters.Tank1TimeAfterRinse
      Else: StateString = "KA: Tank 1 Transfer " & TimerString(Timer.TimeRemaining)
      End If
      If Timer.Finished Then
        If NumberOfRinses > 0 Then
          State = KARinse
          Timer = .Parameters.Tank1RinseTime
        Else
          State = KARinseToDrain
          Timer = .Parameters.Tank1RinseToDrainTime
        End If
      End If
    
    Case KARinseToDrain
      StateString = "KA: Tank 1 Rinse to Drain " & TimerString(Timer.TimeRemaining)
      If Timer.Finished Or (.Tank1Level >= 500) Then
        State = KATransferToDrain
        Timer = .Parameters.Tank1DrainTime
      End If
    
    Case KATransferToDrain
      If (.Tank1Level > 10) Then
        StateString = "KA: Tank 1 Transfer to Drain " & Pad(.Tank1Level / 10, "0.0", 3) & "%"
        Timer = .Parameters.Tank1DrainTime
      Else: StateString = "KA: Tank 1 Transfer to Drain " & TimerString(Timer.TimeRemaining)
      End If
      If Timer.Finished Then
        StepOn = True
        If (Destination = 65) Then .AddReady = True
        If (Destination = 82) Then .ReserveReady = True
        Destination = 0
           
        State = KAOff
        .Tank1Ready = False
        .KP.ACCommand_Cancel
      End If
 
  End Select
  End With
End Sub
Friend Sub ACCommand_Cancel()
  State = KAOff
  Destination = 0
End Sub
Friend Property Get IsOn() As Boolean
  IsOn = (State <> KAOff)
End Property
Friend Property Get IsActive() As Boolean
  IsActive = (State > KAOff) And (State <= KATransferToDrain)
End Property
Friend Property Get IsInterlocked() As Boolean
 IsInterlocked = (State = KAInterlock)
End Property
Friend Property Get IsTransfer() As Boolean
  IsTransfer = (State = KATransfer1) Or (State = KATransfer2) Or (State = KARinseToDrain) Or (State = KATransferToDrain)
End Property
Friend Property Get IsTransferToAddition() As Boolean
 IsTransferToAddition = ((State = KATransfer1) Or (State = KARinse) Or _
                       (State = KATransfer2)) And (Destination = 65)
End Property
Friend Property Get IsTransferToReserve() As Boolean
 IsTransferToReserve = ((State = KATransfer1) Or (State = KARinse) Or _
                       (State = KATransfer2)) And (Destination = 82)
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
Friend Property Get IsWaitReady() As Boolean
  IsWaitReady = (State = KAWaitReady)
End Property
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property
Friend Property Get IsOverrun() As Boolean
  If (State >= KAWaitReady) And OverrunTimer.Finished Then IsOverrun = True
End Property



#End If
End Class
