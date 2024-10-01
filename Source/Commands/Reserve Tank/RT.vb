'American & Efird - Mt. Holly Then Platform
' Version 2024-09-17

Imports Utilities.Translations

<Command("Reserve Transfer", "", "", "", "('StandardTimeReserveTransfer)=5"),
TranslateCommand("es", "Tanque de reserva transferencia", ""),
Description("Transfer the Reserve tank to the machine."),
TranslateDescription("es", ""),
Category("Reserve Tank Commands"), TranslateCategory("es", "Reserve Tank Commands")>
Public Class RT
  Inherits MarshalByRefObject
  Implements ACCommand
  Private Const commandName_ As String = ("RT: ")

  'Command states
  Public Enum EState
    Off
    WaitReady
    WaitForSafe

    Pause
    Restart

    TransferStart
    TransferGravity
    TransferEmpty1
    RinseToTransfer
    TransferEmpty2
    ' Step on at this point

    RinseToDrainPause
    RinseToDrain
    TransferToDrain
    Complete
  End Enum
  Public Property State As EState
  Public Property StateRestart As EState
  Public Property Status As String
  Public Property NumberOfRinses As Integer
  Public Property LevelFill As Integer

  ' Timers
  Public Property Timer As New Timer
  Friend Property TimerMix As New Timer
  Public ReadOnly Property TimeMix As Integer
    Get
      Return TimerMix.Seconds
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
      Start(param)
    End If
  End Sub

  Public Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With controlCode

      ' Cancel Commands
      .AC.Cancel() : .AT.Cancel()
      .KA.Cancel() : .WK.Cancel()
      .DR.Cancel() : .FI.Cancel() : .HC.Cancel() : .HD.Cancel() : .PR.Cancel()
      .RI.Cancel() : .RC.Cancel() : .RH.Cancel() : .TM.Cancel()
      .LD.Cancel() : .PH.Cancel() : .SA.Cancel() : .UL.Cancel()
      ' .RT.Cancel()

      .AirpadOn = False
      .TemperatureControl.Cancel()

      'If Reserve Not Prepared and RP not on, then start it
      If Not .ReserveReady Then
        '    If Not .RP.IsOn Then .RP.Start(0, 0)
      End If

      ' Cancel RD background command
      If .RD.IsOn AndAlso Not .RD.IsCancelling Then .RD.Cancel()

      'Set the default state
      State = EState.WaitReady
      StateRestart = EState.Off

      ' Update Drugroom Preview
      timerWaitReady_.Start()

      'Set default time
      TimerOverrun.Minutes = (.Parameters.StandardTimeReserveTransfer + 5)

      'This is a foreground command so do not step on
      Return False
    End With
  End Function

  Public Function Run() As Boolean Implements ACCommand.Run
    With controlCode
      'In this command we will remember the state of each state machine when we start
      '  we can then loop through the state machine until all state changes have completed 
      '  this prevents dwelling on an intermediate state between control code runs which can occasionally cause IO flicker

      'Pause the state machine under these conditions
      Dim safe As Boolean = (.TempSafe OrElse .HD.HDCompleted) AndAlso .MachineClosed
      Dim pauseCommand As Boolean = .Parent.IsPaused OrElse .EStop

      Static StartState As EState
      Do
        'Remember state and loop until state does not change
        'This makes sure we go through all the state changes before proceeding and setting IO
        StartState = State

        Select Case State

          Case EState.Off
            Status = (" ")
            StateRestart = EState.Off

          Case EState.WaitReady
            Status = Translate("Wait Ready") & Timer.ToString(1)
            If .RP.IsWaitReady Then Status = .RP.StateString : Timer.Seconds = 5
            If .RF.IsForeground Then Status = .RF.Status : Timer.Seconds = 5
            If .RD.IsOn Then Status = .RD.Status : Timer.Seconds = 5
            If Not .ReserveReady Then Timer.Seconds = 1
            ' Kitchen Tank
            If (.KA.IsOn AndAlso .KA.Destination = EKitchenDestination.Reserve) Then Status = .KA.Status : Timer.Seconds = 5

            ' TODO - Improve this
            If .KP.KP1.IsOn Then Status = .KP.KP1.DrugroomDisplay : Timer.Seconds = 5
            'If (.KP.KP1.IsOn AndAlso (Not .KP.KP1.IsBackground) AndAlso (.KP.KP1.Param_Destination = EKitchenDestination.Reserve)) Then StateString = .KP.KP1.DrugroomDisplay : Timer.Seconds = 5
            If (.LA.KP1.IsOn AndAlso (Not .LA.KP1.IsBackground) AndAlso (.LA.KP1.Param_Destination = EKitchenDestination.Reserve)) Then Status = .LA.KP1.DrugroomDisplay : Timer.Seconds = 5
            ' Restart Drugroom Preview
            If (Not .Parent.IsPaused) AndAlso timerWaitReady_.IsPaused Then timerWaitReady_.Restart()
            ' Reserve Tank Ready Delay finished
            If Timer.Finished Then
              State = EState.WaitForSafe
              timerWaitReady_.Pause()
            End If
            ' Update Restart State
            StateRestart = EState.WaitReady

          Case EState.WaitForSafe
            Status = Translate("Wait for Safe") & Timer.ToString(1)
            ' Machine Not Safe To Transfer
            If Not safe Then
              If Not (.TempSafe OrElse .HD.HDCompleted) Then Status = Translate("Temp Not Safe")
              If Not .PressSafe Then Status = Translate("Press Not Safe")
              If Not .MachineClosed Then Status = Translate("Machine Not Closed")
              Timer.Seconds = 1
            End If
            ' Machine Paused
            If pauseCommand Then
              If .Parent.IsPaused Then Status = Translate("Paused")
              If .EStop Then Status = Translate("Paused") & (", ESTOP")
              Timer.Seconds = 1
            End If
            ' Timer finished
            If Timer.Finished Then
              ' Cancel all temperature commands
              .CO.Cancel() : .HC.Cancel() : .HE.Cancel() : .WT.Cancel()
              ' Start Transferring
              State = EState.TransferStart
            End If
            ' Update Restart State
            StateRestart = EState.WaitForSafe

          Case EState.Pause
            Status = Translate("Paused")
            ' Machine Not Safe To Transfer
            If Not safe Then
              If Not (.TempSafe OrElse .HD.HDCompleted) Then Status = Translate("Temp Not Safe")
              If Not .PressSafe Then Status = Translate("Press Not Safe")
              If Not .MachineClosed Then Status = Translate("Machine Not Closed")
              Timer.Seconds = 1
            End If
            ' Machine Paused
            If pauseCommand Then
              If .Parent.IsPaused Then Status = Translate("Paused")
              If .EStop Then Status = Translate("Paused") & (", ESTOP")
              Timer.Seconds = 1
            End If
            If Timer.Finished Then
              State = EState.Restart
              Timer.Seconds = 5
            End If

          Case EState.Restart
            Status = Translate("Restarting") & Timer.ToString(1)
            If Timer.Finished Then
              State = StateRestart
              Timer.Seconds = 5
            End If


            '********************************************************************************************
            '******   TRANSFERRING RESERVE TANK CONTENTS
            '********************************************************************************************
          Case EState.TransferStart
            Status = Translate("Starting") & Timer.ToString(1)
            If Timer.Finished Then
              ' Cancel RT background commands
              .ReserveReady = False
              If .RD.IsOn Then .RD.Cancel()
              If .RP.IsOn Then .RP.Cancel()
              If .RF.IsOn Then .RF.Cancel()

              State = EState.TransferGravity
              Timer.Seconds = MinMax(.Parameters.ReserveTimeGravityTransfer, 5, 300)
            End If
            ' Lost Safe conditions
            If Not safe Then State = EState.WaitForSafe

          Case EState.TransferGravity
            Status = Translate("Transferring") & (" ") & (.ReserveLevel / 10).ToString("#0.0") & "%" & Timer.ToString(1)
            'Complete Reserve Transfer
            If Timer.Finished Then
              .PumpControl.StartAuto()
              State = EState.TransferEmpty1
            End If
            ' Update Restart State
            StateRestart = EState.TransferStart
            ' Lost Safe conditions
            If Not safe Then State = EState.WaitForSafe
            ' Pause conditions
            If pauseCommand Then State = EState.Pause

          Case EState.TransferEmpty1
            If .ReserveLevel > 10 Then
              Status = Translate("Transferring") & (" ") & (.ReserveLevel / 10).ToString("#0.0") & "%"
              Timer.Seconds = .Parameters.AddTransferTimeBeforeRinse
            Else
              Status = Translate("Transferring") & Timer.ToString(1)
            End If
            If Timer.Finished Then
              If .RF.IsOn AndAlso (.RF.FillType = EFillType.Vessel) Then
                ' Used Runback - no more rinsing to machine - rinse to drain
                State = EState.RinseToDrain
                Timer.Seconds = MinMax(.Parameters.ReserveRinseToDrainTime, 5, 120)
                NumberOfRinses = .Parameters.ReserveRinseToDrainNumber

                'Tell the control system to step on
                Return True
              Else
                State = EState.RinseToTransfer
                Timer.Seconds = MinMax(.Parameters.ReserveRinseTime, 5, 60)
              End If
            End If
            ' Update Restart State
            StateRestart = EState.TransferStart
            ' Lost Safe conditions
            If Not safe Then State = EState.WaitForSafe
            ' Pause conditions
            If pauseCommand Then State = EState.Pause

          Case EState.RinseToTransfer
            If Timer.Finished Then
              State = EState.TransferEmpty2
              Timer.Seconds = MinMax(.Parameters.ReserveTimeAfterRinse, 5, 300)
            End If
            ' Update Restart State
            StateRestart = EState.TransferStart
            ' Lost Safe conditions
            If Not safe Then State = EState.WaitForSafe
            ' Pause conditions
            If pauseCommand Then State = EState.Pause
            Status = Translate("Rinsing to machine") & Timer.ToString(1)

          Case EState.TransferEmpty2
            If .ReserveLevel > 10 Then
              Status = commandName_ & (" ") & (.ReserveLevel / 10).ToString("#0.0") & "% " & Timer.ToString
              Timer.Seconds = .Parameters.ReserveTimeAfterRinse
            Else
              Status = commandName_ & Timer.ToString(1)
            End If
            If Timer.Finished Then
              State = EState.RinseToDrain
              Timer.Seconds = MinMax(.Parameters.ReserveRinseToDrainTime, 5, 120)
              NumberOfRinses = .Parameters.ReserveRinseToDrainNumber
              'Tell the control system to step on
              Return True
            End If
            ' Update Restart State
            StateRestart = EState.TransferStart
            ' Lost Safe conditions
            If Not safe Then State = EState.WaitForSafe
            ' Pause conditions
            If pauseCommand Then State = EState.Pause
            Status = Translate("Transferring") & Timer.ToString(1)


            '********************************************************************************************
            '******   WALL WASHING RINSE TO DRAIN
            '********************************************************************************************
          Case EState.RinseToDrainPause
            If pauseCommand Then Timer.Seconds = 2
            If Timer.Finished Then
              State = EState.RinseToDrain
              Timer.Seconds = MinMax(.Parameters.ReserveRinseToDrainTime, 5, 120)
            End If
            Status = Translate("Paused") & Timer.ToString(1)
            If .Parent.IsPaused Then Status = Translate("Paused")
            If .EStop Then Status = Translate("Paused") & (", ESTOP")

          Case EState.RinseToDrain
            If Timer.Finished Then
              NumberOfRinses -= 1
              State = EState.TransferToDrain
              Timer.Seconds = .Parameters.ReserveDrainTime
            End If
            ' Update Restart State
            StateRestart = EState.RinseToDrain
            ' Pause conditions
            If pauseCommand Then State = EState.RinseToDrainPause
            Status = Translate("Rinsing To Drain") & Timer.ToString(1)

          Case EState.TransferToDrain
            If .ReserveLevel > 10 Then
              Status = Translate("Draining") & (" ") & (.ReserveLevel / 10).ToString("#0.0") & "% " & Timer.ToString
              Timer.Seconds = .Parameters.ReserveDrainTime
            Else
              Status = Translate("Draining") & Timer.ToString(1)
            End If
            If Timer.Finished Then
              If (NumberOfRinses > 0) Then
                State = EState.RinseToDrain
                Timer.Seconds = .Parameters.ReserveRinseToDrainTime
              Else
                State = EState.Complete
                Timer.Seconds = 5
              End If
            End If
            ' Update Restart State
            StateRestart = EState.RinseToDrain
            ' Pause conditions
            If pauseCommand Then State = EState.RinseToDrainPause

          Case EState.Complete
            If Timer.Finished Then Cancel()
            Status = Translate("Completing") & Timer.ToString(1)

        End Select
      Loop Until (StartState = State)

    End With
  End Function

  Public Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
    timerWaitReady_.Stop()
  End Sub

  Public ReadOnly Property IsOn() As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property

  Public ReadOnly Property IsActive As Boolean
    Get
      Return (State >= EState.WaitReady) AndAlso (State <= EState.TransferToDrain)
    End Get
  End Property

  Public ReadOnly Property IsWaitReady As Boolean
    Get
      Return (State = EState.WaitReady)
    End Get
  End Property

  Public ReadOnly Property IsOverrun() As Boolean
    Get
      Return (State = EState.WaitReady) AndAlso TimerOverrun.Finished
    End Get
  End Property

  Public ReadOnly Property IsForeground As Boolean
    Get
      Return (State > EState.Off) AndAlso (State <= EState.RinseToDrainPause)
    End Get
  End Property

  Public ReadOnly Property IsBackground As Boolean
    Get
      Return (State >= EState.RinseToDrainPause) AndAlso (State <= EState.Complete)
    End Get
  End Property

  Public ReadOnly Property IsFlowReverse As Boolean
    Get
      Return (State = EState.TransferGravity) OrElse (State = EState.TransferEmpty1)
    End Get
  End Property

  Public ReadOnly Property IsTransfer() As Boolean
    Get
      Return (State >= EState.TransferGravity) AndAlso (State <= EState.TransferEmpty2)
    End Get
  End Property

#Region " Public IO properties "

  Public ReadOnly Property IoTopWash As Boolean
    Get
      Return IsTransfer
    End Get
  End Property

  Public ReadOnly Property IoReserveFillCold() As Boolean
    Get
      Return (State = EState.RinseToTransfer) OrElse (State = EState.RinseToDrain)
    End Get
  End Property

  Public ReadOnly Property IoReserveHeat As Boolean
    Get
      Return False
    End Get
  End Property

  Public ReadOnly Property IoReserveDrain() As Boolean
    Get
      Return (State = EState.RinseToDrain) OrElse (State = EState.TransferToDrain)
    End Get
  End Property

  Public ReadOnly Property IoReserveMixer As Boolean
    Get
      Return False ' TODO
    End Get
  End Property

  ReadOnly Property IoReserveTransfer As Boolean
    Get
      Return IsTransfer OrElse (State = EState.RinseToTransfer) OrElse (State = EState.TransferEmpty2)
    End Get
  End Property

  ReadOnly Property IoSystemBlock As Boolean
    Get
      Return (State = EState.TransferGravity) OrElse IsTransfer
    End Get
  End Property
#End Region


#If 0 Then
  TODO updated copy:'

  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===============================================================================================
'RT - Reserve tank Transfer command
'===============================================================================================
Option Explicit
Implements ACCommand
Public Enum RTState
  RTOff
  RTPause
  RTWaitReady
  RTInterlock
  RTGravityTransfer
  RTTransfer1
  RTRinse
  RTTransfer2
  RTRinseToDrain
  RTTransferToDrain
End Enum
Public State As RTState
Public StateString As String
Public RTStateWas As RTState
Public Timer As New acTimer
Public OverrunTimer As New acTimer

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Name=Reserve Transfer\r\nMinutes=5\r\nHelp=Transfers the reserve tank to the vessel."
'Early bind ControlCode for speed and convenience
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject

  With ControlCode
    .AC.ACCommand_Cancel
    .AT.ACCommand_Cancel: .CO.ACCommand_Cancel: .DR.ACCommand_Cancel
    .FI.ACCommand_Cancel: .HD.ACCommand_Cancel: .HE.ACCommand_Cancel
    .HC.ACCommand_Cancel: .LD.ACCommand_Cancel: .PH.ACCommand_Cancel
    .RC.ACCommand_Cancel: .RI.ACCommand_Cancel: .RW.ACCommand_Cancel
    .SA.ACCommand_Cancel: .TC.ACCommand_Cancel: .TM.ACCommand_Cancel
    .TP.ACCommand_Cancel: .UL.ACCommand_Cancel: .WK.ACCommand_Cancel
    .WT.ACCommand_Cancel
    .AirpadOn = False
    .TemperatureControl.Cancel
    .TemperatureControlContacts.Cancel
    OverrunTimer = 300
    State = RTWaitReady
    .DrugroomPreview.WaitingAtTransferTimer.Start
  End With
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
Dim ControlCode As ControlCode: Set ControlCode = ControlObject
  With ControlCode
  
  'pause the transfer and remember where we were
    If .IO_EStop_PB Then
      If State > RTPause Then
        RTStateWas = State
        State = RTPause
        Timer.Pause
      End If
    End If
    
    Select Case State
      
      Case RTOff
        StateString = ""
        
      Case RTWaitReady
        'Local Add
        If .RP.IsOn And Not .RP.IsReady Then
          StateString = .RP.StateString
          Timer = 10
        End If
        If .RF.IsOn And (.RF.State <> HeatMaintain) Then
          StateString = .RF.StateString
          Timer = 10
        End If
        'Drugroom Add
        If (.KA.IsActive And (.KA.Destination = 82)) Then
          StateString = .KA.StateString
          If Not .ReserveReady Then Timer = 10
        End If
        If (.KP.KP1.IsOn And (.KP.KP1.DispenseTank = 2)) Then
          StateString = .KP.KP1.StateString
          If Not .ReserveReady Then Timer = 10
        End If
        If (.LA.KP1.IsOn And (.LA.KP1.DispenseTank = 2)) Then
          StateString = .LA.KP1.StateString
          If Not .ReserveReady Then Timer = 10
        End If
        If .Parent.IsPaused Then
          .DrugroomPreview.WaitingAtTransferTimer.Pause
        End If
        If (Not .Parent.IsPaused) And .DrugroomPreview.WaitingAtTransferTimer.IsPaused Then
          .DrugroomPreview.WaitingAtTransferTimer.Restart
        End If
        If Timer.Finished Then
          .RF.ACCommand_Cancel
          .DrugroomPreview.WaitingAtTransferTimer.Pause
          OverrunTimer = .Parameters.StandardReserveTransferTime * 60
          State = RTInterlock
        End If
  
      Case RTInterlock
        If .TempSafe And Not .PressSafe Then
          StateString = "RT: Wait Safe " & .SafetyControl.StateString
        Else
          StateString = "RT: Not Safe "
        End If
        If (.MachineSafe Or .HD.HDCompleted) And .LidLocked Then
          State = RTGravityTransfer
          Timer = 5
        End If
        
      Case RTPause
        StateString = "RT: Paused "
        If Not (.IO_EStop_PB) Then
          State = RTStateWas
          Timer.Restart
        End If
     
      Case RTGravityTransfer
        StateString = "RT: Gravity transfer " & TimerString(Timer.TimeRemaining)
        'Make sure lid is locked - else pause
        If Not .LidLocked Then State = RTInterlock
        'Complete Reserve Transfer
        If Timer.Finished Then
           .PumpRequest = True
           State = RTTransfer1
        End If
        
      Case RTTransfer1
        If (.ReserveLevel > 10) Then
          StateString = "RT: Transfer to Machine " & Pad(.ReserveLevel / 10, "0", 3) & "%"
          Timer = .Parameters.ReserveTimeBeforeRinse
        Else: StateString = "RT: Transfer to Machine " & TimerString(Timer.TimeRemaining)
        End If
        'Make sure lid is locked - else pause
        If Not .LidLocked Then State = RTInterlock
        If Timer.Finished Then
          If .RF.IsOn And .RF.FillType = 86 Then
            State = RTRinseToDrain
            Timer = .Parameters.ReserveRinseToDrainTime
          Else
            State = RTRinse
            Timer = .Parameters.ReserveRinseTime
          End If
        End If
      
      Case RTRinse
        StateString = "RT: Reserve Rinse to Machine " & TimerString(Timer.TimeRemaining)
        If Timer.Finished Then
          State = RTTransfer2
          Timer = .Parameters.ReserveTimeAfterRinse
        End If
      
      Case RTTransfer2
        If (.ReserveLevel > 10) Then
          StateString = "RT: Reserve Transfer to Machine " & Pad(.ReserveLevel / 10, "0", 3) & "%"
          Timer = .Parameters.ReserveTimeAfterRinse
        Else: StateString = "RT: Transfer to Machine " & TimerString(Timer.TimeRemaining)
        End If
        'Make sure lid is locked - else pause
        If Not .LidLocked Then State = RTInterlock
        If Timer.Finished Then
          State = RTRinseToDrain
          Timer = .Parameters.ReserveRinseToDrainTime
        End If
      
      Case RTRinseToDrain
        StateString = "RT: Reserve Rinse To Drain " & TimerString(Timer.TimeRemaining)
        If Timer.Finished Then
          State = RTTransferToDrain
          Timer = .Parameters.ReserveDrainTime
        End If
      
      Case RTTransferToDrain
        If (.ReserveLevel > 10) Then
          StateString = "RT: Reserve Transfer to Drain " & Pad(.RecordedLevel / 10, "0", 3) & "%"
          Timer = .Parameters.ReserveDrainTime
        Else: StateString = "RT: Reserve Transfer to Drain " & TimerString(Timer.TimeRemaining)
        End If
        If Timer.Finished Then
          StepOn = True
          .ReserveReady = False
          .ReserveReadyFromPB = False
          State = RTOff
          .RP.ACCommand_Cancel
          If .RF.IsOn Then .RF.ACCommand_Cancel
          .AirpadOn = True
        End If
   
    End Select
  End With
End Sub
Friend Sub ACCommand_Cancel()
  State = RTOff
End Sub
Friend Property Get IsOn() As Boolean
  IsOn = (State <> RTOff)
End Property
Friend Property Get IsActive() As Boolean
  IsActive = (State > RTWaitReady) And (State <= RTTransferToDrain)
End Property
Friend Property Get IsInterlocked() As Boolean
  If (State = RTInterlock) Then IsInterlocked = True
End Property
Friend Property Get IsGravityTransfer() As Boolean
  If (State = RTGravityTransfer) Then IsGravityTransfer = True
End Property
Friend Property Get IsTransfer() As Boolean
  IsTransfer = (State = RTTransfer1) Or (State = RTRinse) Or (State = RTTransfer2)
End Property
Friend Property Get IsTransferToDrain() As Boolean
 IsTransferToDrain = IsRinseToDrain Or (State = RTTransferToDrain)
End Property
Friend Property Get IsRinse() As Boolean
  If (State = RTRinse) Then IsRinse = True
End Property
Friend Property Get IsRinseToDrain() As Boolean
  If (State = RTRinseToDrain) Then IsRinseToDrain = True
End Property
Friend Property Get IsPaused() As Boolean
  IsPaused = (State = RTPause)
End Property
Friend Property Get IsWaitReady() As Boolean
  IsWaitReady = (State = RTWaitReady)
End Property
Friend Property Get IsDelayedWaitReady() As Boolean
  IsDelayedWaitReady = OverrunTimer.Finished And IsWaitReady
End Property
Friend Property Get IsDelayedTransfer() As Boolean
  IsDelayedTransfer = OverrunTimer.Finished And (State > RTWaitReady)
End Property
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property




#End If
End Class
