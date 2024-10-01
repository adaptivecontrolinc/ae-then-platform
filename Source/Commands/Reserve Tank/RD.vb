'American & Efird - MX Then Multiflex Package
' Version 2024-09-04

Imports Utilities.Translations

<Command("Reserve Drain", "Rinses:|0-9|", "", "", ""), _
TranslateCommand("es", "Reserve Drain", "Rinses:|0-9|"), _
Description("Drain and Rinse the reserve tank the specified number of times."), _
TranslateDescription("es", "Escurra y enjuague el tanque de reserva el número especificado de veces."),
Category("Reserve Tank Commands"), TranslateCategory("es", "Reserve Tank Commands")>
Public Class RD : Inherits MarshalByRefObject : Implements ACCommand
  Private Const commandName_ As String = ("RD: ")
  Private ReadOnly controlCode As ControlCode

  Public Enum EState
    Off
    WaitIdle
    RinseToDrain
    Drain

    ' Repeat if necessary
    Complete
  End Enum
  Public State As EState
  Public StateRestart As EState
  Public Status As String
  Public NumberOfRinses As Integer
  Property IsCancelling As Boolean

  Public Property Timer As New Timer
  Friend TimerOverrun As New Timer
  Public ReadOnly Property TimeOverrun As Integer
    Get
      Return TimerOverrun.Seconds
    End Get
  End Property

  Public Sub New(ByVal controlCode As ControlCode)
    Me.controlCode = controlCode
    Me.controlCode.Commands.Add(Me)
  End Sub

  Public Sub ParametersChanged(ByVal ParamArray param() As Integer) Implements ACCommand.ParametersChanged
    'If the command is not running then do nothing the changes will be picked up when the command starts
    If Not IsOn Then Exit Sub

    'Check array bounds just to be on the safe side
    If param.GetUpperBound(0) >= 1 Then NumberOfRinses = param(1)
  End Sub

  Public Function Start(ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With controlCode

      'Check array bounds just to be on the safe side
      If param.GetUpperBound(0) >= 1 Then NumberOfRinses = param(1)

      'Initialize Timers
      Timer.TimeRemaining = 5

      'Set default state
      State = EState.WaitIdle

      'This is a background command, but wait until after Sync State to step on
      Return False
    End With
  End Function

  Public Function Run() As Boolean Implements ACCommand.Run
    With controlCode

      Select Case State
        Case EState.Off
          Status = (" ")
          StateRestart = EState.Off

        Case EState.WaitIdle
          Status = commandName_ & Translate("Wait Idle") & Timer.ToString(1)
          ' Cancel any other Background commands:
          If .RF.IsOn Then .RF.Cancel()
          If .RP.IsOn Then .RP.Cancel()
          ' Cancel active ReserveTransfer commands (we've jumped to this command)
          If .RT.IsForeground Then .RT.Cancel()
          If .RT.IsBackground Then Status = .RT.Status : Timer.Seconds = 1
          ' Kitchen Tank
          If (.KA.IsOn AndAlso .KA.Destination = EKitchenDestination.Reserve) Then Status = .KA.Status : Timer.Seconds = 1
          If (.KP.KP1.IsForeground AndAlso (.KP.KP1.DispenseTank = 2)) Then Status = .KP.KP1.DrugroomDisplay : Timer.Seconds = 1
          If (.LA.KP1.IsOn AndAlso (Not .LA.KP1.IsBackground) AndAlso (.LA.KP1.Param_Destination = EKitchenDestination.Reserve)) Then Status = .LA.KP1.DrugroomDisplay : Timer.Seconds = 1
          If (.Parent.IsPaused OrElse .EStop) Then Timer.Seconds = 1
          If .Parent.IsPaused Then Status = commandName_ & Translate("Paused") : Timer.Seconds = 1
          If .EStop Then Status = commandName_ & Translate("EStop") : Timer.Seconds = 1
          ' Reserve tank idle
          If Timer.Finished Then
            'Step to the next state
            State = EState.RinseToDrain
            Timer.Seconds = MinMax(.Parameters.ReserveRinseToDrainTime, 5, 60)
            TimerOverrun.Seconds = (NumberOfRinses * .Parameters.ReserveDrainTime) + 300

            'This is a background command, now we can step on (it's ok to step on after the RT has completed the foreground transfer)
            Return True
          End If


          '********************************************************************************************
          '******   RESERVE RINSE TO DRAIN LOGIC
          '********************************************************************************************
        Case EState.RinseToDrain
          If Timer.Finished OrElse (.ReserveLevel > 750) Then
            Timer.Seconds = MinMax(.Parameters.ReserveDrainTime, 5, 60)
            State = EState.Drain
          End If
          Status = commandName_ & Translate("Rinsing To Drain") & (" ") & (.ReserveLevel / 10).ToString("#0.0") & Timer.ToString(1)

        Case EState.Drain
          If .ReserveLevel > 10 Then
            Status = commandName_ & Translate("Draining") & (" ") & (.ReserveLevel / 10).ToString("#0.0") & ("%")
            Timer.Seconds = MinMax(.Parameters.ReserveDrainTime, 5, 60)
          Else
            Status = commandName_ & Translate("Draining") & Timer.ToString(1)
          End If
          If Timer.Finished Then
            NumberOfRinses -= 1
            If (NumberOfRinses > 0) AndAlso (Not IsCancelling) Then
              ' Repeat total rinse/wash/drain
              State = EState.RinseToDrain
              Timer.Seconds = MinMax(.Parameters.ReserveRinseToDrainTime, 5, 60)
            Else
              State = EState.Complete
              Timer.Seconds = 5
            End If
          End If


        Case EState.Complete
          Status = commandName_ & Translate("Completing") & Timer.ToString(1)
          If Timer.Finished Then
            ' Cancelling or finished rinses
            State = EState.Off
            NumberOfRinses = 0
            IsCancelling = False
          End If

      End Select
    End With
  End Function

  Public Sub Cancel() Implements ACCommand.Cancel
    NumberOfRinses = 0
    If (State = EState.WaitIdle) OrElse (State = EState.RinseToDrain) Then
      IsCancelling = True
      State = EState.Drain
    Else
      Status = ""
      State = EState.Off
    End If
    TimerOverrun.Seconds = 0    ' Clear the timer
  End Sub

  Public ReadOnly Property IsOn As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property

  Public ReadOnly Property IsWaitIdle As Boolean
    Get
      Return (State = EState.WaitIdle)
    End Get
  End Property

  Public ReadOnly Property IsOverrun() As Boolean
    Get
      Return (State > EState.WaitIdle) AndAlso TimerOverrun.Finished
    End Get
  End Property

#Region " PUBLIC I/O PROPERTIES "

  Public ReadOnly Property IoReserveFillCold() As Boolean
    Get
      Return (State = EState.RinseToDrain)
    End Get
  End Property

  Public ReadOnly Property IoDrain() As Boolean
    Get
      Return (State = EState.RinseToDrain) OrElse (State = EState.Drain)
    End Get
  End Property

#End Region


#If 0 Then
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===============================================================================================
'RD - Kitchen tank prepare command
'===============================================================================================
Option Explicit
Implements ACCommand
Public Enum RDState
  RDOff
  RDWaitForRF
  RDDrainEmpty
  RDRinseToDrain
  RDDrainRinse
End Enum
Public State As RDState
Public StateString As String
Public NumberOfRinsesDrains As Long
Public RDStateWas As RDState
Public Timer As New acTimer

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Parameters= |0-9|\r\nName=Reserve Drain\r\nHelp=drain and rinse the reserve tank for the specified number of times."
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  NumberOfRinsesDrains = Param(1)
  
  With ControlCode
    .RT.ACCommand_Cancel
    .RW.ACCommand_Cancel

    'start timer for 5 minutes: if RF takes more than 5 minutes, cancel the RF to prevent background funciton overlap
    'example: KP waits for RD waiting for RF....
    Timer = 300
    State = RDWaitForRF
    StepOn = True
  End With
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
  Dim ControlCode As ControlCode: Set ControlCode = ControlObject
  With ControlCode
  
    Select Case State
      
      Case RDOff
        StateString = ""
        
      Case RDWaitForRF
        If .RF.StateString <> "" Then
          StateString = "RD: " & .RF.StateString
        Else: StateString = "RD: Wait for RF " & TimerString(Timer.TimeRemaining)
        End If
        If Timer.Finished Then .RF.ACCommand_Cancel
        If (Not .RF.IsOn) Or (.RF.State = HeatMaintain) Then
           .RF.ACCommand_Cancel
           State = RDDrainEmpty
           Timer = .Parameters.ReserveDrainTime
        End If
        
      Case RDDrainEmpty
        If (.ReserveLevel > 10) Then
          StateString = "RD: Drain Empty " & Pad(.ReserveLevel / 10, "0.0", 3) & "%"
          Timer = .Parameters.ReserveDrainTime
        Else: StateString = "RD: Drain Empty " & TimerString(Timer.TimeRemaining)
        End If
        If Timer.Finished Then
          State = RDRinseToDrain
          Timer = .Parameters.ReserveRinseToDrainTime
        End If
       
      Case RDRinseToDrain
        StateString = "RD: Rinse to Drain " & TimerString(Timer.TimeRemaining)
        If Timer.Finished Then
          State = RDDrainRinse
          Timer = .Parameters.ReserveDrainTime
        End If
       
      Case RDDrainRinse
        If (.ReserveLevel > 10) Then
          StateString = "RD: Drain Empty " & Pad(.ReserveLevel / 10, "0.0", 3) & "%"
          Timer = .Parameters.ReserveDrainTime
        Else: StateString = "RD: Drain Empty " & TimerString(Timer.TimeRemaining)
        End If
        If Timer.Finished Then
          NumberOfRinsesDrains = NumberOfRinsesDrains - 1
          If NumberOfRinsesDrains > 0 Then
             State = RDRinseToDrain
             Timer = .Parameters.ReserveRinseToDrainTime
          Else
             State = RDOff
          End If
        End If
   
    End Select
  End With
End Sub

Friend Sub ACCommand_Cancel()
  State = RDOff
End Sub
Friend Property Get IsOn() As Boolean
  IsOn = (State <> RDOff)
End Property
Friend Property Get IsActive() As Boolean
  IsActive = (State > RDOff) And (State <= RDDrainRinse)
End Property
Friend Property Get IsWaitingForRF() As Boolean
  If (State = RDWaitForRF) Then IsWaitingForRF = True
End Property
Friend Property Get IsTransferToDrain() As Boolean
 IsTransferToDrain = IsRinseToDrain Or (State = RDDrainRinse) Or (State = RDDrainEmpty)
End Property
Friend Property Get IsRinseToDrain() As Boolean
  If (State = RDRinseToDrain) Then IsRinseToDrain = True
End Property
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property
  
#End If
End Class