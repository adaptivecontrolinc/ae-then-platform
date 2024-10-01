'American & Efird - Mt. Holly Then Platform
' Version 2024-09-11

Imports Utilities.Translations

<Command("Load", "", "", "", "'StandardTimeLoad=30"), TranslateCommand("es", "Introducir Cargador"),
Description("Signal Operator to Unload Machine."),
TranslateDescription("es", "Señala a operador para cargar la máquina y comienza a registrar el tiempo de carga"),
Category("Operator Functions"), TranslateCategory("es", "Operator Functions")>
Public Class LD : Inherits MarshalByRefObject : Implements ACCommand
  Private Const commandName_ As String = ("LD: ")

  'Command States
  Public Enum Estate
    Off
    Interlock
    DrainEmpty
    Flush
    DrainEmpty2
    Load
    CloseLid
    Done
  End Enum
  Property State As Estate
  Property Status As String

  Public ReadOnly Timer As New Timer
  Friend ReadOnly TimerAdvance As New Timer
  Public ReadOnly Property TimeAdvance As Integer
    Get
      Return TimerAdvance.Seconds
    End Get
  End Property
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
  End Sub

  Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With controlCode

      ' Cancel Commands
      .AC.Cancel()
      If .AT.IsForeground Then .AT.Cancel()
      .KA.Cancel() : .WK.Cancel()
      .DR.Cancel() : .FI.Cancel() : .HC.Cancel() : .HD.Cancel()
      .RI.Cancel() : .RC.Cancel() : .RH.Cancel() : .TM.Cancel()
      '.LD.Cancel() : 
      .PH.Cancel() : .SA.Cancel() : .UL.Cancel()
      If .RF.IsForeground AndAlso .RF.FillType = EFillType.Vessel Then .RF.Cancel()   ' TODO Check
      If .RT.IsForeground Then .RT.Cancel()
      .WT.Cancel()

      ' Cancel PressureOn command if active during Atmospheric Commands
      .PR.Cancel()

      .AirpadOn = False

      ' Clear Parent.Signal 
      If .Parent.Signal <> "" Then .Parent.Signal = ""

      'Set default start state
      State = Estate.Interlock
      Timer.Seconds = 1
      TimerOverrun.Minutes = .Parameters.StandardTimeLoad

      'This is a foreground command so do not step on
      Return False
    End With
  End Function

  Function Run() As Boolean Implements ACCommand.Run
    With controlCode

      ' Force state machine to interlock state if the machine is not safe
      Dim safe As Boolean = .MachineSafe
      If (State > Estate.Interlock) AndAlso (Not safe) Then State = Estate.Interlock

      ' Reset AdvancePB timer 
      If (Not .AdvancePb) OrElse (.Parent.Signal <> "") Then TimerAdvance.Seconds = 2

      Select Case State

        Case Estate.Off
          Status = (" ")

        Case Estate.Interlock
          Status = Translate("Wait Safe") & Timer.ToString(1)
          If Not safe Then
            If Not .MachineSafe Then
              ' Waiting for SafetyControl.IsDepressurized or PressSafe or TempSafe
              Status = Translate("Machine Not Safe")
            End If
            Timer.Seconds = 5
          Else
            Status = Translate("Load") & (" ") & Translate("Starting") & Timer.ToString(1)
          End If
          If Timer.Finished Then
            ' Confirm Manual Drain, if level exists
            If (.MachineLevel > 100) Then
              Status = Translate("Not Empty, Hold Advance to Drain") & TimerAdvance.ToString(1)
              If .AdvancePb Then Status = Translate("Drain Starting") & TimerAdvance.ToString(1)
              If TimerAdvance.Finished Then
                ' Cancel Temperature Commands
                .CO.Cancel() : .HE.Cancel()
                .GetRecordedLevel = False
                State = Estate.DrainEmpty
                Timer.Seconds = .Parameters.DrainMachineTime
              End If
            Else
              ' Step to Loading State
              State = Estate.Load
              .Parent.Signal = Translate("Load")
            End If
          End If

        Case Estate.DrainEmpty
          .PumpControl.StopMainPump()
          .GetRecordedLevel = False
          .Alarms.LosingLevelInTheVessel = False
          If (.MachineLevel > 50) Then Timer.Seconds = .Parameters.DrainMachineTime
          If Timer.Finished Then
            Timer.Seconds = MinMax(.Parameters.LevelGaugeFlushTime, 10, 60)
            State = Estate.Flush
          End If
          If .MachineLevel > 50 Then
            Status = Translate("Draining") & (" ") & (.MachineLevel / 10).ToString("#0.0") & "%"
          Else
            Status = Translate("Draining") & Timer.ToString(1)
          End If

        Case Estate.Flush
          If Timer.Finished Then
            State = Estate.DrainEmpty2
            Timer.Seconds = MinMax(.Parameters.LevelGaugeSettleTime, 10, 120)
          End If
          Status = Translate("Flushing Level") & Timer.ToString(1)

        Case Estate.DrainEmpty2
          If Timer.Finished Then  ' Wait for Level to settle first
            If .MachineLevel > 50 Then Timer.Seconds = MinMax(.Parameters.DrainMachineTime, 10, 120)
            If Timer.Finished Then
              '.MachineVolume = 0 ' TODO
              State = Estate.Load
              .Parent.Signal = Translate("Load")
            End If
            If .MachineLevel > 50 Then
              Status = Translate("Draining") & (" 2 ") & (.MachineLevel / 10).ToString("#0.0") & "%"
            Else
              Status = Translate("Draining") & Timer.ToString(1)
            End If
          End If

        Case Estate.Load
          Status = Translate("Load")
          If .Parent.Signal = "" Then
            If Not .MachineClosed Then
              If Not .MachineClosed Then Status = Translate("Machine Not Closed")
            Else
              ' Step to Confirm the machine is closed
              State = Estate.CloseLid
              TimerAdvance.Seconds = 2
            End If
          End If

        Case Estate.CloseLid
          If Not .AdvancePb Then
            Status = Translate("Load") & (", ") & Translate("Hold Advance")
            TimerAdvance.Seconds = 2
          Else
            'Attempting to Advance
            If .Parent.Signal <> "" Then .Parent.Signal = ""
            Status = Translate("Advancing") & (" ") & TimerAdvance.ToString
            If TimerAdvance.Finished Then
              .IO.AdvancePb = False      ' Clear Simulation
              State = Estate.Done
              Timer.Seconds = 5
            End If
          End If

        Case Estate.Done
          Status = Translate("Completing") & Timer.ToString(1)
          If Timer.Finished Then
            If .Parent.Signal <> "" Then .Parent.Signal = ""
            Cancel()
          End If

      End Select
    End With
  End Function

  Sub Cancel() Implements ACCommand.Cancel
    State = Estate.Off
    Status = ""
    Timer.Cancel()
    TimerOverrun.Cancel()
  End Sub

  ReadOnly Property IsOn As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> Estate.Off)
    End Get
  End Property

  ReadOnly Property IsConfirmPortSwitch As Boolean
    Get
      Return (State = Estate.CloseLid)
    End Get
  End Property

  ReadOnly Property IsOverrun As Boolean
    Get
      Return State = Estate.Load AndAlso TimerOverrun.Finished
    End Get
  End Property

  ReadOnly Property IoDrain As Boolean
    Get
      Return (State = Estate.DrainEmpty) OrElse (State = Estate.Flush) OrElse (State = Estate.DrainEmpty2)
    End Get
  End Property

  ReadOnly Property IoLampSignal As Boolean
    Get
      Return ((State = Estate.Interlock) AndAlso (controlCode.MachineLevel > 100) AndAlso controlCode.FlashFast) OrElse
              (((State = Estate.Load) OrElse (State = Estate.CloseLid)) AndAlso controlCode.FlashSlow)
    End Get
  End Property

  ReadOnly Property IoTopWash As Boolean
    Get
      Return (State = Estate.DrainEmpty) OrElse (State = Estate.Flush) OrElse (State = Estate.DrainEmpty2)
    End Get
  End Property



#If 0 Then
' TODO

'Load command
Option Explicit
Implements ACCommand
Public Enum LDState
  LDOff
  LDUnsafe
  LDCloseLid
  LDOn
End Enum
Public State As LDState
Public StateString As String
Public LDSignalOnRequest As Boolean
Public Timer As New acTimer
Public AdvanceTimer As New acTimer
Public OverrunTimer As New acTimer

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Name=Load Packages\r\nMinutes=5\r\nHelp=Signals the operator to Load the machine and starts recording the load time. If the load time exceeds the parameter Standard Load Time then the overrun time is logged as Load Delay."
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode
     .CO.ACCommand_Cancel: .DR.ACCommand_Cancel
    .FI.ACCommand_Cancel: .HD.ACCommand_Cancel: .HE.ACCommand_Cancel
    .PH.ACCommand_Cancel: .PR.ACCommand_Cancel: .RC.ACCommand_Cancel
    .RH.ACCommand_Cancel: .RI.ACCommand_Cancel
    If .RF.FillType = 86 Then .RF.ACCommand_Cancel: .RT.ACCommand_Cancel
    .RW.ACCommand_Cancel:.TC.ACCommand_Cancel
    .TM.ACCommand_Cancel: 

    .TemperatureControl.Cancel
    .TemperatureControlContacts.Cancel
    .AirpadOn = False
    .PumpRequest = False
    OverrunTimer = .Parameters.StandardLoadTime * 60
    AdvanceTimer = 2
    State = LDUnsafe

  End With
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode
    
    Select Case State
      Case LDOff
        StateString = ""
        
      Case LDUnsafe
        If Not .MachineSafe Then
          StateString = .SafetyControl.StateString
        Else
          State = LDOn
          AdvanceTimer = 2
          OverrunTimer = .Parameters.StandardLoadTime * 60
          LDSignalOnRequest = True
        End If
  
      Case LDCloseLid
        StateString = "Close Lid to Continue"
        If .LidLocked Then
          AdvanceTimer = 2
          State = LDOn
        End If
  
      Case LDOn
        StateString = "Load, Hold Run to Complete " & TimerString(AdvanceTimer.TimeRemaining)
        If Not .IO_RemoteRun Then AdvanceTimer = 2
        If .IO_RemoteRun And Not .LidLocked Then
          State = LDCloseLid
        End If
        If Not .MachineSafe Then State = LDUnsafe
        If AdvanceTimer.Finished Then
           ACCommand_Cancel
        End If
        
    End Select
  End With
End Sub

Friend Sub ACCommand_Cancel()
  State = LDOff
  LDSignalOnRequest = False
End Sub
Friend Property Get IsOn() As Boolean
  IsOn = (State <> LDOff)
End Property
Friend Property Get IsLoad() As Boolean
  If State = LDOn Then IsLoad = True
End Property
Friend Property Get IsCloseLid() As Boolean
  If State = LDCloseLid Then IsCloseLid = True
End Property
Friend Property Get IsUnSafe() As Boolean
  If State = LDUnsafe Then IsUnSafe = True
End Property
Friend Property Get IsOverrun() As Boolean
  If ((State = LDOn) Or (State = LDCloseLid)) And OverrunTimer.Finished Then IsOverrun = True
End Property
Private Property Get ACCommand_IsOn() As Boolean
ACCommand_IsOn = IsOn
End Property

#End If
End Class
