'American & Efird - Mt. Holly Then Platform to GMX
' Version 2024-09-17

Imports Utilities.Translations

<Command("Unload Machine", "", "", "", "'StandardTimeUnload=10"),
TranslateCommand("es", "Sacar Cargador"),
Description("Signal Operator to Unload Machine."),
TranslateDescription("es", "Se�ala a operador para descargar el paquete de la m�quina."),
Category("Operator Functions"), TranslateCategory("es", "Operator Functions")>
Public Class UL
  Inherits MarshalByRefObject
  Implements ACCommand
  Private Const commandName_ As String = ("UL: ")

  'Command States
  Public Enum EState
    Off
    Interlock
    Depressurize
    DrainEmpty
    UnloadSignal
    Unload
    DrainEmpty2
    Done
  End Enum
  Property State As EState
  Property StateString As String
  Private flashSlow_ As Boolean
  Private flashFast_ As Boolean

  Property Timer As New Timer
  Friend TimerAdvance As New Timer
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
      .DR.Cancel() : .FI.Cancel() : .HC.Cancel() : .HD.Cancel() : .PR.Cancel()
      .RC.Cancel() : .RH.Cancel() : .RI.Cancel() : .TM.Cancel() ' TC.Cancel
      .LD.Cancel() : .PH.Cancel() : .SA.Cancel() ': .UL.Cancel()
      If .RT.IsForeground Then .RT.Cancel()
      .WT.Cancel()

      ' Clear Parent.Signal 
      If .Parent.Signal <> "" Then .Parent.Signal = ""

      State = EState.Interlock
      Timer.Seconds = 1
      TimerOverrun.Minutes = .Parameters.StandardTimeUnload

    End With
  End Function
  Function Run() As Boolean Implements ACCommand.Run
    With controlCode

      ' Flashers in sync
      flashSlow_ = .FlashSlow
      flashFast_ = .FlashFast

      ' Force state machine to interlock state if the machine is not safe
      Dim Safe As Boolean = .MachineSafe

      ' Reset AdvancePB timer 
      If Not .AdvancePb Then TimerAdvance.Seconds = 2

      Select Case State

        Case EState.Off
          StateString = (" ")

        Case EState.Interlock
          StateString = commandName_ & Translate("Wait Safe") & Timer.ToString(1)
          If Not .TempSafe Then Timer.Seconds = 5
          If Timer.Finished Then
            ' Cancel Temperature Commands
            .CO.Cancel() : .HE.Cancel() : .TP.Cancel()
            .TemperatureControl.Cancel()
            .PumpControl.StopMainPump()
            .GetRecordedLevel = False
            .AirpadOn = False
            .PR.Cancel()
            State = EState.Depressurize
          End If

        Case EState.Depressurize
          StateString = commandName_ & Translate("Depressurizing") & Timer.ToString(1)
          If Not .TempSafe Then State = EState.Interlock
          If Not .PressSafe Then Timer.Seconds = 5
          If Timer.Finished Then
            State = EState.DrainEmpty
          End If

        Case EState.DrainEmpty
          If .MachineLevel > 50 Then
            StateString = commandName_ & Translate("Draining") & (" ") & (.MachineLevel / 10).ToString("#0.0") & "%"
            Timer.Seconds = MinMax(.Parameters.DrainMachineTime, 30, 300)
          Else
            StateString = commandName_ & Translate("Draining") & Timer.ToString(1)
          End If
          If .AdvancePb Then StateString = commandName_ & Translate("Advancing") & TimerAdvance.ToString(1)
          If Not Safe Then State = EState.Interlock
          If Timer.Finished OrElse TimerAdvance.Finished Then
            State = EState.Unload
            .Parent.Signal = Translate("Unload")
          End If

        Case EState.UnloadSignal
          If .IO.AdvancePb OrElse (.Parent.Signal = "") Then
            State = EState.Unload
          End If
          StateString = commandName_ & "Unload"

        Case EState.Unload
          StateString = commandName_ & Translate("Unload, Hold Run to Complete") & TimerAdvance.ToString(1)
          If TimerAdvance.Finished Then
            .IO.AdvancePb = False      ' Clear Simulation
            State = EState.DrainEmpty2
            Timer.Seconds = 5
          End If

        Case EState.DrainEmpty2
          If .MachineLevel > 50 Then
            StateString = commandName_ & Translate("Draining") & (" ") & (.MachineLevel / 10).ToString("#0.0") & "%"
            Timer.Seconds = .Parameters.DrainMachineTime
          Else
            StateString = commandName_ & Translate("Draining") & Timer.ToString(1)
          End If
          If Not Safe Then State = EState.Interlock
          If Timer.Finished Then
            State = EState.Done
            Timer.Seconds = 5
          End If

        Case EState.Done
          If Timer.Finished Then
            .IO.AdvancePb = False
            If .Parent.Signal <> "" Then .Parent.Signal = ""
            Cancel()
          End If
          StateString = commandName_ & Translate("Completing") & Timer.ToString(1)

      End Select
    End With
  End Function

  Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
    StateString = ""
    Timer.Cancel()
    TimerOverrun.Cancel()
    TimerAdvance.Cancel()
  End Sub

  ReadOnly Property IsOn As Boolean Implements ACCommand.IsOn
    Get
      IsOn = State <> EState.Off
    End Get
  End Property

  Public ReadOnly Property IsOverrun() As Boolean
    Get
      Return IsOn AndAlso TimerOverrun.Finished
    End Get
  End Property

  ReadOnly Property IoDrain As Boolean
    Get
      Return (State = EState.DrainEmpty) OrElse (State = EState.Unload) OrElse (State = EState.DrainEmpty2)
    End Get
  End Property

  ReadOnly Property IoLampSignal As Boolean
    Get
      Return (State = EState.Unload) AndAlso flashSlow_
    End Get
  End Property


  ' TODO Check & remove
#If 0 Then

  VERSION 1.0 CLASS

'Unload command
Option Explicit
Implements ACCommand
Public Enum ULState
  ULOff
  ULUnsafe
  ULWaitingToUnload
  ULUnloading
End Enum
Public State As ULState
Public StateString As String
Public Timer As New acTimer
Public ULSignalOnRequest As Boolean
Public AdvanceTimer As New acTimer
Public OverrunTimer As New acTimer

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Name=Unload Packages\r\nMinutes=5\r\nHelp=Signals the operator to unload the package from the machine."
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode
    .AC.ACCommand_Cancel
    .AT.ACCommand_Cancel: .CO.ACCommand_Cancel: .DR.ACCommand_Cancel
    .FI.ACCommand_Cancel: .HD.ACCommand_Cancel: .HE.ACCommand_Cancel
    .HC.ACCommand_Cancel: .LD.ACCommand_Cancel: .PH.ACCommand_Cancel
    .RC.ACCommand_Cancel:
    If .RF.FillType = 86 Then .RF.ACCommand_Cancel
    .RH.ACCommand_Cancel: .RI.ACCommand_Cancel
    .RT.ACCommand_Cancel: .RW.ACCommand_Cancel: .SA.ACCommand_Cancel
    .TC.ACCommand_Cancel: .TM.ACCommand_Cancel: .WK.ACCommand_Cancel
    .WT.ACCommand_Cancel
    .TemperatureControl.Cancel
    .TemperatureControlContacts.Cancel
    .AirpadOn = False
    .PumpRequest = False
    OverrunTimer = .Parameters.StandardUnloadTime * 60
    
  End With
  AdvanceTimer = 2
  State = ULUnsafe
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode
    
    Select Case State
      Case ULOff
        StateString = ""
        
      Case ULUnsafe
        If Not .MachineSafe Then
          StateString = .SafetyControl.StateString
        Else
          State = ULWaitingToUnload
          OverrunTimer = .Parameters.StandardUnloadTime * 60
          ULSignalOnRequest = True
        End If
  
      Case ULWaitingToUnload
        StateString = "Waiting to Unload "
        AdvanceTimer = 2
        If .IO_RemoteRun Then State = ULUnloading
        If Not .MachineSafe Then State = ULUnsafe
      
      Case ULUnloading
        StateString = "Unload, Hold Run to Complete " & TimerString(AdvanceTimer.TimeRemaining)
        If Not .IO_RemoteRun Then AdvanceTimer = 2
        If AdvanceTimer.Finished Then ACCommand_Cancel
        If Not .MachineSafe Then
          State = ULUnsafe
        End If
    
    End Select
  End With
End Sub
  
#End If
End Class
