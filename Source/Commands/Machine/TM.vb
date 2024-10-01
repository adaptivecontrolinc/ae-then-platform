'American & Efird
' Version 2024-09-17 - Mt Holly
' Version 2016-09-13 - MX Package

Imports Utilities.Translations

<Command("Run for ", "|0-99|mins", " ", "", "'1"),
TranslateCommand("es", "Mantener Durante", "|0-99|mins"),
Description("Hold for specified time, in minutes."),
TranslateDescription("es", "Mantener durante un tiempo determinado, en cuestión de minutos"),
Category("Machine Functions"), TranslateCategory("es", "Machine Functions")>
Public Class TM : Inherits MarshalByRefObject : Implements ACCommand
  Private Const commandName_ As String = ("TM: ")

  Public Enum EState
    Off
    Start
    HoldForTime
    Done
    PauseCrashCool
    PauseParent
    PausePump
    PauseLid
    Restart

    Interlock
    CheckMachineClosed
    StartPump
    Pressurize
    Paused

  End Enum
  Property State As EState
  Property Status As String
  Property HoldTime As Integer
  Property MachineNotClosed As Boolean

  Property Timer As New Timer

  Friend TimerHold As New Timer
  Public ReadOnly Property TimeHold As Integer
    Get
      Return TimerHold.Seconds
    End Get
  End Property
  Friend Property TimerOverrun As New Timer
  Public ReadOnly Property TimeOverrun As Integer
    Get
      Return TimerOverrun.Seconds
    End Get
  End Property
  Friend TimerAdvance As New Timer
  Public ReadOnly Property TimeAdvance As Integer
    Get
      Return TimerAdvance.Seconds
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
      ' Update Hold Time
      If param.GetUpperBound(0) >= 1 Then HoldTime = param(1) * 60

      ' Set timer
      TimerHold.Seconds = HoldTime
      TimerOverrun.Seconds = HoldTime + 60
    End If
  End Sub

  Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With controlCode

      'Cancel Commands
      .AC.Cancel() : .AT.Cancel()
      .KA.Cancel() : .WK.Cancel()
      .DR.Cancel() : .FI.Cancel() : .HC.Cancel() : .HD.Cancel()
      .RI.Cancel() : .RC.Cancel() : .RH.Cancel() ': .TM.Cancel()
      .LD.Cancel() : .PH.Cancel() : .SA.Cancel() : .UL.Cancel()

      If .RF.FillLevel = EFillType.Vessel Then .RF.Cancel()         ' TODO: Check this
      If .RT.IsForeground Then .RT.Cancel()
      '  .CO.Cancel() : .HC.Cancel() : .HE.Cancel() : 
      .WT.Cancel()

      ' Check array bounds just to be on the safe side
      HoldTime = 0
      If param.GetUpperBound(0) >= 1 Then HoldTime = param(1) * 60

      ' Set timer
      TimerHold.Seconds = HoldTime
      TimerHold.Pause()

      TimerOverrun.Seconds = HoldTime + 60

      ' Set default state
      State = EState.Interlock
      Timer.Seconds = 1
      If Not (.MachineClosed) Then MachineNotClosed = True

      'This is a foreground command so do not step on
      Return False
    End With
  End Function

  Function Run() As Boolean Implements ACCommand.Run
    With controlCode

      ' Safe to continue
      Dim safe As Boolean = (Not .EStop) AndAlso .MachineClosed
      If (State > EState.Interlock) AndAlso Not safe Then
        If Not (.MachineClosed) Then
          MachineNotClosed = True
          .Parent.Signal = "Machine Not Closed"
        End If
        State = EState.Interlock
        TimerHold.Pause()
        Timer.Seconds = 2
      End If

      ' Pause if active and problem occurs
      Dim pauseCommand As Boolean = .EStop OrElse (Not .IO.PumpRunning) OrElse (.Parent.IsPaused) OrElse (.TemperatureControl.IsCrashCoolOn)
      If (State > EState.Paused) AndAlso pauseCommand Then
        State = EState.Paused
        TimerHold.Pause()
      End If

      ' Select State 
      Select Case State
        Case EState.Off
          Status = (" ")

        Case EState.Interlock
          Status = Translate("Interlocked")
          If Not .IO.PumpRunning Then Status = commandName_ & Translate("Pump Off") : Timer.Seconds = 1
          If .TemperatureControl.IsCrashCoolOn Then Status = commandName_ & Translate("Crash-Cooling") : Timer.Seconds = 1
          If .Parent.IsPaused Then Status = commandName_ & Translate("Paused") : Timer.Seconds = 1
          If Not .MachineClosed Then Status = commandName_ & Translate("Machine Not Closed") : Timer.Seconds = 10
          If .EStop Then Status = commandName_ & Translate("EStop") : Timer.Seconds = 1
          If Not (.MachineClosed) Then MachineNotClosed = True
          If Timer.Finished Then
            State = EState.CheckMachineClosed
          End If

        Case EState.CheckMachineClosed
          Status = commandName_ & Translate("Checking Lids") & Timer.ToString(1)
          If Not .MachineClosed Then Status = commandName_ & Translate("Machine Not Closed") : Timer.Seconds = 10
          If Not (.MachineClosed) Then MachineNotClosed = True
          ' Flag set where machine/sample/expansion limit switch was lost - operator must hold advance to clear and continue <<Redundancy>>
          If MachineNotClosed AndAlso safe Then
            If Not .AdvancePb Then
              Status = commandName_ & Translate("Hold Advance") & (" ") & TimerAdvance.ToString
              If .FlashSlow Then Status = commandName_ & ("Lost Machine Closed")
              TimerAdvance.Seconds = 2
            Else
              'Attempting to Advance
              If .Parent.Signal <> "" Then .Parent.Signal = ""
              Status = commandName_ & Translate("Advancing") & (" ") & TimerAdvance.ToString
              If TimerAdvance.Finished Then MachineNotClosed = False
            End If
          End If
          ' Machine is Safe to continue
          If Timer.Finished AndAlso (Not MachineNotClosed) Then
            State = EState.StartPump
            ' Start pump if not running
            If Not .PumpControl.IsActive Then
              .PumpControl.StartAuto()
              Timer.Seconds = .Parameters.FillPrimePumpTime
            End If
          End If

        Case EState.StartPump
          If Not .PumpControl.IsRunning Then Timer.Seconds = .PumpControl.Parameters_PumpAccelerationTime
          If Timer.Finished Then
            .AirpadOn = True
            State = EState.Pressurize
            Timer.Seconds = 5
          End If
          Status = commandName_ & .PumpControl.StateString

        Case EState.Pressurize
          If Not .SafetyControl.IsPressurized Then Timer.Seconds = 5
          If Timer.Finished Then
            State = EState.HoldForTime
            TimerHold.Seconds = HoldTime
          End If
          Status = commandName_ & Translate("Pressurizing") & Timer.ToString(1)

        Case EState.Paused
          If Not .PumpControl.PumpRunning Then Status = commandName_ & Translate("Pump Off")
          If .TemperatureControl.IsCrashCoolOn Then Status = commandName_ & Translate("Crash-Cooling") : Timer.Seconds = 1
          If .Parent.IsPaused Then Status = commandName_ & Translate("Paused") : Timer.Seconds = 1
          If .EStop Then Status = commandName_ & Translate("EStop") : Timer.Seconds = 1
          If Timer.Finished Then
            State = EState.HoldForTime
            TimerHold.Restart()
            .AirpadOn = True
          End If
          Status = commandName_ & Translate("Paused")

        Case EState.HoldForTime
          If TimerHold.Finished Then
            Timer.Seconds = 2
            State = EState.Done
          End If
          If TimerHold.Paused Then Timer.Restart()
          Status = commandName_ & Translate("Holding") & TimerHold.ToString(1)

        Case EState.Done
          If Timer.Finished Then
            Cancel()
          End If
          Status = commandName_ & "Completing" & Timer.ToString(1)

      End Select

    End With
  End Function

  Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
    Status = ""
    Timer.Cancel()
    TimerHold.Cancel()
    TimerAdvance.Cancel()
    TimerOverrun.Seconds = 0
  End Sub

  ReadOnly Property IsOn As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property

  ReadOnly Property IsOverrun As Boolean
    Get
      Return IsOn AndAlso TimerOverrun.Finished
    End Get
  End Property


#If 0 Then


Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Parameters=|0-99|mins\r\nName=Run\r\nMinutes='1\r\nHelp=Holds for specified time. If the main pump is not running the time is paused."
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  HoldTime = Param(1) * 60
  Timer = HoldTime
  
  With ControlCode
    .AC.ACCommand_Cancel
    .AT.ACCommand_Cancel: .DR.ACCommand_Cancel: .FI.ACCommand_Cancel
    .HD.ACCommand_Cancel: .HC.ACCommand_Cancel: .LD.ACCommand_Cancel
    .PH.ACCommand_Cancel: .RC.ACCommand_Cancel: .RH.ACCommand_Cancel
    If .RF.FillType = 86 Then .RF.ACCommand_Cancel:
    .RI.ACCommand_Cancel: .RT.ACCommand_Cancel: .RW.ACCommand_Cancel
    .SA.ACCommand_Cancel: .UL.ACCommand_Cancel: .WK.ACCommand_Cancel
    .WT.ACCommand_Cancel
    
  End With
  OverrunTimer = HoldTime + 60
  State = TMOn
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode
    
    Select Case State
    
      Case TMOff
        StateString = ""
        
      Case TMPause
        If Timer.Finished Then State = TMOff
        If Not .IO_MainPumpRunning Then
          StateString = "TM: Pump Off "
        ElseIf .Parent.IsPaused Then
          StateString = "TM: System  Paused "
        ElseIf .IO_EStop_PB Then
          StateString = "TM: Estop Active "
        Else
          State = TMOn
          Timer.Restart
        End If
      
      Case TMOn
        StateString = "TM: Holding " & TimerString(Timer.TimeRemaining)
        If (Not .IO_MainPumpRunning) Or .Parent.IsPaused Or .IO_EStop_PB Then     ' Pause timer if pump stopped
          State = TMPause
          Timer.Pause
        End If
        If Timer.Finished Then State = TMOff
        
    End Select
  End With
End Sub
Friend Sub ACCommand_Cancel()
  State = TMOff
End Sub
Friend Property Get IsOn() As Boolean
  IsOn = (State <> TMOff)
End Property
Friend Property Get IsOverrun() As Boolean
  IsOverrun = IsOn And OverrunTimer.Finished
End Property
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property



#End If
End Class
