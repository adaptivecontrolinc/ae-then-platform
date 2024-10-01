'American & Efird - Mt. Holly Then Platform to GMX
' Version 2024-07-22

Imports Utilities.Translations

<Command("Hold for", "|0-99| contacts", " ", "", "'1 * 1"), _
TranslateCommand("es", "Mantenga pulsado durante", "|0-99| contactos"), _
Description("Holds for specified number of contacts. If interlocks are lost, the count is paused."), _
TranslateDescription("es", "Bodegas para determinado número de contactos. Si los bloqueos se pierden, la cuenta está en pausa."), _
Category("Machine Functions")> _
Public Class HC : Inherits MarshalByRefObject : Implements ACCommand
  Private Const commandName_ As String = ("HC: ")

  'Command States
  Public Enum EState
    Off
    Interlock
    CheckMachineClosed
    StartPump
    Pressurize
    Paused
    HoldForTime
  End Enum
  Property State As EState
  Property Status As String
  Property HoldTurnsSv As Integer
  Property HoldTurnsPv As Integer
  Property HoldTurnsVolume As Integer
  Property MachineNotClosed As Boolean

  Property TimeHold As Integer
  Property TimerHold As New Timer
  Property TimerOverrun As New Timer
  Property TimerAdvance As New Timer

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
    If controlCode.Parameters.EnableCommandChg = 1 Then Start(param)
  End Sub

  Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With controlCode

      ' Cancel Commands
      .AC.Cancel() : .AT.Cancel()
      .KA.Cancel() : .WK.Cancel()
      .DR.Cancel() : .FI.Cancel() ': .HC.Cancel() :
      .HD.Cancel()
      .RI.Cancel() : .RC.Cancel() : .RH.Cancel() : .TM.Cancel()
      .LD.Cancel() : .PH.Cancel() : .SA.Cancel() : .UL.Cancel()
      If .RT.IsForeground Then .RT.Cancel()
      '  .CO.Cancel() : .HC.Cancel() : .HE.Cancel() : 
      .WT.Cancel()

      ' Check array bounds just to be on the safe side
      HoldTurnsSv = 0
      If param.GetUpperBound(0) >= 1 Then HoldTurnsSv = param(1)

      ' Set timer
      TimerHold.Milliseconds = 1000
      TimerOverrun.Minutes = (HoldTurnsSv * 3) + 1
      HoldTurnsPv = 0

      'Set default state
      State = EState.Interlock
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
        If Not (.MachineClosed) Then MachineNotClosed = True
        State = EState.Interlock
        TimerHold.Pause()
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
          If Not .MachineClosed Then Status = Translate("Machine Not Closed") : TimerHold.Seconds = 10
          If Not .IO.PumpRunning Then Status = Translate("Pump Off") : TimerHold.Seconds = 1
          If .TemperatureControl.IsCrashCoolOn Then Status = Translate("Crash-Cooling") : TimerHold.Seconds = 1
          If .Parent.IsPaused Then Status = Translate("Paused") : TimerHold.Seconds = 1
          If .EStop Then Status = Translate("EStop") : TimerHold.Seconds = 1
          If Not (.MachineClosed) Then MachineNotClosed = True
          If TimerHold.Finished Then
            State = EState.CheckMachineClosed
          End If

        Case EState.CheckMachineClosed
          Status = Translate("Checking Lids") & (" ") & TimerHold.ToString
          If Not .MachineClosed Then Status = Translate("Machine Not Closed") : TimerHold.Seconds = 10
          If Not (.MachineClosed) Then MachineNotClosed = True
          If Not .AdvancePb Then
            TimerAdvance.Seconds = 2
            Status = Translate("Hold Advance") & (" ") & TimerAdvance.ToString
          Else
            'Attempting to Advance
            If .Parent.Signal <> "" Then .Parent.Signal = ""
            Status = Translate("Advancing") & (" ") & TimerAdvance.ToString
            If TimerAdvance.Finished Then MachineNotClosed = False
          End If
          If Not .MachineClosed Then Status = Translate("Machine Not Closed") : TimerHold.Seconds = 10
          If Not (.MachineClosed) Then MachineNotClosed = True
          If TimerHold.Finished AndAlso (Not MachineNotClosed) Then
            ' Machine is closed, all is ok
            State = EState.StartPump
            ' Start pump if not running
            If Not .PumpControl.IsActive Then
              .PumpControl.StartAuto()
              TimerHold.Seconds = .Parameters.FillPrimePumpTime
            End If
          End If

        Case EState.StartPump
          Status = .PumpControl.StateString
          If Not .PumpControl.IsRunning Then TimerHold.Seconds = .PumpControl.Parameters_PumpAccelerationTime
          If TimerHold.Finished Then
            State = EState.Pressurize
            TimerHold.Seconds = 5
          End If

        Case EState.Pressurize
          Status = Translate("Pressurizing") & (" ") & TimerHold.ToString
          If Not .SafetyControl.IsPressurized Then TimerHold.Seconds = 5
          If Not .AirpadOn Then .AirpadOn = True
          If TimerHold.Finished Then
            State = EState.HoldForTime
            TimerHold.Seconds = TimeHold
          End If

          If Not .PumpControl.PumpRunning Then Status = Translate("Pump Off")
          If .TemperatureControl.IsCrashCoolOn Then Status = Translate("Crash-Cooling") : TimerHold.Seconds = 1
          If .Parent.IsPaused Then Status = Translate("Paused") : TimerHold.Seconds = 1
          If .EStop Then Status = Translate("EStop") : TimerHold.Seconds = 1
          If TimerHold.Finished Then
            State = EState.HoldForTime
            TimerHold.Milliseconds = 1000
            If Not .AirpadOn Then .AirpadOn = True
          End If

        Case EState.HoldForTime
          Status = Translate("Holding") & (" ") & (HoldTurnsSv - HoldTurnsPv).ToString("#0") & " / " & (HoldTurnsSv).ToString("#0")
          ' Calculate Batch Turnover
          If TimerHold.Finished Then
            HoldTurnsVolume += CInt(.FlowRate / 60)
            If HoldTurnsVolume >= (.VolumeBasedOnLevel) Then
              HoldTurnsVolume -= .VolumeBasedOnLevel
              HoldTurnsPv += 1
            End If
            TimerHold.Milliseconds = 1000
          End If
          ' Pause Reasons
          If (Not .IO.PumpRunning) OrElse .Parent.IsPaused OrElse .EStop Then
            State = EState.Paused
          End If
          ' Reached desired turns
          If HoldTurnsPv > +HoldTurnsSv Then State = EState.Off

      End Select

    End With
  End Function

  Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
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



  ' TODO
#If 0 Then


  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Time command
Option Explicit
Implements ACCommand
Public Enum HCState
  HCOff
  HCPause
  HCOn
End Enum
Public State As HCState
Public StateString As String
Public OneSecondTimer As New acTimer, OverrunTimer As New acTimer
Public NumberOfHoldingTurns As Long
Public NumberOfTurns As Long
Public HCVolume As Double

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Parameters=|0-99| contacts\r\nName=Hold for \r\nMinutes=('1*3)\r\nHelp=Holds for specified number of contacts. "
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  NumberOfHoldingTurns = Param(1)
  NumberOfTurns = 0
  HCVolume = 0
  With ControlCode
    .AC.ACCommand_Cancel
    .AT.ACCommand_Cancel: .DR.ACCommand_Cancel: .FI.ACCommand_Cancel
    .HD.ACCommand_Cancel: .LD.ACCommand_Cancel: .PH.ACCommand_Cancel
    .RC.ACCommand_Cancel: .RH.ACCommand_Cancel: .RI.ACCommand_Cancel
    If .RF.FillType = 86 Then .RF.ACCommand_Cancel:
    .RT.ACCommand_Cancel: .RW.ACCommand_Cancel: .SA.ACCommand_Cancel
    .TM.ACCommand_Cancel: .UL.ACCommand_Cancel: .WK.ACCommand_Cancel
    .WT.ACCommand_Cancel
    
  End With
  OverrunTimer = ((NumberOfHoldingTurns * 3) * 60) + 60
  OneSecondTimer = 1
  State = HCOn
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode
    
    Select Case State
    
      Case HCOff
        StateString = ""
        
      Case HCPause
        If Not .IO_MainPumpRunning Then
          StateString = "HC Paused: Pump Off "
        ElseIf .Parent.IsPaused Then
          StateString = "HC Paused: "
        ElseIf .IO_EStop_PB Then
          StateString = "HC Paused: EStop "
        Else: State = HCOn
        End If
      
      Case HCOn
        StateString = "HC: Turns Remaining " & (NumberOfHoldingTurns - NumberOfTurns) & " / " & NumberOfHoldingTurns
        If OneSecondTimer.Finished Then
          HCVolume = HCVolume + (.FlowRate / 60)
          If HCVolume >= (.VolumeBasedOnLevel) Then  ' volume based on level in gallons
             HCVolume = HCVolume - (.VolumeBasedOnLevel)
             NumberOfTurns = NumberOfTurns + 1
          End If
          OneSecondTimer = 1
        End If
        If NumberOfTurns >= NumberOfHoldingTurns Then State = HCOff
        If (Not .IO_MainPumpRunning) Or .Parent.IsPaused Or .IO_EStop_PB Then     ' Pause timer if pump stopped
          State = HCPause
        End If
         
    End Select
  End With
End Sub

Friend Sub ACCommand_Cancel()
  State = HCOff
  NumberOfTurns = 0
  HCVolume = 0
End Sub
Friend Property Get IsOn() As Boolean
  IsOn = (State <> HCOff)
End Property
Friend Property Get IsOverrun() As Boolean
  IsOverrun = IsOn And OverrunTimer.Finished
End Property
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property






#End If
End Class
