'American & Efird - Mt Holly Then Multiflex Package
' Version 2024-07-31

Imports Utilities.Translations

<Command("Drain", "", "", "", "('StandardTimeDrain)=5", CommandType.Standard),
TranslateCommand("es", "Drenaje", ""), Description("Drain the machine  to the specified level."),
TranslateDescription("es", "Vaciar el vacío de la máquina."),
Category("Machine Functions"), TranslateCategory("es", "Machine Functions")>
Public Class DR : Inherits MarshalByRefObject : Implements ACCommand
  Private ReadOnly controlCode As ControlCode
  Private Const commandName_ As String = ("DR: ")

  'Command states
  Public Enum EState
    Off
    Interlock
    Depressurize
    DrainEmpty1
    LevelFlush
    DrainEmpty2
    Complete
  End Enum
  Public State As EState
  Public Status As String
  Public FlushCount As Integer

  Public Timer As New Timer
  Public TimerAlarm As New Timer
  Public TimerBeforeFlush As New Timer
  Public TimerOverrun As New Timer

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

      'Cancel Commands
      .AC.Cancel() : .AT.Cancel()
      .WK.Cancel()

      '.DR.Cancel() : 
      .FI.Cancel() : .HC.Cancel() : .HD.Cancel()
      .RC.Cancel() : .RH.Cancel() : .RI.Cancel() : .TM.Cancel()
      .LD.Cancel() : .PH.Cancel() : .SA.Cancel() : .UL.Cancel()
      If .RF.FillType = EFillType.Vessel Then .RF.Cancel()
      If .RT.IsForeground Then .RT.Cancel()
      ' .RW.Cancel TODO?
      .WT.Cancel()

      FlushCount = 0

      'Set the overrun to the higher value
      TimerOverrun.Minutes = .Parameters.StandardTimeDrain
      TimerAlarm.Minutes = .Parameters.DrainAlarmTime

      ' Set flow direction
      .PumpControl.SwitchInToOut()

      'Set default state
      State = EState.Interlock
      Timer.Seconds = 5

      'This is a foreground command so do not step on
      Return False
    End With
  End Function

  Function Run() As Boolean Implements ACCommand.Run
    With controlCode


      Select Case State
        Case EState.Off
          Status = (" ")


        Case EState.Interlock
          Status = commandName_ & Translate("Interlocked") & Timer.ToString(1)
          If Not .TempSafe Then Status = commandName_ & Translate("Temp Not Safe") : Timer.Seconds = 2
          ' If Not .PressSafe Then Status = commandName_ & Translate("Press Not Safe") : Timer.Seconds = 2
          If .Parent.IsPaused Then Status = commandName_ & Translate("Paused") : Timer.Seconds = 2
          If .EStop Then Status = commandName_ & Translate("EStop") : Timer.Seconds = 2
          If Timer.Finished Then
            .CO.Cancel() : .HE.Cancel() : .TP.Cancel() ' TODO .TC.Cancel
            .TemperatureControl.Cancel()
            .AirpadOn = False
            .PR.Cancel()
            .GetRecordedLevel = False
            State = EState.Depressurize
          End If


        Case EState.Depressurize
          If .SafetyControl.IsDePressurizing Then Status = .SafetyControl.StateString
          If Not .PressSafe Then Timer.Seconds = 5
          If Timer.Finished Then
            .GetRecordedLevel = False
            .PumpControl.StopMainPump()
            State = EState.DrainEmpty1
            Timer.Seconds = MinMax(.Parameters.DrainMachineTime, 30, 300)
            TimerBeforeFlush.Seconds = MinMax(.Parameters.DrainMachineFlushTime, 60, 600)
          End If
          If Not .TempSafe Then State = EState.Interlock
          Status = Translate("Depressurizing") & Timer.ToString(1)


        Case EState.DrainEmpty1
          If .MachineLevel > 10 Then
            Timer.Seconds = MinMax(.Parameters.DrainMachineTime, 30, 300)
          End If
          .GotRecordedLevel = False
          .Alarms.LosingLevelInTheVessel = False
          ' Continue draining until delay timer has completed
          If Timer.Finished Then
            Timer.Seconds = MinMax(.Parameters.LevelGaugeFlushTime, 10, 60)
            State = EState.LevelFlush
          End If
          ' Waiting too long to drain empty - may be issue with DP level transmitter
          If Not .PressSafe Then TimerBeforeFlush.Seconds = .Parameters.LevelGaugeFlushTime + 1
          If TimerBeforeFlush.Finished AndAlso (.Parameters.DrainMachineFlushTime > 0) Then
            Timer.Seconds = MinMax(.Parameters.LevelGaugeFlushTime, 10, 60)
            State = EState.LevelFlush
          End If
          ' Lost machine safe
          If Not .TempSafe Then State = EState.Interlock
          If .MachineLevel > 10 Then
            Status = Translate("Draining") & (" ") & (.MachineLevel / 10).ToString("#0.0") & "%"
          Else
            Status = Translate("Draining") & Timer.ToString(1)
          End If


        Case EState.LevelFlush
          If Timer.Finished Then
            FlushCount += 1
            State = EState.DrainEmpty2
            Timer.Seconds = MinMax(.Parameters.LevelGaugeSettleTime, 10, 120)
            TimerBeforeFlush.Seconds = MinMax(.Parameters.DrainMachineFlushTime, 60, 600)
          End If
          ' Lost Machine safe
          If Not .TempSafe Then State = EState.Interlock
          Status = Translate("Flush Level") & Timer.ToString(1)


        Case EState.DrainEmpty2
          If Timer.Finished Then  ' Wait for Level to settle first
            If .MachineLevel > 50 Then Timer.Seconds = MinMax(.Parameters.DrainMachineTime, 10, 120)
            ' Continue draining until delay timer has completed
            If Timer.Finished Then
              State = EState.Complete
              Timer.Seconds = 5
            Else
              ' Waiting too long to drain empty - may be issue with DP level transmitter
              If TimerBeforeFlush.Finished AndAlso (.Parameters.DrainMachineFlushTime > 0) Then
                Timer.Seconds = MinMax(.Parameters.LevelGaugeFlushTime, 10, 60)
                State = EState.LevelFlush
              End If
            End If
          End If
          ' Lost machine safe
          If Not .TempSafe Then State = EState.Interlock
          If .MachineLevel > 50 Then
            Status = Translate("Draining") & (" 2 ") & (.MachineLevel / 10).ToString("#0.0") & "%"
          Else
            Status = Translate("Draining") & Timer.ToString(1)
          End If

        Case EState.Complete
          If Timer.Finished Then
            .HD.HDCompleted = False
            Status = ""
            State = EState.Off
          End If
          Status = Translate("Completing") & Timer.ToString(1)

      End Select
    End With
  End Function

  Public Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
  End Sub

  ReadOnly Property IsOn As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property

  ReadOnly Property IsActive As Boolean
    Get
      Return (State > EState.Interlock) AndAlso (State <= EState.Complete)
    End Get
  End Property

  ReadOnly Property IsAlarm As Boolean
    Get
      Return IsOn AndAlso TimerAlarm.Finished AndAlso (controlCode.Parameters.DrainAlarmTime > 0)
    End Get
  End Property

  ReadOnly Property IsOverrun As Boolean
    Get
      Return IsOn AndAlso TimerOverrun.Finished AndAlso (controlCode.Parameters.StandardTimeDrain > 0)
    End Get
  End Property

  ReadOnly Property IoLevelFlush As Boolean
    Get
      Return (State = EState.LevelFlush)
    End Get
  End Property

  ReadOnly Property IoDrain As Boolean
    Get
      Return (State >= EState.DrainEmpty1) AndAlso (State <= EState.DrainEmpty2)
    End Get
  End Property

  ReadOnly Property IoTopWash As Boolean
    Get
      Return (State >= EState.Depressurize) AndAlso (State <= EState.DrainEmpty2)
    End Get
  End Property

#If 0 Then
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Drain command
Option Explicit
Implements ACCommand

Public Enum DRState
  DROff
  DRInterlock
  DRDrainEmpty
  DRLevelFlush
  DRLevelSettle
End Enum
Public State As DRState
Attribute State.VB_VarUserMemId = 0
Public StateString As String
Public LevelGaugeTimer As New acTimer
Public LevelGaugeFlushCount As Integer
Public Timer As New acTimer, OverrunTimer As New acTimer

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Name=Drain\r\nMinutes=3\r\nHelp=Drains the machine for parameter time DR Drain Time (s) after the expansion tank level float input is lost."
  Dim ControlCode As ControlCode: Set ControlCode = ControlObject
  With ControlCode
    .AC.ACCommand_Cancel: .AT.ACCommand_Cancel:
    .FI.ACCommand_Cancel: .HD.ACCommand_Cancel
    .HC.ACCommand_Cancel: .LD.ACCommand_Cancel: .PH.ACCommand_Cancel
    .PR.ACCommand_Cancel: .RC.ACCommand_Cancel: .RH.ACCommand_Cancel
    If .RF.FillType = 86 Then .RF.ACCommand_Cancel:
    .RI.ACCommand_Cancel: .RT.ACCommand_Cancel: .RW.ACCommand_Cancel
    .SA.ACCommand_Cancel: .TM.ACCommand_Cancel: .UL.ACCommand_Cancel
    .WK.ACCommand_Cancel: .WT.ACCommand_Cancel
    .AirpadOn = False
    LevelGaugeFlushCount = 0
    State = DRInterlock
    OverrunTimer = 420
  End With
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode
        
    Select Case State
    
      Case DROff
        StateString = ""
        
      Case DRInterlock
        If .TempSafe And Not .PressSafe Then
          StateString = "DR: " & .SafetyControl.StateString
        Else
          StateString = "DR: Unsafe to Drain "
        End If
        If .MachineSafe Then
           If .RF.IsFillWithVessel Then .RF.ACCommand_Cancel
           .TP.ACCommand_Cancel
           .CO.ACCommand_Cancel
           .HE.ACCommand_Cancel
           .TemperatureControl.Cancel
           .TC.ACCommand_Cancel
           .TemperatureControlContacts.Cancel
           .PumpRequest = True
           .PumpOnCount = 0
           State = DRDrainEmpty
           Timer = .Parameters_DRDrainTime
        End If
                
      Case DRDrainEmpty
        If (.VesselLevel > 10) Then
          StateString = "DR: Draining " & Pad(.VesselLevel / 10, "0", 3) & "%"
          Timer = .Parameters_DRDrainTime
        Else: StateString = "DR: Draining " & TimerString(Timer.TimeRemaining)
        End If
        .GetRecordedLevel = False
        .Alarms_LosingLevelInTheVessel = False
        If Timer.Finished Then
           Timer = .Parameters_LevelGaugeFlushTime
           .PumpRequest = False
           State = DRLevelFlush
        End If
        ' Level gauge flush override timer
        If Not .PressSafe Then LevelGaugeTimer = .Parameters_LevelFlushOverrideTime + 1
        If (.Parameters_LevelFlushOverrideTime > 0) And LevelGaugeTimer.Finished Then
          State = HDLevelFlush
          Timer = .Parameters_LevelGaugeFlushTime
          LevelGaugeFlushCount = LevelGaugeFlushCount + 1
        End If
      
      Case DRLevelFlush
        StateString = "DR: Flushing Level Gauge " & Pad(LevelGaugeFlushCount, "0", 1) & TimerString(Timer.TimeRemaining)
        If Timer.Finished Then
          Timer = .Parameters_LevelGaugeSettleTime
          State = DRLevelSettle
        End If
        
      Case DRLevelSettle
        StateString = "Settling Level Gauge " & Pad(LevelGaugeFlushCount, "0", 1) & " " & TimerString(Timer.TimeRemaining)
        If Timer.Finished Then
          ' Check empty with ability to auto-flush level transmitter for false level
          If (.VesselLevel > 10) Then
           State = DRDrainEmpty
           Timer = .Parameters_DRDrainTime
          Else
            State = DROff
          End If
        End If

    End Select
  End With
End Sub
Friend Sub ACCommand_Cancel()
  State = DROff
End Sub
Friend Property Get IsOn() As Boolean
  IsOn = (State <> DROff)
End Property
Friend Property Get IsOverrun() As Boolean
  IsOverrun = IsOn And OverrunTimer.Finished
End Property
Friend Property Get IsActive() As Boolean
  If IsDraining Then IsActive = True
End Property
Friend Property Get IsDraining() As Boolean
  If (State = DRDrainEmpty) Then IsDraining = True
End Property
Friend Property Get IsInterlocked() As Boolean
  If (State = DRInterlock) Then IsInterlocked = True
End Property

Friend Property Get IoFlushingLevel() As Boolean
  If (State = DRLevelFlush) Then IoFlushingLevel = True
End Property
Friend Property Get IoDrain() As Boolean
  If (State = DRDrainEmpty) Or (State = DRLevelFlush) Or (State = DRLevelSettle) Then IoDrain = True
End Property

Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property

#End If
End Class
