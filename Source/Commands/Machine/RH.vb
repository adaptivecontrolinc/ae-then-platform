'American & Efird - Then Multiflex Package
' Version 2024-09-11

Imports Utilities.Translations

<Command("Rinse with Heat", "|0-180|F for |0-99|mins", "", "", "'2 + 2"),
TranslateCommand("es", "Enjuague con calor", "|0-180|F for |0-99|mins"),
Description("Rinses with water at the specified temperature for the specified time. Does not cancel previous temperature command."),
TranslateDescription("es", "Enjuagues con agua a la temperatura especificada durante el tiempo especificado. No anula el anterior comando de temperatura."),
Category("Machine Functions"), TranslateCategory("es", "Machine Functions")>
Public Class RH : Inherits MarshalByRefObject : Implements ACCommand
  Private Const commandName_ As String = ("RH: ")
  Private ReadOnly controlCode As ControlCode

  'Command States
  Public Enum EState
    Off
    Interlock
    Depressurize
    ResetMeter
    Pause
    FillStart
    Filling
    Rinsing
    Done
  End Enum
  Public State As EState
  Public Status As String
  Public RinseTemp As Integer
  Public RinseTime As Integer

  Public BlendControl As New BlendControl

  Public ReadOnly Timer As New Timer
  Friend ReadOnly TimerAdvance As New Timer
  Public ReadOnly Property TimeAdvance As Integer
    Get
      Return TimerAdvance.Seconds
    End Get
  End Property
  Friend ReadOnly TimerOverrun As New Timer
  Public ReadOnly Property TimeOverrun As Integer
    Get
      Return TimerOverrun.Seconds
    End Get
  End Property
  Private ReadOnly timerRinseFlush_ As New Timer
  Friend ReadOnly TimerTopWashBlocked As New Timer
  Public ReadOnly Property TimeTopWashBlocked As Integer
    Get
      Return TimerTopWashBlocked.Seconds
    End Get
  End Property

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

  Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With controlCode

      ' Cancel Commands
      .AC.Cancel() : .AT.Cancel()
      .KA.Cancel() : .WK.Cancel()
      .DR.Cancel() : .FI.Cancel() : .HC.Cancel() : .HD.Cancel() : .PR.Cancel()
      .RI.Cancel() : .RC.Cancel() ': .RH.Cancel() : 
      .TM.Cancel()
      .LD.Cancel() : .PH.Cancel() : .SA.Cancel() : .UL.Cancel()
      If .RF.IsForeground AndAlso .RF.FillType = EFillType.Vessel Then .RF.Cancel() ' TODO CHeck
      If .RT.IsForeground Then .RT.Cancel()
      .WT.Cancel()

      ' Cancel PressureOn command if active during Atmospheric Commands
      .PR.Cancel()

      ' Check Parameters
      If param.GetUpperBound(0) >= 1 Then RinseTemp = param(1) * 10
      If param.GetUpperBound(0) >= 2 Then RinseTime = param(2) * 60

      ' Confirm limits and update RinseType
      If RinseTemp > 1800 Then RinseTemp = 1800

      TimerOverrun.Seconds = RinseTime + 120
      State = EState.Interlock

      'This is a foreground command so do not step on
      Return False
    End With
  End Function

  Function Run() As Boolean Implements ACCommand.Run
    With controlCode
      ' Only run the state machine if we are active
      If State = EState.Off Then Exit Function

      ' Set blend control parameters
      BlendControl.Parameters(.Parameters.BlendDeadBand, .Parameters.BlendFactor, .Parameters.BlendSettleTime, .Parameters.ColdWaterTemperature, .Parameters.HotWaterTemperature)

      ' Safe to rinse
      Dim safe As Boolean = (.TempSafe OrElse .HD.HDCompleted) AndAlso .MachineClosed AndAlso (Not .TemperatureControl.IsCrashCoolOn)
      Dim pauseCommand As Boolean = .EStop OrElse .Parent.IsPaused OrElse (Not .PumpControl.IsRunning)

      ' Select State
      Select Case State

        Case EState.Off
          Status = (" ")


        Case EState.Interlock
          Status = commandName_ & Translate("Interlocked") & Timer.ToString(1)
          If Not safe Then
            If Not (.TempSafe OrElse .HD.HDCompleted) Then Status = commandName_ & Translate("Temp Not Safe")
            If Not .MachineClosed Then Status = commandName_ & Translate("Machine Not Closed")
            If Not .PressSafe Then Status = commandName_ & Translate("Press Not Safe")
            If .TemperatureControl.IsCrashCoolOn Then Status = commandName_ & Translate("Crash-Cooling")
            Timer.Seconds = 1
          End If
          If .Parent.IsPaused OrElse .EStop Then
            ' If Not .IO.PumpRunning Then Status = Translate("Pump Off")
            If .Parent.IsPaused Then Status = commandName_ & Translate("Paused")
            If .EStop Then Status = commandName_ & Translate("EStop")
            Timer.Seconds = 1
          End If
          If Timer.Finished Then
            ' Cancel Temp Commands
            '  .CO.Cancel() : .HE.Cancel() : .TP.Cancel() ': .TC.Cancel()
            '  .TemperatureControl.Cancel()
            .GetRecordedLevel = False
            .AirpadOn = False
            State = EState.Depressurize
            Timer.Seconds = 5
          End If


        Case EState.Depressurize
          If Not .PressSafe Then Timer.Seconds = 5
          If Timer.Finished Then
            ' Step to next state
            State = EState.ResetMeter
          End If
          If Not safe Then State = EState.Interlock
          Status = Translate("Depressuzing") & Timer.ToString(1)


        Case EState.ResetMeter
          'If .MachineVolume = 0 Then
          '  State = EState.FillStart
          '  Timer.Seconds = MinMax(.Parameters.TopWashOpenDelay, 5, 60)
          'End If
          State = EState.FillStart
          Timer.Seconds = MinMax(.Parameters.TopWashOpenDelay, 5, 60)
          If Not safe Then State = EState.Interlock
          Status = commandName_ & Translate("Reset flowmeter")


        Case EState.Pause
          Status = commandName_ & Translate("Paused") & Timer.ToString(1)
          If Not safe Then
            If Not (.TempSafe OrElse .HD.HDCompleted) Then Status = commandName_ & Translate("Temp Not Safe")
            If Not .MachineClosed Then Status = commandName_ & Translate("Machine Not Closed")
            If .TemperatureControl.IsCrashCoolOn Then Status = commandName_ & Translate("Crash-Cooling")
            Timer.Seconds = 1
          End If
          If pauseCommand Then
            If Not .IO.PumpRunning Then Status = commandName_ & Translate("Pump Off")
            If .Parent.IsPaused Then Status = commandName_ & Translate("Paused")
            If .EStop Then Status = commandName_ & Translate("EStop")
            Timer.Seconds = 1
          End If
          If Timer.Finished Then
            ' Reset all timers
            TimerTopWashBlocked.Seconds = .Parameters.TopWashBlockageTimer
            ' Restart Filling
            State = EState.Filling
            Timer.Seconds = MinMax(.Parameters.TopWashOpenDelay, 5, 60)
          End If
          ' Reset timers with advance pushbutton
          If Not .AdvancePb Then TimerAdvance.Seconds = 2
          If TimerAdvance.Finished Then
            TimerTopWashBlocked.Seconds = .Parameters.TopWashBlockageTimer + 1
          End If
          If Not .TempSafe Then State = EState.Interlock


        Case EState.FillStart
          If Timer.Finished Then
            BlendControl.Start(RinseTemp)
            ' Reset all timers
            TimerTopWashBlocked.Seconds = .Parameters.TopWashBlockageTimer
            ' Start Filling - start the rinse timer once the machine has reached 100%
            State = EState.Filling
          End If
          Status = commandName_ & Translate("Starting") & Timer.ToString(1)


        Case EState.Filling
          BlendControl.Run(.IO.FillTemp)
          If .MachineLevel >= 850 Then
            State = EState.Rinsing
            ' Restart the main pump
            If Not .PumpControl.IsActive Then .PumpControl.StartAuto()
            ' Reset all timers
            TimerTopWashBlocked.Seconds = .Parameters.TopWashBlockageTimer + 1
            Timer.Seconds = RinseTime
          End If
          If Not (.TempSafe AndAlso .PressSafe) Then State = EState.Interlock
          Status = commandName_ & Translate("Filling") & (" ") & (.MachineLevel / 10).ToString("#0.0") & " / 85.0 %"


        Case EState.Rinsing
          BlendControl.Run(.IO.FillTemp)
          ' Top Wash Blockage
          If IsTopWashBlocked Then Status = commandName_ & Translate("Top Wash Blocked, Hold Run to Reset") & (" ") & TimerAdvance.ToString
          ' Reset blockage timer [plastic blocks topwash valve]
          If .IO.PressureInterlock OrElse (.Parameters.TopWashBlockageTimer = 0) Then TimerTopWashBlocked.Seconds = .Parameters.TopWashBlockageTimer + 1
          ' Pause if necessary
          If pauseCommand Then
            RinseTime = Timer.Seconds
            Timer.Seconds = 5
            State = EState.Pause
          End If
          ' Interlock if necessary
          If Not (safe) Then
            RinseTime = Timer.Seconds
            Timer.Seconds = 5
            State = EState.Interlock
          End If
          ' Reset timers with advance pushbutton
          If Not .AdvancePb Then TimerAdvance.Seconds = 2
          If TimerAdvance.Finished Then
            TimerTopWashBlocked.Seconds = .Parameters.TopWashBlockageTimer + 1
          End If
          ' Rinse Completed
          If Timer.Finished Then
            Timer.Seconds = 1
            State = EState.Done
          End If
          Status = commandName_ & Translate("Rinsing") & Timer.ToString(1)


        Case EState.Done
          If Timer.Finished Then
            .HD.HDCompleted = False
            State = EState.Off
          End If
          Status = commandName_ & Translate("Completing")

      End Select
    End With
  End Function

  Public Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
    TimerOverrun.Cancel()
    Timer.Cancel()
  End Sub

  Public ReadOnly Property IsOn() As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property

  Public ReadOnly Property IsActive As Boolean
    Get
      Return (State >= EState.Depressurize) AndAlso (State <= EState.Rinsing)
    End Get
  End Property

  ReadOnly Property IsResetMeter As Boolean
    Get
      Return State = EState.ResetMeter
    End Get
  End Property

  ReadOnly Property IsOverrun As Boolean
    Get
      Return IsOn AndAlso TimerOverrun.Finished
    End Get
  End Property

  ReadOnly Property IsTopWashBlocked As Boolean
    Get
      Return (State = EState.Rinsing) AndAlso TimerTopWashBlocked.Finished
    End Get
  End Property

  ReadOnly Property IoMachineFill As Boolean
    Get
      Return (State = EState.Filling) OrElse (State = EState.Rinsing)
    End Get
  End Property

  'ReadOnly Property IoFillCold As Boolean
  '  Get
  '    Return IoMachineFill AndAlso ((RinseType = EFillType.Cold) OrElse (RinseType = EFillType.Mix))
  '  End Get
  'End Property

  'ReadOnly Property IoFillHot As Boolean
  '  Get
  '    Return IoMachineFill AndAlso ((RinseType = EFillType.Hot) OrElse (RinseType = EFillType.Mix))
  '  End Get
  'End Property

  ReadOnly Property IoTopWash As Boolean
    Get
      Return (State = EState.FillStart) OrElse (State = EState.Filling) OrElse (State = EState.Rinsing)
    End Get
  End Property

  ReadOnly Property IOBlendOutput As Integer
    Get
      If IoMachineFill Then Return BlendControl.IOOutput
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
Attribute VB_Name = "RH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Rinse Heat Command
'Same as RI but doesn't cancel temperature control
'If you want to heat with this, use a TP command before it.
Option Explicit
Implements ACCommand
Public Enum RHState
  RHOff
  RHInterLock
  RHNotSafe
  RHFillToLevel
  RHPause
  RHRinse
End Enum
Public State As RHState
Public StateString As String
Public Timer As New acTimer, OverrunTimer As New acTimer
Public AdvanceTimer As New acTimer
Public RinseTime As Long
Public RinseTemperature As Long
Public RinseFlushTimer As New acTimer
Private BlendControl As New acBlendControl
Public TemperatureHigh As Boolean
Public TemperatureLow As Boolean
Public temperaturehightimer As New acTimer
Public temperaturelowtimer As New acTimer
Public topwashblocktimer As New acTimer
Public TopWashBlocked As Boolean

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Parameters=for |0-99|mins\r\nName=Rinse And Heat\r\nMinutes='1\r\nHelp=Rinses with water for the specified time."
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  RinseTemperature = Param(1) * 10
  RinseTime = Param(2) * 60
  
  With ControlCode
    .AC.ACCommand_Cancel
    .AT.ACCommand_Cancel: .DR.ACCommand_Cancel: .FI.ACCommand_Cancel
    .HD.ACCommand_Cancel: .HC.ACCommand_Cancel: .LD.ACCommand_Cancel
    If .RF.FillType = 86 Then .RF.ACCommand_Cancel:
    .PH.ACCommand_Cancel: .RC.ACCommand_Cancel: .RI.ACCommand_Cancel
    .RT.ACCommand_Cancel: .RW.ACCommand_Cancel: .SA.ACCommand_Cancel
    .TM.ACCommand_Cancel: .UL.ACCommand_Cancel: .WK.ACCommand_Cancel
    .WT.ACCommand_Cancel
     BlendControl.Params .Parameters.BlendFactor, .Parameters.BlendDeadBand, 5, 100, _
                        .Parameters.ColdWaterTemperature, .Parameters.HotWaterTemperature
    .AirpadOn = False
  End With
 
  OverrunTimer = RinseTime + 60
  State = RHInterLock
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode
    Select Case State
    
      Case RHOff
        StateString = ""
        
      Case RHInterLock
        If .TempSafe And Not .PressSafe Then
          StateString = .SafetyControl.StateString
        Else
          StateString = "RH: Unsafe to rinse"
        End If
        If (.MachineSafe Or .HD.HDCompleted) Then
           State = RHFillToLevel
           BlendControl.Start RinseTemperature
        End If
        
      Case RHNotSafe
        If .TempSafe And Not .PressSafe Then
          StateString = .SafetyControl.StateString
        Else
          StateString = "RH: Unsafe to rinse"
        End If
        TopWashBlocked = False
        Timer.Pause
        If .TempSafe Then State = RHPause
     
      Case RHFillToLevel
        StateString = "RH: Filling " & Pad(.VesselLevel / 10, "0", 3) & "% / " & Pad(850 / 10, "0", 3) & "% "
       'Run blend control code
        BlendControl.Run .IO_BlendFillTemp
        If (.VesselLevel >= 850) Then
          temperaturelowtimer = .Parameters.RinseTemperatureAlarmDelay
          temperaturehightimer = .Parameters.RinseTemperatureAlarmDelay
          topwashblocktimer.TimeRemaining = .Parameters.TopWashBlockageTimer
          Timer = RinseTime
          State = RHRinse
          .PumpRequest = True
        End If
        
      Case RHPause
        StateString = "RH: Paused "
        Timer.Pause
        If .IO_MainPumpRunning And (Not .Parent.IsPaused) And (Not .IO_EStop_PB) Then
           Timer.Restart
           State = RHRinse
        End If
        If Not .TempSafe Then State = RHNotSafe
        
      Case RHRinse
        If TopWashBlocked Then
          StateString = "RH: Top Wash Blocked, Hold Run to Reset " & TimerString(AdvanceTimer.TimeRemaining)
        Else
          StateString = "RH: Rinsing " & TimerString(Timer.TimeRemaining)
        End If
        BlendControl.Run .IO_BlendFillTemp
        'If pump and reel are not running pause rinse
        If (Not .IO_MainPumpRunning) Or .Parent.IsPaused Or .IO_EStop_PB Then State = RHPause
        If Not .TempSafe Then State = RHNotSafe
        'temp alarms
        If (.IO_BlendFillTemp > (RinseTemperature - .Parameters.RinseTemperatureAlarmBand)) Or (RinseTemperature = 0) Or (RinseTemperature = 1800) Then
         temperaturelowtimer = .Parameters.RinseTemperatureAlarmDelay
        End If
        If (.IO_BlendFillTemp < (RinseTemperature + .Parameters.RinseTemperatureAlarmBand)) Or (RinseTemperature = 0) Or (RinseTemperature = 1800) Then
         temperaturehightimer = .Parameters.RinseTemperatureAlarmDelay
        End If
        If temperaturelowtimer.Finished Then
          TemperatureLow = True
        Else
           TemperatureLow = False
        End If
        If temperaturehightimer.Finished Then
          TemperatureHigh = True
        Else
          TemperatureHigh = False
        End If
        'top wash timer - look at pressure switch to see if plastic is blocking topwash
        If .IO_PressureInterlock Then topwashblocktimer.TimeRemaining = .Parameters.TopWashBlockageTimer
        If (.Parameters.TopWashBlockageTimer > 0) And topwashblocktimer.Finished Then
          TopWashBlocked = True
        End If
        If Not .IO_RemoteRun Then AdvanceTimer.TimeRemaining = 2
        If AdvanceTimer.Finished Then
          topwashblocktimer.TimeRemaining = .Parameters.TopWashBlockageTimer
          TopWashBlocked = False
        End If
        If Timer.Finished Then
          State = RHOff
          TemperatureHigh = False
          TemperatureLow = False
          TopWashBlocked = False
          .HD.HDCompleted = False
        End If

    End Select
  End With
End Sub
Friend Sub ACCommand_Cancel()
  State = RHOff
  TemperatureHigh = False
  TemperatureLow = False
  TopWashBlocked = False
End Sub
Friend Property Get IsOn() As Boolean
  IsOn = (State <> RHOff)
End Property
Friend Property Get IsActive() As Boolean
  If IsRinsing Or IsFilling Then IsActive = True
End Property
Friend Property Get IsInterlocked() As Boolean
  If State = RHInterLock Then IsInterlocked = True
End Property
Friend Property Get IsNotSafe() As Boolean
  If State = RHNotSafe Then IsNotSafe = True
End Property
Friend Property Get IsFilling() As Boolean
  IsFilling = (State = RHFillToLevel)
End Property
Friend Property Get IsRinsing() As Boolean
  IsRinsing = (State = RHRinse)
End Property
Friend Property Get IsPaused() As Boolean
  If (State = RHPause) Then IsPaused = True
End Property
Friend Property Get IsOverrun() As Boolean
  IsOverrun = IsOn And OverrunTimer.Finished
End Property
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property
Friend Property Get BlendOutput() As Long

'Blend output
  BlendOutput = BlendControl.Output
'Limit output
  MinMax BlendOutput, 0, 1000
End Property



#End If
End Class
