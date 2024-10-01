'American & Efird - 
' Version 2024-07-30 Mt.Holly Then Platform
' Version 2022-05-06 GMX Tina14

Imports Utilities.Translations

<Command("Fill", "At |0-180|F To |0-100|%", "", "'1", "'StandardTimeFillMachine=10"),
TranslateCommand("es", "Relleno ", "a |0-180|F To |0-100|%"),
Description("Fills the machine with water at the specified temperature to the specified level."),
TranslateDescription("es", "La máquina se llena de agua a la temperatura especificada para el valor del parámetro rellenar nivel."),
Category("Machine Functions"), TranslateCategory("es", "Machine Functions")>
Public Class FI : Inherits MarshalByRefObject : Implements ACCommand
  Private Const commandName_ As String = ("FI: ")
  Private ReadOnly controlCode As ControlCode

  'Command States
  Public Enum Estate
    Off
    Interlock
    Depressurize
    ResetMeter

    NotSafe

    FillStart
    FillToLevel
    FillSettle1
    PumpSettle
    TopUp
    FillSettle2
    Complete
  End Enum
  Public State As Estate
  Public Status As String
  Public FillLevel As Integer
  Public FillTemp As Integer
  '  Public FillType As EFillType
  Public NumberOfTopups As Integer

  Public BlendControl As New BlendControl

  Public Timer As New Timer
  Public TimerAlarm As New Timer
  Public TimerOverrun As New Timer

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
      .DR.Cancel() ': .FI.Cancel() 
      .HC.Cancel() : .HD.Cancel()
      .RI.Cancel() : .RC.Cancel() : .RH.Cancel() : .TM.Cancel()
      .LD.Cancel() : .PH.Cancel() : .SA.Cancel() : .UL.Cancel()
      If .RF.FillType = EFillType.Vessel Then .RF.Cancel() ' TODO Check
      If .RT.IsForeground Then .RT.Cancel() : .RW.Cancel()
      .WT.Cancel()

      ' Cancel PressureOn command if active during Atmospheric Commands
      .PR.Cancel()

      ' Check Parameters
      If param.GetUpperBound(0) >= 1 Then FillTemp = param(1) * 10
      If FillTemp > 1800 Then FillTemp = 1800

      ' Set Fill Type
      'FillType = EFillType.Cold
      'If FillTemp > 400 Then FillType = EFillType.Mix
      'If FillTemp > 900 Then FillType = EFillType.Hot
      'If (.Parameters.FillEnableBlend = 0) AndAlso FillType = EFillType.Mix Then FillType = EFillType.Hot
      'If (.Parameters.FillEnableHot = 0) Then FillType = EFillType.Cold


      If param.GetUpperBound(0) >= 2 Then FillLevel = param(2) * 10
      If FillLevel > 1000 Then FillLevel = 1000

      ' Use Command FillLevel or Working Level based on NP & PT:
      If (.Parameters.EnableFillByWorkingLevel = 1) And (.WorkingLevel > 0) Then
        FillLevel = .WorkingLevel
        If FillLevel > 1000 Then FillLevel = 1000
      End If

      ' Initialize OverrunTimer and First State
      TimerOverrun.Minutes = .Parameters.StandardTimeFillMachine
      TimerAlarm.Minutes = .Parameters.FillAlarmTime

      ' Reset Flags
      NumberOfTopups = 0

      State = Estate.Interlock
      Timer.Seconds = 2

      'This is a foreground command so do not step on
      Return False
    End With
  End Function

  Function Run() As Boolean Implements ACCommand.Run
    With controlCode

      ' Only run the state machine if we are active
      If State = Estate.Off Then Exit Function

      ' Set blend control parameters
      BlendControl.Parameters(.Parameters.BlendDeadBand, .Parameters.BlendFactor, .Parameters.BlendSettleTime, .Parameters.ColdWaterTemperature, .Parameters.HotWaterTemperature)

      ' Safe-To-Fill conditions
      Dim safe As Boolean = (.TempSafe OrElse .HD.HDCompleted) AndAlso .PressSafe AndAlso .MachineClosed

      ' Issue with inconsistent state value in histories - 
      Static StartState As Estate
      Do
        ' Remember state and loop until state does not change
        ' This makes sure we go through all the state changes before proceeding and setting IO
        StartState = State

        ' Force state machine to interlock state if the machine is not safe
        If IsActive AndAlso Not safe Then State = Estate.Interlock

        ' STATE LOGIC ************************************************************************
        Select Case State

          Case Estate.Off
            Status = (" ")

          Case Estate.Interlock
            Status = commandName_ & Translate("Interlocked")
            If Not safe Then
              If Not (.TempSafe OrElse .HD.HDCompleted) Then Status = commandName_ & Translate("Temp Not Safe")
              If Not .MachineClosed Then Status = commandName_ & Translate("Machine Not Closed")
              If .Parent.IsPaused Then Status = commandName_ & Translate("Paused")
              If .EStop Then Status = commandName_ & Translate("EStop")
              Timer.Seconds = 1
            End If
            If Timer.Finished Then
              .CO.Cancel() : .HE.Cancel() : .TP.Cancel()
              .TemperatureControl.Cancel()
              .AirpadOn = False
              .GetRecordedLevel = False
              State = Estate.Depressurize
            End If

          Case Estate.Depressurize
            If Not .PressSafe Then Timer.Seconds = 5
            If Timer.Finished Then
              ' Step to next state
              State = Estate.ResetMeter
            End If
            Status = commandName_ & Translate("Depressuzing") & Timer.ToString(1)

          Case Estate.ResetMeter
            'If .MachineVolume = 0 Then
            If .MachineLevel >= FillLevel Then
                State = Estate.PumpSettle
                Timer.Seconds = .Parameters.FillPrimePumpTime
                If Not .PumpControl.IsActive Then .PumpControl.StartAuto()
              Else
                .PumpControl.StopMainPump()
                State = Estate.FillStart
                Timer.Seconds = MinMax(.Parameters.FillOpenDelayTime, 5, 60)
              End If
            'End If
            Status = commandName_ & Translate("Reset flowmeter")

          Case Estate.NotSafe
            Status = commandName_ & Translate("Interlocked")
            If Not (safe AndAlso .PressSafe) Then
              If Not .PressSafe Then Status = commandName_ & Translate("Press Not Safe")
              If Not (.TempSafe OrElse .HD.HDCompleted) Then Status = commandName_ & Translate("Temp Not Safe")
              If Not .MachineClosed Then Status = commandName_ & Translate("Machine Not Closed")
              If .Parent.IsPaused Then Status = commandName_ & Translate("Paused")
              If .EStop Then Status = commandName_ & Translate("EStop")
              Timer.Seconds = 1
            End If
            If Timer.Finished Then
              State = Estate.FillStart
              .PumpControl.StopMainPump()
              Timer.Seconds = MinMax(.Parameters.FillOpenDelayTime, 5, 60)
            End If
            If .MachineLevel >= FillLevel Then
              State = Estate.PumpSettle
              If Not .PumpControl.IsActive Then .PumpControl.StartAuto()
            End If

          Case Estate.FillStart
            If Timer.Finished Then
              BlendControl.Start(FillTemp)
              State = Estate.FillToLevel
            End If
            Status = commandName_ & Translate("Starting") & Timer.ToString(1)

          Case Estate.FillToLevel
            BlendControl.Run(.IO.FillTemp)
            If .MachineLevel >= FillLevel Then
              State = Estate.FillSettle1
              Timer.Seconds = .Parameters.FillSettleTime
            End If
            Status = commandName_ & Translate("Filling") & (" ") & (.MachineLevel / 10).ToString("#0.0") & " / " & (FillLevel / 10).ToString("#0.0") & "%"

          Case Estate.FillSettle1
            If Timer.Finished Then
              State = Estate.PumpSettle
              .PumpControl.StartAuto()
              Timer.Seconds = .Parameters.FillPrimePumpTime
            End If
            Status = commandName_ & Translate("Settling") & Timer.ToString(1)

          Case Estate.PumpSettle
            'If Not .IO.PumpInAutoSw Then StateString = Translate("Pump Switch Not in Auto")
            If Not .IO.PumpRunning Then
              Timer.Seconds = .Parameters.FillPrimePumpTime
            Else
              If Not .PumpControl.IsRunning Then Timer.Seconds = .Parameters.FillPrimePumpTime
              ' Timer has finished and pump is running - Also consider pump started and loss of water in expansion tank could run pump dry, begin filling immediately
              If Timer.Finished OrElse (.IO.PumpRunning AndAlso (.MachineLevel < (.PumpControl.Parameters_PumpMinimumLevel - 50))) Then
                State = Estate.TopUp
              End If
            End If
            Status = commandName_ & Translate("Settle with pump") & Timer.ToString(1)
            If .PumpControl.IsStarting Then Status = commandName_ & .PumpControl.StateString

          Case Estate.TopUp
            BlendControl.Run(.IO.FillTemp)
            If .MachineLevel >= FillLevel Then
              NumberOfTopups += 1
              Timer.Seconds = MinMax(.Parameters.FillSettleTime, 10, 120)
              State = Estate.FillSettle2
            End If
            Status = commandName_ & Translate("Top Up") & (" ") & (.MachineLevel / 10).ToString("#0.0") & " / " & (FillLevel / 10).ToString("#0.0") & "%"

          Case Estate.FillSettle2
            If Timer.Finished Then
              ' Check Level after TopUp and Settle, may need to TopUp again
              If (.MachineLevel >= FillLevel) Then
                .HD.HDCompleted = False
                .AirpadOn = True
                State = Estate.Complete
                Timer.Seconds = 2
              Else
                State = Estate.TopUp
                Timer.Seconds = .Parameters.OverFillTime
              End If
            End If
            Status = commandName_ & Translate("Settle TopUp") & (" ") & (NumberOfTopups.ToString("#0")) & Timer.ToString(1)

          Case Estate.Complete
            If Timer.Finished Then
              ' Record Level once airpad has equalized after complete delay
              .GetRecordedLevel = True
              State = Estate.Off
            End If
            Status = commandName_ & Translate("Completing") & Timer.ToString(1)

        End Select

      Loop Until (StartState = State)

    End With
  End Function

  Public Sub Cancel() Implements ACCommand.Cancel
    Timer.Cancel()
    TimerOverrun.Cancel()
    FillLevel = 0
    FillTemp = 0
    '   FillType = EFillType.Cold
    State = Estate.Off
    Status = ""
  End Sub

  ReadOnly Property IsOn As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> Estate.Off)
    End Get
  End Property

  ReadOnly Property IsActive As Boolean
    Get
      Return (State > Estate.Interlock) AndAlso (State < Estate.Complete)
    End Get
  End Property

  ReadOnly Property IsAlarm As Boolean
    Get
      Return IsOn AndAlso TimerAlarm.Finished AndAlso (controlCode.Parameters.FillAlarmTime > 0)
    End Get
  End Property

  ReadOnly Property IsOverrun As Boolean
    Get
      Return IsActive AndAlso TimerOverrun.Finished AndAlso (controlCode.Parameters.StandardTimeFillMachine > 0)
    End Get
  End Property

  ReadOnly Property ResetFlowmeter As Boolean
    Get
      Return State = Estate.ResetMeter
    End Get
  End Property

  ReadOnly Property IoFillSelect As Boolean
    Get
      Return (State = Estate.FillToLevel) OrElse (State = Estate.TopUp)
    End Get
  End Property

  'ReadOnly Property IoFillCold As Boolean
  '  Get
  '    Return IoFillSelect AndAlso ((FillType = EFillType.Cold) OrElse (FillType = EFillType.Mix))
  '  End Get
  'End Property

  'ReadOnly Property IoFillHot As Boolean
  '  Get
  '    Return IoFillSelect AndAlso ((FillType = EFillType.Hot) OrElse (FillType = EFillType.Mix))
  '  End Get
  'End Property

  ReadOnly Property IoTopWash As Boolean
    Get
      Return (State >= Estate.FillStart) AndAlso (State <= Estate.FillSettle2)
    End Get
  End Property

  ReadOnly Property IOBlendOutput As Integer
    Get
      If IoFillSelect Then Return BlendControl.IOOutput
    End Get
  End Property


  ' TODO
#If 0 Then
  

'Fill command
Option Explicit
Implements ACCommand
Public Enum FIState
  FIOff
  FIInterlock
  FINotSafe
  FIFillAndFlush
  FIFillToLevel
  FISettle
  FIPumpSettle
  FISettleSecondTime
  FITopUp
End Enum

Public FlushTimer As New acTimer



  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode
    .


        
    State = FIInterlock
    BlendControl.Params .Parameters.BlendFactor, .Parameters.BlendDeadBand, 5, 100, _
                        .Parameters.ColdWaterTemperature, .Parameters.HotWaterTemperature
 End With
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode:
    
    If (State > FIInterlock) And ((Not .MachineSafe) Or .HD.HDCompleted) And (Not .LidLocked) Then State = FINotSafe
    
    Select Case State
    
      Case FIOff
        StateString = ""
        
      Case FIInterlock
        If .TempSafe And Not .PressSafe Then
          StateString = "FI: " & .SafetyControl.StateString
        ElseIf Not .LidLocked Then
          StateString = "FI: Lid Not Closed "
        Else
          StateString = "FI: Unsafe to Fill "
        End If
        If (.MachineSafe Or .HD.HDCompleted) And .LidLocked Then
          .TP.ACCommand_Cancel
          .HE.ACCommand_Cancel
          .CO.ACCommand_Cancel
          .TemperatureControl.Cancel
          .TC.ACCommand_Cancel
          .TemperatureControlContacts.Cancel
          BlendControl.Start FillTemperature
          If (.VesselLevel >= FillLevel) Then
            State = FIPumpSettle
            Timer = .Parameters.FillSettleTime
            .PumpRequest = True
          Else
            State = FIFillToLevel
            .PumpRequest = False
            FlushTimer = .Parameters.LevelGaugeFlushTime
          End If
        End If
      
      Case FINotSafe
        If .TempSafe And Not .PressSafe Then
          StateString = "FI: " & .SafetyControl.StateString
        ElseIf Not .LidLocked Then
          StateString = "FI: Lid Not Closed "
        Else
          StateString = "FI: Unsafe to Fill "
        End If
        If .MachineSafe And .LidLocked Then State = FIFillToLevel
      
      Case FIFillAndFlush
        StateString = "FI: Filling " & Pad(.VesselLevel / 10, "0", 3) & "%"
        BlendControl.Run .IO_BlendFillTemp
        If FlushTimer.Finished Then State = FIFillToLevel
        
      Case FIFillToLevel
        StateString = "FI: Filling " & Pad(.VesselLevel / 10, "0", 3) & " / " & Pad(FillLevel / 10, "0", 3) & "% "
        BlendControl.Run .IO_BlendFillTemp
        If (.VesselLevel >= FillLevel) Then
          State = FISettle
          Timer = .Parameters.FillSettleTime
        End If
        
      Case FISettle
        StateString = "FI: Settle " & TimerString(Timer.TimeRemaining)
        BlendControl.Run .IO_BlendFillTemp
        If Timer.Finished Then
          State = FIPumpSettle
          Timer = .Parameters.FillPrimePumpTime
          .PumpRequest = True
          .PumpOnCount = 0
        End If
      
      Case FIPumpSettle
        StateString = "FI: Settle with Pump " & TimerString(Timer.TimeRemaining)
        BlendControl.Run .IO_BlendFillTemp
        If Timer.Finished Then
          Timer = .Parameters.FillSettleTime
          State = FISettleSecondTime
        End If
      
      Case FISettleSecondTime
        StateString = "FI: Settle " & TimerString(Timer.TimeRemaining)
        BlendControl.Run .IO_BlendFillTemp
        If Timer.Finished Then
          State = FITopUp
        End If
        
      Case FITopUp
        StateString = "FI: TopUp " & Pad(.VesselLevel / 10, "0", 3) & " / " & Pad(FillLevel / 10, "0", 3) & "% "
        BlendControl.Run .IO_BlendFillTemp
        If (.VesselLevel >= FillLevel) Then
          State = FIOff
          .HD.HDCompleted = False
          .GetRecordedLevel = True
          .AirpadOn = True 'pressurizes the machine
        End If
        
    End Select
  End With
End Sub

Friend Sub ACCommand_Cancel()
  State = FIOff
End Sub
Friend Property Get IsOn() As Boolean
  IsOn = (State <> FIOff)
End Property
Friend Property Get IsActive() As Boolean
  If (State >= FIFillAndFlush) Then IsActive = True
End Property
Friend Property Get IsFilling() As Boolean
  If (State = FIFillToLevel) Or (State = FITopUp) Or (State = FIFillAndFlush) Then IsFilling = True
End Property
Friend Property Get IsNotSafe() As Boolean
  IsNotSafe = (State = FINotSafe) Or (State = FIInterlock)
End Property
Friend Property Get IsFillToLevel() As Boolean
  If (State = FIFillToLevel) Or (State = FIFillAndFlush) Then IsFillToLevel = True
End Property
Friend Property Get IsFillAndFlushing() As Boolean
 If (State = FIFillAndFlush) Then IsFillAndFlushing = True
End Property
Friend Property Get IsSettle() As Boolean
  If (State = FISettle) Or (State = FISettleSecondTime) Then IsSettle = True
End Property
Friend Property Get IsSettleWithPump() As Boolean
  If (State = FIPumpSettle) Then IsSettleWithPump = True
End Property
Friend Property Get IsTopUp() As Boolean
  If (State = FITopUp) Then IsTopUp = True
End Property
Friend Property Get IsOverrun() As Boolean
  If IsFilling And FillTimer.Finished Then IsOverrun = True
End Property
Friend Property Get FITimer() As Long
  FITimer = Timer.TimeRemaining
End Property
Private Property Get ACCommand_IsOn() As Boolean
ACCommand_IsOn = IsOn
End Property
Friend Property Get Output() As Long
'Blend output
  Output = BlendControl.Output

'Limit output
  MinMax Output, 0, 1000
End Property


#End If
End Class
