'American & Efird - Mt Holly Then Platform
' Version 2024-08-21

Imports Utilities.Translations

<Command("Reserve Fill", "|HCMV| |0-99|% |0-180|F Mix? |0-1|", "", "", ""),
TranslateCommand("es", "Agregar Relleno", "|HCMV| |0-99|% |0-180|F Mix:|0-1|"),
Description("Fills the reserve tank to the desired level (in percent). Heats and Mixes."),
TranslateDescription("es", "Llene el tanque de reserva a nivel desired."),
Category("Reserve Tank Commands"), TranslateCategory("es", "Reserve Tank Commands")>
Public Class RF
  Inherits MarshalByRefObject
  Implements ACCommand
  Private Const commandName_ As String = ("RF: ")

  'Command states
  Public Enum EState
    Off
    Interlock
    Depressurize
    FillFromMachineStart
    FillFromMachine
    Fill
    Heating
    HeatMaintain
    Done
  End Enum
  Public State As EState
  Public Status As String

  Public FillLevel As Integer
  Public FillType As Integer
  Public TempDesired As Integer
  Public HeatOn As Boolean
  Public AirpadOn As Boolean

  Public Enum EMixState
    Off
    Mixer
  End Enum
  Public MixState As EMixState
  Public MixValue As Integer

  Public Timer As New Timer
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

    'Enable command parameter adjustments, if parameter set
    If controlCode.Parameters.EnableCommandChg = 1 Then
      Start(param)
    End If
  End Sub

  Public Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With controlCode

      ' Cancel Reserve Commands
      .RD.Cancel()
      .RW.Cancel()

      ' Reset parameters
      FillType = 0
      FillLevel = 0
      TempDesired = 0
      MixValue = 0

      'Command parameters - check array bounds just to be on the safe side
      ' Fill Type [H=72, C=67, M=77, V=86]
      Dim intFillType As Integer = 0
      If param.GetUpperBound(0) >= 1 Then intFillType = param(1)
      If param.GetUpperBound(0) >= 2 Then FillLevel = param(2) * 10
      If param.GetUpperBound(0) >= 3 Then TempDesired = MinMax(param(3) * 10, 0, 800)
      If param.GetUpperBound(0) >= 4 Then MixValue = (param(4))

      'Set the default fill level 
      FillLevel = MinMax(FillLevel, 0, 1000)

      ' Fill Type [H=72, C=67, M=77, V=86]
      FillType = EFillType.Cold
      Select Case intFillType
        Case 86
          FillType = EFillType.Vessel
        Case 77
          FillType = EFillType.Mix
        Case 72
          FillType = EFillType.Hot
        Case Else 'C=67
          FillType = EFillType.Cold
      End Select

      ' Mix Value - Then Platform has a Reserve Mixer
      If MixValue = 0 Then MixState = EMixState.Off
      If MixValue = 1 Then MixState = EMixState.Mixer

      ' If Heating desired - check level against minimum level required for heating
      If (TempDesired > 0) Then
        ' Limit Fill Level
        If FillLevel < .Parameters.ReserveMixerOnLevel Then FillLevel = .Parameters.ReserveMixerOnLevel
        If FillLevel < .Parameters.ReserveHeatEnableLevel Then FillLevel = .Parameters.ReserveHeatEnableLevel
        If FillLevel > 750 Then FillLevel = 750

        ' Limit Temp Desired
        If .Parameters.ReserveTankHighTempLimit > 1800 Then .Parameters.ReserveTankHighTempLimit = 1800
        If TempDesired > .Parameters.ReserveTankHighTempLimit Then TempDesired = .Parameters.ReserveTankHighTempLimit
        If TempDesired > 1800 Then TempDesired = 1800

        ' Override MixOn parameter if heating
        If MixValue = 0 Then MixState = EMixState.Mixer
      End If

      ' Clear Reserve Ready flag
      If FillType = EFillType.Vessel AndAlso .KP.Tank1Destination = EKitchenDestination.Reserve Then
        '  Do not clear reserve ready flag
      Else
        '  .ReserveReady = False
      End If

      'Use Command FillLevel or Percent of Working Level based on NP & PT:
      If (.Parameters.EnableFillByWorkingLevel = 1) Then
        If (.Parameters.EnableFillByWorkingLevel = 1) And (.WorkingLevel > 0) Then
          FillLevel = CInt(param(2) * (.WorkingLevel / 100))
          If FillLevel > 1000 Then FillLevel = 1000
        End If
        If TempDesired > 0 Then
          If FillLevel < .Parameters.ReserveMixerOnLevel Then FillLevel = .Parameters.ReserveMixerOnLevel
        End If
      End If

      ' Step On if fresh fill
      If FillType <> EFillType.Vessel Then Return True

      ' Set initial state
      State = EState.Interlock
      HeatOn = False
      Timer.Seconds = 2
      TimerOverrun.Minutes = 5
    End With
  End Function

  Public Function Run() As Boolean Implements ACCommand.Run
    With controlCode

      Select Case State
        Case EState.Off
          Status = (" ")

        Case EState.Interlock
          Status = Translate("Wait Idle") & Timer.ToString(1)
          If .KA.IsOn AndAlso (.KA.Destination = EKitchenDestination.Reserve) Then Status = commandName_ & .KA.Status : Timer.Seconds = 2
          If .KP.IsOn AndAlso (.KP.KP1.Param_Destination = EKitchenDestination.Reserve) AndAlso (.KP.KP1.IsForeground) Then Status = commandName_ & .KP.KP1.DrugroomDisplay : Timer.Seconds = 2
          If .LA.IsOn AndAlso (.LA.KP1.Param_Destination = EKitchenDestination.Reserve) AndAlso (.LA.KP1.IsForeground) Then Status = commandName_ & .LA.KP1.DrugroomDisplay : Timer.Seconds = 2
          If Timer.Finished Then
            ' Kitchen complete or idle
            If FillType = EFillType.Vessel Then
              ' Check Machine Safe conditions
              If .TempSafe AndAlso .MachineClosed Then
                .CO.Cancel() : .HE.Cancel() : .TP.Cancel()
                .TemperatureControl.Cancel()
                .AirpadOn = False
                .PR.Cancel()
                .PumpControl.StopMainPump()
                State = EState.Depressurize
                Timer.Seconds = 2
              Else
                If Not .MachineClosed Then Status = commandName_ & Translate("Runback wait for Machine Closed")
                If Not .TempSafe Then Status = commandName_ & Translate("Temp Not Safe") & Timer.ToString(1)
              End If
            Else
              State = EState.Fill
            End If
          End If


        Case EState.Depressurize
          If Not .MachineClosed Then Status = commandName_ & Translate("Runback wait for Machine Closed") : Timer.Seconds = 2
          If Not .PressSafe Then Status = commandName_ & Translate("Depressuzing") & (" ") & .MachinePressureDisplay : Timer.Seconds = 2
          .GetRecordedLevel = False
          .Alarms.LosingLevelInTheVessel = False
          .AirpadOn = False
          If .MachineClosed AndAlso .TempSafe AndAlso .PressSafe AndAlso Timer.Finished Then
            State = EState.FillFromMachineStart
            Timer.Seconds = MinMax(.Parameters.ReserveBackFillStartTime, 10, 60)
          End If
          Status = commandName_ & Translate("Depressuzing") & Timer.ToString(1)


        Case EState.FillFromMachineStart
          If Timer.Finished Then
            AirpadOn = True
            State = EState.FillFromMachine
          End If
          If .ReserveLevel >= FillLevel Then
            Return True
            AirpadOn = False
            State = EState.Heating
          End If
          If Not (.TempSafe AndAlso .MachineClosed) Then State = EState.Depressurize
          Status = commandName_ & Translate("Fill From Vessel Starting") & Timer.ToString(1)


        Case EState.FillFromMachine
          If Not .MachineSafe Then State = EState.Depressurize
          If (.MachinePressure <= 1500) Then                    'SafePressure = 2758 (4psi), so only pressure up to 1.5bar so that backfill works
            AirpadOn = True
          Else : AirpadOn = False
          End If
          If .ReserveLevel >= FillLevel Then
            AirpadOn = False
            State = EState.Heating
          End If
          Status = commandName_ & Translate("Fill fom machine") & (" ") & (.ReserveLevel / 10).ToString("#0.0") & " / " & (FillLevel / 10).ToString("#0.0") & "% "

        Case EState.Fill
          If .ReserveLevel < FillLevel Then
            Timer.Seconds = .Parameters.ReserveOverfillTime
          Else
            Status = Translate("Filling") & Timer.ToString(1)
          End If
          If Timer.Finished AndAlso (.ReserveLevel > FillLevel) Then
            State = EState.Heating
            Timer.Seconds = 2
          End If
          Status = commandName_ & Translate("Filling") & (" ") & (.ReserveLevel / 10).ToString("#0.0") & " / " & (FillLevel / 10).ToString("#0.0") & "% "


        Case EState.Heating
          If (TempDesired > 0) Then
            ' Maintain temp limit in case parameter updated
            If TempDesired > .Parameters.ReserveTankHighTempLimit Then TempDesired = .Parameters.ReserveTankHighTempLimit
            If TempDesired > 1400 Then TempDesired = 1400 : If .Parameters.ReserveTankHighTempLimit > 1400 Then .Parameters.ReserveTankHighTempLimit = 1400
            ' Monitor Tank Temp 
            If (.IO.ReserveTemp >= (TempDesired - .Parameters.ReserveHeatDeadband)) Then
              State = EState.HeatMaintain
              ' This is a background command so tell the control system to step on
              Return True
            End If
          Else
            State = EState.Done
            Timer.Seconds = 2
          End If
          ' Check Level continuously
          If (.ReserveLevel < FillLevel) OrElse (.Parameters.ReserveTempProbe = 0) Then
            State = EState.Done
            Timer.Seconds = 2
          End If
          Status = Translate("Heating") & (" ") & (.IO.ReserveTemp / 10).ToString("#0.0") & " / " & (TempDesired / 10).ToString("#0.0") & " F"

        Case EState.HeatMaintain
          Status = Translate("Heating") & (" ") & (.IO.ReserveTemp / 10).ToString("#0.0") & " / " & (TempDesired / 10).ToString("#0.0") & " F"
          If (.IO.ReserveTemp < (TempDesired - .Parameters.ReserveHeatDeadband)) Then HeatOn = True
          If (.IO.ReserveTemp >= (TempDesired - .Parameters.ReserveHeatDeadband)) Then HeatOn = False
          ' Check Level continuously
          If (.ReserveLevel < FillLevel) OrElse (.Parameters.ReserveTempProbe = 0) Then
            State = EState.Done
            Timer.Seconds = 2
          End If

        Case EState.Done
          If Timer.Finished Then Cancel()
          Status = Translate("Completing") & Timer.ToString(1)
          If (TempDesired > 0) AndAlso (.Parameters.ReserveTempProbe = 0) Then Status = Translate("Heat Not Enabled") & Timer.ToString(1)

      End Select
    End With
  End Function

  Public Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
    Status = ""
    MixState = EMixState.Off
    AirpadOn = False
    HeatOn = False
    MixValue = 0
    FillLevel = 0
    TempDesired = 0
    FillType = EFillType.Cold
  End Sub

  Public ReadOnly Property IsOn() As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property

  Public ReadOnly Property IsInterlocked As Boolean
    Get
      Return (State = EState.Interlock)
    End Get
  End Property

  Public ReadOnly Property IsRunback As Boolean
    Get
      Return (State = EState.FillFromMachine)
    End Get
  End Property

  Public ReadOnly Property IsForeground As Boolean
    Get
      Return (State = EState.Interlock) OrElse (State = EState.Depressurize) OrElse (State = EState.FillFromMachine) 'OrElse (State = EState.Heating)
      ' If (State = Interlock) Or (State = Depressurize) Or (State = FillFromVesselStart) Or (State = FillFromVessel) Then IsForeground = True TODO Check
    End Get
  End Property

  Public ReadOnly Property IsFillFromVessel As Boolean
    Get
      Return (State = EState.FillFromMachineStart) OrElse (State = EState.FillFromMachine)
    End Get
  End Property

  Public ReadOnly Property IsOverrun As Boolean
    Get
      Return IsForeground AndAlso TimerOverrun.Finished
      ' vb6: If IsOn And (State <> HeatMaintain) And OverrunTimer.Finished Then IsOverrun = True
    End Get
  End Property

  Public ReadOnly Property IsHeatMaintain As Boolean
    Get
      Return (State = EState.HeatMaintain)
    End Get
  End Property

  Public ReadOnly Property IoAirpad As Boolean
    Get
      Return (State = EState.FillFromMachine) AndAlso AirpadOn
    End Get
  End Property

  Public ReadOnly Property IoReserveFillCold As Boolean
    Get
      Return (State = EState.Fill) AndAlso (FillType = EFillType.Cold OrElse FillType = EFillType.Mix)
    End Get
  End Property

  Public ReadOnly Property IoReserveFillHot As Boolean
    Get
      Return (State = EState.Fill) AndAlso (FillType = EFillType.Hot OrElse FillType = EFillType.Mix)
    End Get
  End Property

  Public ReadOnly Property IoReserveHeat As Boolean
    Get
      Return (State = EState.Heating) OrElse ((State = EState.HeatMaintain) AndAlso HeatOn)
    End Get
  End Property

  Public ReadOnly Property IoReserveMixer As Boolean
    Get
      Return (State = EState.Heating) OrElse (State = EState.HeatMaintain)
    End Get
  End Property

  Public ReadOnly Property IoReserveRunback As Boolean
    Get
      Return (State = EState.FillFromMachineStart) OrElse (State = EState.FillFromMachine)
    End Get
  End Property

#If 0 Then  ' TODO Check & Remove vb6

'===============================================================================================
'RF - Reserve tank fill command
'  2022-02-14
'===============================================================================================
Option Explicit
Implements ACCommand
Public Enum RFState
  Off
  Interlock
  Depressurize
  FillFromVesselStart
  FillFromVessel
  Fill
  Heating
  HeatMaintain
End Enum
Public State As RFState
Public StateString As String
Public FillLevel As Long
Public FillType As Long
Public DesiredTemperature As Long
Public HeatOn As Boolean
Public ReserveMixing As Boolean
Public Timer As New acTimer
Public OverrunTimer As New acTimer

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Parameters= |HCMV| |0-99|% |0-180|F Mix?|0-1|\r\nName=Reserve Fill\r\nHelp=Fills the reserve tank to the specified level with Hot, Cold, Mixed or from the vessel. Heats the reserve tank to the desired temperature and signals the operator to prepare an a RP."
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  
  ' Command Parameters (Attributes):
  '   Parameters= |HCMV| |0-99|% |0-180|F Mix?|0-1|
  '   Name=Reserve Fill
  '   Help=Fills the reserve tank to the specified level with the specified water type. Heats to the desired temperature.

  FillType = 0
  FillLevel = 0
  DesiredTemperature = 0
  ReserveMixing = False
  
  ' Notes for fill type H=72, C=67, M=77, V=86
  ' Letter Parameters would normally be a variant type, but Plant explorer doesn't display variants so convert to long
  If Param(1) = 86 Then
    FillType = 86 'Runback from vessel
  ElseIf Param(1) = 77 Then
    FillType = 77
  ElseIf Param(1) = 72 Then
    FillType = 72
  Else: FillType = 67
  End If
  
  FillLevel = Param(2) * 10
  DesiredTemperature = Param(3) * 10
  If DesiredTemperature >= 1800 Then DesiredTemperature = 1800
  
  If (FillType <> 86) Then StepOn = True
  If Param(4) = 1 Then ReserveMixing = True
  
  With ControlCode
    .RW.ACCommand_Cancel
    
    'Use Command FillLevel or Percent of Working Level based on NP & PT:
    If (.Parameters_EnableReserveFillByWorkingLevel = 1) Then
      If (.Parameters_EnableFillByWorkingLevel = 1) And (.WorkingLevel > 0) Then
        FillLevel = Param(2) * (.WorkingLevel) / 100
        If FillLevel > 1000 Then FillLevel = 1000
      End If
      If DesiredTemperature > 0 Then
        If FillLevel < .Parameters_ReserveMixerOnLevel Then FillLevel = .Parameters_ReserveMixerOnLevel
      End If
    End If
  End With

  OverrunTimer.TimeRemaining = 300
  State = Interlock
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
  Dim ControlCode As ControlCode: Set ControlCode = ControlObject

  With ControlCode
    Select Case State
       
      Case Off
        StateString = ""
        
      Case Interlock
        OverrunTimer.TimeRemaining = 300
        If (.KP.KP1.DispenseTank = 2) Then
          StateString = "RF Drugroom: " & .KP.KP1.StateString
        ElseIf .KA.Destination = 82 Then
          StateString = "RF Drugroom: " & .KA.StateString
        Else
          If FillType = 86 Then
            If Not .LidLocked Then
              Timer = 2
              StateString = "RF: Runback Wait for Lid Locked"
            End If
            If Not .TempSafe Then
              Timer = 2
              StateString = "RF: Runback Wait for TempSafe"
            End If
          End If
        End If
        If Timer.Finished Then
          If FillType = 86 Then
            .AirpadOn = False
            .PumpRequest = False
            Timer.TimeRemaining = 2
            State = Depressurize
          Else
            State = Fill
          End If
        End If
       
      Case Depressurize
        If .TempSafe And Not .PressSafe Then
          StateString = "RF: Wait Safe " & .SafetyControl.StateString
          Timer.TimeRemaining = 2
        Else
          StateString = "RF: Interlock " & TimerString(Timer.TimeRemaining)
        End If
        .GetRecordedLevel = False
        .Alarms_LosingLevelInTheVessel = False
        .AirpadOn = False
        If .MachineSafe And .LidLocked And Timer.Finished Then
          State = FillFromVesselStart
          Timer.TimeRemaining = .Parameters_ReserveBackFillStartTime
          If Timer.TimeRemaining < 10 Then Timer.TimeRemaining = 10
        End If

      Case FillFromVesselStart
        StateString = "RF: Fill From Vessel Start " & TimerString(Timer.TimeRemaining)
        If Not .TempSafe Then State = Depressurize
        If Timer.Finished Then
          State = FillFromVessel
        End If
        If (.ReserveLevel >= FillLevel) Then
          StepOn = True
          .AirpadOn = False
          State = Heating
        End If
      
      Case FillFromVessel
        StateString = "RF: Fill From Vessel " & Pad(.ReserveLevel / 10, "0", 3) & "% / " & Pad(FillLevel / 10, "0", 3) & "% "
        If Not .TempSafe Then State = Depressurize
        .AirpadOn = True
        If (.ReserveLevel >= FillLevel) Then
          StepOn = True
          .AirpadOn = False
          State = Heating
        End If
      
      Case Fill
        StateString = "RF: Fill Fresh " & Pad(.ReserveLevel / 10, "0", 3) & "% / " & Pad(FillLevel / 10, "0", 3) & "% "
        If (.ReserveLevel >= FillLevel) Then State = Heating
      
      Case Heating
        If DesiredTemperature > 0 Then
          If Not .ReserveMixerOn Then
            StateString = "RF: Reserve level too low for mixer "
          Else: StateString = "RF: Heating " & Pad(.IO_ReserveTemp / 10, "0.0", 3) & "F / " & Pad(DesiredTemperature / 10, "0.0", 3) & "F"
          End If
          If (.IO_ReserveTemp >= (DesiredTemperature - .Parameters_ReserveHeatDeadband)) Then State = HeatMaintain
        Else
          ACCommand_Cancel
        End If
        
      Case HeatMaintain
        If Not .ReserveMixerOn Then
          StateString = "RF: Reserve level too low for mixer "
        Else: StateString = "RF: Heat Maintain " & Pad(.IO_ReserveTemp / 10, "0.0", 3) & "F / " & Pad(((DesiredTemperature - .Parameters_ReserveHeatDeadband) / 10), "0.0", 3) & "F"
        End If
        If (.IO_ReserveTemp < (DesiredTemperature - .Parameters_ReserveHeatDeadband)) Then HeatOn = True
        If (.IO_ReserveTemp >= (DesiredTemperature - .Parameters_ReserveHeatDeadband)) Then HeatOn = False
        
    End Select
  End With
End Sub

Public Sub ACCommand_Cancel()
  HeatOn = False
  State = Off
  FillType = 0
  FillLevel = 0
  DesiredTemperature = 0
  ReserveMixing = False
End Sub
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property
Friend Property Get IsOn() As Boolean
  If (State > Off) Then IsOn = True
End Property
Friend Property Get IsForeground() As Boolean
  If (State = Interlock) Or (State = Depressurize) Or (State = FillFromVesselStart) Or (State = FillFromVessel) Then IsForeground = True
End Property
Friend Property Get IsInterlocked() As Boolean
  If (State = Interlock) Then IsInterlocked = True
End Property
Friend Property Get IsFilling() As Boolean
  If (State = Fill) Then IsFilling = True
End Property
Friend Property Get IsFillWithCold() As Boolean
  If (State = Fill) And ((FillType = 67) Or (FillType = 77)) Then IsFillWithCold = True
End Property
Friend Property Get IsFillWithHot() As Boolean
  If (State = Fill) And ((FillType = 72) Or (FillType = 77)) Then IsFillWithHot = True
End Property
Friend Property Get IsFillWithVessel() As Boolean
  If ((State = FillFromVesselStart) Or (State = FillFromVessel)) And (FillType = 86) Then IsFillWithVessel = True
End Property
Friend Property Get IsHeating() As Boolean
  If (State = Heating) Or HeatOn Then IsHeating = True
End Property
Friend Property Get IsOverrun() As Boolean
  If IsOn And (State <> HeatMaintain) And OverrunTimer.Finished Then IsOverrun = True
End Property

#End If
End Class
