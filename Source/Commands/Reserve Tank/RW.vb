'American & Efird - Mt. Holly Then Platform
' Version 2024-08-23
Imports Utilities.Translations

<Command("Reserve Wash", "|0-99|%  |0-180|F |0-99| mins", "", "", "'3 + 5"),
TranslateCommand("es", "Tanque de reserva transferencia", ""),
Description("Fill the reserve tank to the specified level. Heat the tank to the specified temperature. Transfer and fill the reserve tank to the machine and overflow wash it."),
TranslateDescription("es", ""),
Category("Reserve Tank Commands"), TranslateCategory("es", "Reserve Tank Commands")>
Public Class RW : Inherits MarshalByRefObject : Implements ACCommand
  Private ReadOnly controlCode As ControlCode
  Private Const commandName_ As String = ("RW: ")

  Public Enum EState
    Off
    Start
    Fill
    Heat
    Interlock
    Paused
    TransferAndFill
    TransferNoFill
    Transfer1
    Rinse
    Transfer2
    RinseToDrainPause
    RinseToDrain
    TransferToDrain
    Done
  End Enum
  Property State As EState
  Property StateWas As EState
  Property Status As String
  Public Property NumberOfRinses As Integer
  Property HoldTime As Integer
  Property MachineNotClosed As Boolean
  Property FillLevel As Integer
  Property DesiredTemp As Integer
  Property WashTime As Integer

  Public ReadOnly Timer As New Timer
  Friend TimerHold As New Timer
  Public ReadOnly Property TimeHold As Integer
    Get
      Return TimerHold.Seconds
    End Get
  End Property
  Friend Property TimerOverrun As New Timer
  Friend TimerAdvance As New Timer

  Public Sub New(ByVal controlCode As ControlCode)
    Me.controlCode = controlCode
    Me.controlCode.Commands.Add(Me)
  End Sub

  Public Sub ParametersChanged(ParamArray param() As Integer) Implements ACCommand.ParametersChanged
    'If the command is not running then do nothing the changes will be picked up when the command starts
    If Not IsOn Then Exit Sub

    'Enable command parameter adjustments, if parameter set
    If controlCode.Parameters.EnableCommandChg = 1 Then
      Start(param)
    End If
  End Sub

  Public Function Start(ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With controlCode
      FillLevel = 0
      DesiredTemp = 0
      HoldTime = 0
      If param.GetUpperBound(0) >= 1 Then FillLevel = param(1) * 10
      If param.GetUpperBound(0) >= 2 Then DesiredTemp = param(2) * 10
      If param.GetUpperBound(0) >= 3 Then HoldTime = param(3) * 60

      If DesiredTemp > 1800 Then DesiredTemp = 1800
      If DesiredTemp > 0 Then
        If DesiredTemp < .Parameters.ReserveMixerOnLevel Then FillLevel = .Parameters.ReserveMixerOnLevel
      End If

      ' Cancel Commands
      .AC.Cancel() : .AT.Cancel()
      .KA.Cancel() : .WK.Cancel()
      .DR.Cancel() : .FI.Cancel() : .HC.Cancel() : .HD.Cancel()
      .RC.Cancel() : .RH.Cancel() : .RI.Cancel() : .TM.Cancel() ' TC?
      .LD.Cancel() : .PH.Cancel() : .SA.Cancel() : .UL.Cancel()
      .RD.Cancel() : .RT.Cancel()  ' RP, RT? .RF.Cancel() :
      .WT.Cancel()

      If .RF.FillLevel = EFillType.Vessel Then .RF.Cancel()

      TimerOverrun.Seconds = HoldTime + 180
      State = EState.Start
      StateWas = EState.Start

    End With
  End Function

  Public Function Run() As Boolean Implements ACCommand.Run
    With controlCode

      ' Pause if active and problem occurs
      Dim pauseCommand As Boolean = .EStop OrElse (.Parent.IsPaused) ' OrElse (Not .IO.PumpRunning) OrElse (.TemperatureControl.IsCrashCoolOn)
      If (State > EState.Paused) AndAlso (State < EState.Done) AndAlso pauseCommand Then
        ' StateWas = State
        State = EState.Paused
        TimerHold.Pause()
      End If

      ' State Logic
      Select Case State
        Case EState.Off
          Status = ""
          StateWas = EState.Off


        Case EState.Start
          If Timer.Finished Then
            State = EState.Fill
            StateWas = EState.Fill
          End If


        Case EState.Fill
          If .ReserveLevel < FillLevel Then Timer.Seconds = .Parameters.ReserveOverfillTime
          If Timer.Finished AndAlso (.ReserveLevel > FillLevel) Then
            .AirpadOn = False
            .ReserveMixerOn = True
            State = EState.Heat
            Timer.Seconds = 2
          End If
          If .ReserveLevel < FillLevel Then
            Status = commandName_ & Translate("Filling") & (" ") & (.ReserveLevel / 10).ToString("#0.0") & " / " & (FillLevel / 10).ToString("#0.0") & "% "
          Else
            Status = commandName_ & Translate("Filling") & Timer.ToString(1)
          End If


        Case EState.Heat
          'If (DesiredTemp > 0) Then
          ' Maintain temp limit in case parameter updated
          If DesiredTemp > .Parameters.ReserveTankHighTempLimit Then DesiredTemp = .Parameters.ReserveTankHighTempLimit
          If DesiredTemp > 750 Then DesiredTemp = 750 : If .Parameters.ReserveTankHighTempLimit > 750 Then .Parameters.ReserveTankHighTempLimit = 750
          ' Monitor Tank Temp 
          If (.IO.ReserveTemp >= (DesiredTemp - .Parameters.ReserveHeatDeadband)) Then
            State = EState.Interlock
            .ReserveMixerOn = False
            'This is a background command so tell the control system to step on
            Return True
          End If


        Case EState.Interlock
          If .TempSafe AndAlso Not .PressSafe Then
            Status = commandName_ & Translate("Wait Safe") & (" ") & .SafetyControl.StateString
          Else
            Status = commandName_ & Translate("Not Safe")
          End If
          If (.MachineSafe Or .HD.HDCompleted) Then
            .CO.Cancel()
            .HE.Cancel()
            'TC.Cancel() ' TODO
            .TP.Cancel()
            .TemperatureControl.Cancel()
            '   .TemperatureControlContacts.Cancel 'TODO
            State = EState.TransferAndFill
            Timer.Seconds = HoldTime
          End If


        Case EState.Paused
          If Not pauseCommand Then
            State = StateWas
            Timer.Restart()
          End If
          Status = commandName_ & Translate("Paused")



        Case EState.TransferAndFill
          ' If (.PumpRequest = False) Then .PumpRequest = True
          .PumpControl.StartAuto()
          If .ReserveLevel > FillLevel Then
            State = EState.TransferNoFill
          End If
          If Timer.Finished Then
            State = EState.Transfer1
          End If
          Status = commandName_ & Translate("Transfer with Fill") & Timer.ToString(1)


        Case EState.TransferNoFill
          If (.ReserveLevel < FillLevel) Then
            State = EState.TransferAndFill
          End If
          If Timer.Finished Then
            State = EState.Transfer1
          End If
          Status = commandName_ & Translate("Transfer") & Timer.ToString(1)


        Case EState.Transfer1
          If .ReserveLevel > 10 Then Timer.Seconds = .Parameters.AddTransferTimeBeforeRinse
          If Timer.Finished Then
            State = EState.Rinse
            Timer.Seconds = .Parameters.ReserveRinseTime
          End If
          If .ReserveLevel > 10 Then
            Status = commandName_ & Translate("Transferring") & (" ") & (.ReserveLevel / 10).ToString("#0.0") & "%"
          Else Status = commandName_ & Translate("Transferring") & Timer.ToString(1)
          End If


        Case EState.Rinse
          If Timer.Finished Then
            State = EState.Transfer2
            StateWas = State
            Timer.Seconds = MinMax(.Parameters.ReserveTimeAfterRinse, 5, 300)
          End If
          Status = commandName_ & Translate("Rinse to Machine") & Timer.ToString(1)


        Case EState.Transfer2
          If .ReserveLevel > 10 Then Timer.Seconds = .Parameters.ReserveTimeAfterRinse
          If Timer.Finished Then
            State = EState.RinseToDrain
            Timer.Seconds = MinMax(.Parameters.ReserveRinseToDrainTime, 5, 120)
            'Tell the control system to step on
            '   Return True
          End If
          If .ReserveLevel > 10 Then
            Status = commandName_ & (" ") & Translate("Transfer") & "(2) " & (.ReserveLevel / 10).ToString("#0.0") & "% " & Timer.ToString
          Else : Status = commandName_ & (" ") & Translate("Transfer") & "(2) " & Timer.ToString(1)
          End If


            '********************************************************************************************
            '******   WALL WASHING RINSE TO DRAIN
            '********************************************************************************************
        Case EState.RinseToDrainPause
          If Not pauseCommand Then
            If .Parent.IsPaused Then Status = Translate("Paused")
            If .EStop Then Status = Translate("Paused") & (", ESTOP")
            Timer.Seconds = 2
          End If
          If Timer.Finished Then
            State = EState.RinseToDrain
            Timer.Seconds = MinMax(.Parameters.ReserveRinseToDrainTime, 5, 120)
          End If
          Status = Translate("Paused") & Timer.ToString(1)


        Case EState.RinseToDrain
          If Timer.Finished Then
            NumberOfRinses -= 1
            State = EState.TransferToDrain
            Timer.Seconds = .Parameters.ReserveDrainTime
          End If
          ' Update Restart State
          StateWas = EState.RinseToDrain
          ' Pause conditions
          If pauseCommand Then State = EState.RinseToDrainPause
          Status = commandName_ & Translate("Rinsing To Drain") & Timer.ToString(1)


        Case EState.TransferToDrain
          If .ReserveLevel > 10 Then
            Status = commandName_ & Translate("Draining") & (" ") & (.ReserveLevel / 10).ToString("#0.0") & "% " & Timer.ToString
            Timer.Seconds = .Parameters.ReserveDrainTime
          Else
            Status = commandName_ & Translate("Draining") & Timer.ToString(1)
          End If
          If Timer.Finished Then
            If (NumberOfRinses > 0) Then
              State = EState.RinseToDrain
              Timer.Seconds = .Parameters.ReserveRinseToDrainTime
            Else
              State = EState.Done
              .ReserveReady = False
              .HD.HDCompleted = False
              Timer.Seconds = 2
            End If
          End If
          ' Update Restart State
          StateWas = EState.RinseToDrain
          ' Pause conditions
          If pauseCommand Then State = EState.RinseToDrainPause


        Case EState.Done
          Status = commandName_ & "Completing" & Timer.ToString(1)
          If Timer.Finished Then Cancel()


      End Select
    End With
  End Function
  Public Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
    StateWas = EState.Off
    Status = ""
  End Sub

  Public ReadOnly Property IsOn As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property

  Friend ReadOnly Property IsActive As Boolean
    Get
      Return (State >= EState.Fill) AndAlso (State <= EState.RinseToDrain)
    End Get
  End Property

  Friend ReadOnly Property IsFilling As Boolean
    Get
      Return (State = EState.Fill) OrElse (State = EState.TransferAndFill)
    End Get
  End Property

  ReadOnly Property IsInterlocked As Boolean
    Get
      Return (State = EState.Interlock)
    End Get
  End Property

  ' TODO ' need?
  ReadOnly Property IsTransferAndNoFilling As Boolean
    Get
      Return (State = EState.TransferNoFill)
    End Get
  End Property

  ReadOnly Property IsTransfer As Boolean
    Get
      Return (State >= EState.TransferAndFill) AndAlso (State <= EState.Transfer2)
    End Get
  End Property

  ReadOnly Property IsHeating As Boolean
    Get
      Return (State = EState.Heat)
    End Get
  End Property

  ReadOnly Property IsTransferToDrain As Boolean
    Get
      Return (State = EState.RinseToDrain) OrElse (State = EState.TransferToDrain)
    End Get
  End Property

  ReadOnly Property IsRinse As Boolean
    Get
      Return (State = EState.Rinse)
    End Get
  End Property

  ReadOnly Property IsRinseToDrain As Boolean
    Get
      Return (State = EState.RinseToDrain)
    End Get
  End Property

  ReadOnly Property IsPaused As Boolean
    Get
      Return (State = EState.Paused)
    End Get
  End Property

  Friend ReadOnly Property IsOverrun As Boolean
    Get
      Return IsOn AndAlso TimerOverrun.Finished
    End Get
  End Property
  Public ReadOnly Property TimeOverrun As Integer
    Get
      Return TimerOverrun.Seconds
    End Get
  End Property
  Public ReadOnly Property TimeAdvance As Integer
    Get
      Return TimerAdvance.Seconds
    End Get
  End Property

End Class
