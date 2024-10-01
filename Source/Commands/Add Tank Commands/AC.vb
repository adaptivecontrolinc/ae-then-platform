'American & Efird - Mt. Holly Then Platform
' Version 2024-08-25

Imports Utilities.Translations

<Command("Add By Contacts", "|0-99|", "", "", "3+'1"), TranslateCommand("es", "Agregar contactos", "|0-99|"),
Description("Transfer the Add Tank to the machine, over the defined number of batch contacts."),
  TranslateDescription("es", "Transferir el depósito añadir a la máquina, sobre el número definido de contactos por lotes."),
 Category("Add Tank Commands"), TranslateCategory("es", "Add Tank Commands")>
Public Class AC : Inherits MarshalByRefObject : Implements ACCommand
  Private Const commandName_ As String = ("AC: ")
  Private ReadOnly controlCode As ControlCode

  Public Enum EState
    Off
    WaitReady
    WaitForSafe

    Start
    Pause
    Restart

    ' Dosing using Add Pump and Add Rate (0-100%) valves
    Dose
    DoseMix

    ' Transfer Empty
    TransferEmpty1
    RinseToTransfer
    SettleTank
    TransferEmpty2
    RinseToTransferDose
    TransferEmptyDose

    ' Fill & Mix - Send to Drain
    DrainPause
    FillToMix
    MixForTime
    CloseMixForTime
    TransferDrain
    RinseToDrain
    TransferDrain2

    Done
  End Enum
  Public State As EState
  Public StateRestart As EState
  Public Status As String
  Public ContactsAdd As Integer
  Public ContactsStart As Integer
  Public ContactsRemain As Integer

  Public LevelActual As Integer
  Public LevelDesired As Integer
  Public LevelFill As Integer
  Public LevelStart As Integer
  Public NumberOfRinses As Integer

  ' Timers
  Public Timer As New Timer
  Public TimerMix As New Timer
  Public TimerOverrun As New Timer

  Private timerWaitReady_ As New TimerUp
  Public ReadOnly Property TimeWaitReady As Integer
    Get
      If (Not IsOn) OrElse (State > EState.WaitReady) Then
        timerWaitReady_.Stop()
      End If
      Return timerWaitReady_.Seconds
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

  Public Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With controlCode

      ' Cancel Commands
      .AT.Cancel()
      .KA.Cancel() : .WK.Cancel()
      .DR.Cancel() : .FI.Cancel() : .HC.Cancel() : .HD.Cancel() : .PR.Cancel()
      .RI.Cancel() : .RC.Cancel() : .RH.Cancel() : .TM.Cancel()
      .LD.Cancel() : .PH.Cancel() : .SA.Cancel() : .UL.Cancel()
      If .RT.IsForeground Then .RT.Cancel()
      .RW.Cancel()
      .WT.Cancel()

      ' TODO 
      ' KA Not cancelled in vb6
      ' RW cancelled
      ' .DrugroomPreview.WaitingAtTransferTimer.Start


      'If AddTank Not Prepared and AP not on, then start it
      If (Not .AddReady) OrElse (.KP.KP1.Param_Destination = EKitchenDestination.Add) Then
        If Not .AP.IsOn Then .AP.Start(0, 0)
      End If

      'Check array bounds just to be on the safe side
      ContactsAdd = 0
      If param.GetUpperBound(0) >= 1 Then ContactsAdd = param(1)

      'Set the default state
      State = EState.WaitReady
      NumberOfRinses = MinMax(.Parameters.AddRinses, 0, 10)

      ' Update Drugroom Preview
      timerWaitReady_.Start()

      'Set default time
      TimerOverrun.Minutes = (.Parameters.StandardTimeAddPrepare + 5) + (ContactsAdd * 2)

      'This is a foreground command so do not step on
      Return False
    End With
  End Function

  Public Function Run() As Boolean Implements ACCommand.Run
    With controlCode
      'In this command we will remember the state of each state machine when we start
      '  we can then loop through the state machine until all state changes have completed 
      '  this prevents dwelling on an intermediate state between control code runs which can occasionally cause IO flicker

      'Pause the state machine under these conditions
      Dim pauseCommand As Boolean = .Parent.IsPaused OrElse .EStop OrElse (Not .PumpControl.PumpRunning) OrElse (Not .TempSafe) OrElse (Not .MachineClosed)
      Dim pauseDose As Boolean = pauseCommand

      Static StartState As EState
      Do
        'Remember state and loop until state does not change
        'This makes sure we go through all the state changes before proceeding and setting IO
        StartState = State

        'Pause the state machine if necessary
        If (State > EState.Restart) AndAlso (State < EState.DrainPause) AndAlso pauseCommand Then
          State = EState.Pause
        End If
        ' Pause the state machine for rinsing to drain, if necessary
        If (State > EState.DrainPause) AndAlso (State < EState.Done) AndAlso (.Parent.IsPaused OrElse .EStop) Then
          State = EState.DrainPause
        End If


        '********************************************************************************************
        '******   STATE LOGIC
        '********************************************************************************************
        Select Case State
          Case EState.Off
            Status = (" ")


          Case EState.WaitReady
            ' Wait for Add Tank Ready
            Status = Translate("Wait Ready")
            If .AP.IsWaitReady Then Status = commandName_ & .AP.Status
            If .AF.IsActive Then Status = commandName_ & .AF.Status
            ' Kitchen Tank
            If .KA.IsActive AndAlso (.KA.Destination = EKitchenDestination.Add) Then Status = .KA.Status
            If .KP.KP1.IsForeground AndAlso (.KP.KP1.DispenseTank = EKitchenDestination.Add) Then Status = .KP.KP1.DrugroomDisplay
            If .LA.KP1.IsForeground AndAlso (.LA.KP1.Param_Destination = EKitchenDestination.Add) Then Status = .LA.KP1.DrugroomDisplay
            ' TODO : If .Parent.IsPaused Then .DrugroomPreview.WaitingAtTransferTimer.Pause
            ' Restart Drugroom Preview
            If Not .Parent.IsPaused AndAlso timerWaitReady_.IsPaused Then timerWaitReady_.Restart() ' .DrugroomPreview.WaitingAtTransferTimer.Restart
            ' Reset Ready Delay Timer
            If Not .AddReady Then
              Timer.Seconds = 5
            Else
              If Timer.Finished Then
                .AD.Cancel() : .AF.Cancel() : .AP.Cancel()
                ' TODO add low pressure request to PressureControl (Help with addpump transfer difficulty)
                State = EState.WaitForSafe
                timerWaitReady_.Pause()
                Timer.Seconds = 2
              End If
            End If
            ' Update Restart State
            StateRestart = EState.WaitReady


          Case EState.WaitForSafe
            Status = commandName_ & Translate("Wait for Safe") & Timer.ToString(1)
            If Not (.TempSafe OrElse .HD.HDCompleted) Then Status = commandName_ & Translate("Temp Not Safe") : Timer.Seconds = 1
            If Not .MachineClosed Then Status = commandName_ & Translate("Machine Not Closed") : Timer.Seconds = 1
            If .Parent.IsPaused Then Status = commandName_ & Translate("Paused") : Timer.Seconds = 1
            If .EStop Then Status = commandName_ & Translate("EStop") : Timer.Seconds = 1
            If Timer.Finished Then State = EState.Start
            ' Update Restart State
            StateRestart = EState.WaitReady


          Case EState.Start
            .AP.Cancel() : .AF.Cancel()
            ' TODO VB6: .DrugroomPreview.WaitingAtTransferTimer.Pause
            If ContactsAdd > 0 Then
              State = EState.DoseMix
              LevelStart = .AddLevel
              ContactsStart = .TotalNumberOfContacts
            Else
              .AddControl.Cancel()
              State = EState.TransferEmpty1
              Timer.Seconds = .Parameters.AddTransferTimeBeforeRinse + 1
            End If
            ' Update Restart State
            StateRestart = EState.WaitReady
            Status = commandName_ & Translate("Starting")


          Case EState.Pause
            Status = commandName_ & Translate("Paused")
            If Not .PumpControl.PumpRunning Then Status = commandName_ & Translate("Pump Off") : Timer.Seconds = 2
            If Not (.TempSafe OrElse .HD.HDCompleted) Then Status = commandName_ & Translate("Temp Not Safe") : Timer.Seconds = 2
            If Not .MachineClosed Then Status = commandName_ & Translate("Machine Not Closed") : Timer.Seconds = 2
            If .EStop Then Status = commandName_ & Translate("EStop") : Timer.Seconds = 2
            If Not pauseDose AndAlso Timer.Finished Then State = EState.Restart


          Case EState.Restart
            If (ContactsAdd > 0) AndAlso (.AddLevel > LevelDesired) Then
              State = EState.DoseMix
              Timer.Restart()
              'TODO add restart dose timer using last paused time remaining
            Else
              State = StateRestart
              Timer.Seconds = 5
            End If
            Status = commandName_ & Translate("Restarting") & Timer.ToString(1)


            '********************************************************************************************
            '******   DOSING THROUGH ADD PUMP
            '********************************************************************************************

          Case EState.Dose
            LevelDesired = Setpoint(.TotalNumberOfContacts)
            ' Check level - if level low switch to circulate
            If (.AddLevel < LevelDesired) Then State = EState.DoseMix
            ' Calculate remaining contacts
            ContactsRemain = ContactsAdd - (.TotalNumberOfContacts - ContactsStart)
            ' Finished dosing, transfer empty
            If (ContactsRemain = 0) OrElse (.AddLevel <= 10) Then
              State = EState.TransferEmpty1
              Timer.Seconds = .Parameters.AddTransferTimeBeforeRinse
            End If
            ' Update Restart State
            StateRestart = EState.DoseMix
            'Status = Translate("Dosing") & (" ") & (.AddLevel / 10).ToString("#0.0") & " / " & (LevelDesired / 10).ToString("#0.0") & "%"
            Status = Translate("Dosing") & (" ") & ContactsRemain.ToString("#0") & " / " & ContactsAdd.ToString("#0") & " Contacts"


          Case EState.DoseMix
            LevelDesired = Setpoint(.TotalNumberOfContacts)
            ' Check level - if level high then switch to transfer
            If (.AddLevel >= LevelDesired) AndAlso (.PumpControl.IsInToOut) Then State = EState.Dose
            ' Calculate remaining contacts
            ContactsRemain = ContactsAdd - (.TotalNumberOfContacts - ContactsStart)
            ' Finished dosing, transfer empty
            If (ContactsRemain <= 0) OrElse (.AddLevel < 10) Then
              State = EState.TransferEmpty1
              Timer.Seconds = .Parameters.AddTransferTimeBeforeRinse + 1
            End If
            ' Update Restart State
            StateRestart = EState.DoseMix
            ' Status = Translate("Mixing") & (" ") & (.AddLevel / 10).ToString("#0.0") & " / " & (LevelDesired / 10).ToString("#0.0") & "%"
            Status = commandName_ & Translate("Dose Paused") & (" ") & ContactsRemain.ToString("#0") & " / " & ContactsAdd.ToString("#0") & " Contacts"



            '********************************************************************************************
            '******   TRANSFERRING EMPTY THROUGH MAIN ADD TRANSFER LINE
            '********************************************************************************************
          Case EState.TransferEmpty1
            If .AddLevel > 10 Then Timer.Seconds = .Parameters.AddTransferTimeBeforeRinse + 1
            If Timer.Finished Then
              State = EState.RinseToTransfer
              Timer.Seconds = MinMax(.Parameters.AddTimeToRinseToMachine, 5, 120)
            End If
            Status = commandName_ & Translate("Transferring") & Timer.ToString(1)
            If .AddLevel > 10 Then Status = commandName_ & Translate("Transferring") & (" ") & (.AddLevel / 10).ToString("#0.0") & "%"
            ' Update Restart State
            StateRestart = EState.TransferEmpty1


          Case EState.RinseToTransfer
            If Timer.Finished Then
              State = EState.SettleTank
              Timer.Seconds = .Parameters.AddTransferSettleTime
            End If
            ' Update Restart State
            StateRestart = EState.TransferEmpty1
            Status = commandName_ & Translate("Rinse to machine") & Timer.ToString(1)


          Case EState.SettleTank
            If Timer.Finished Then
              State = EState.TransferEmpty2
              Timer.Seconds = .Parameters.AddTimeAfterRinse
            End If
            ' Update Restart State
            StateRestart = EState.TransferEmpty1
            Status = commandName_ & Translate("Settle Tank") & Timer.ToString(1)


          Case EState.TransferEmpty2
            If .AddLevel > 10 Then Timer.Seconds = .Parameters.AddTimeAfterRinse
            If Timer.Finished Then
              State = EState.RinseToTransferDose
              Timer.Seconds = .Parameters.AddTimeToRinseToMachine
            End If
            Status = commandName_ & Translate("Transferring") & Timer.ToString(1)
            If .AddLevel > 10 Then
              Status = commandName_ & Translate("Transferring") & (" ") & (.AddLevel / 10).ToString("#0.0") & "%"
            End If
            ' Update Restart State
            StateRestart = EState.TransferEmpty1


          Case EState.RinseToTransferDose
            If Timer.Finished Then
              State = EState.TransferEmptyDose
              Timer.Seconds = .Parameters.AddTimeAfterRinse
            End If
            Status = commandName_ & Translate("Rinse to machine") & Timer.ToString(1)
            ' Update Restart State
            StateRestart = EState.TransferEmpty1


          Case EState.TransferEmptyDose
            If .AddLevel > 10 Then Timer.Seconds = .Parameters.AddTimeAfterRinse
            If Timer.Finished Then
              .AddReady = False
              .AP.Cancel() : .AF.Cancel()
              State = EState.FillToMix
              Return True  'Step On
            End If
            Status = commandName_ & Translate("Transfer Rinse") & Timer.ToString(1)
            If .AddLevel > 10 Then Status = commandName_ & Translate("Transfer Rinse") & (" ") & (.AddLevel / 10).ToString("#0.0") & "%"
            ' Update Restart State
            StateRestart = EState.TransferEmpty1


            '********************************************************************************************
            '******   FILL, CIRCULATE, AND SEND TO DRAIN
            '********************************************************************************************
          Case EState.DrainPause
            If Not (.EStop OrElse .Parent.IsPaused) Then State = StateRestart
            Status = commandName_ & Translate("Paused")
            If .EStop Then Status = commandName_ & Translate("EStop")


          Case EState.FillToMix
            LevelFill = MinMax(.Parameters.AddRinseFillLevel, .Parameters.AddMixOnLevel, 750)
            If .AddLevel > LevelFill Then
              State = EState.MixForTime
              Timer.Seconds = MinMax(.Parameters.AddRinseMixTime, 5, 120)
              TimerMix.Seconds = .Parameters.AddRinseMixPulseTime
            End If
            ' Update Restart State
            StateRestart = EState.FillToMix
            Status = commandName_ & Translate("Filling") & (" ") & (.AddLevel / 10).ToString("#0.0") & (" / ") & (LevelFill / 10).ToString("#0.0") & "% "


          Case EState.MixForTime
            If TimerMix.Finished Then
              State = EState.CloseMixForTime
              TimerMix.Seconds = 1
            End If
            If Timer.Finished Then
              State = EState.TransferDrain
              Timer.Seconds = MinMax(.Parameters.AddTimeToDrain, 5, 120)
            End If
            ' Update Restart State
            StateRestart = EState.FillToMix
            Status = commandName_ & Translate("Clean Mix") & Timer.ToString(1)

          Case EState.CloseMixForTime
            If TimerMix.Finished Then
              TimerMix.Seconds = .Parameters.AddRinseMixPulseTime
              State = EState.MixForTime
            End If
            If Timer.Finished Then
              State = EState.TransferDrain
              Timer.Seconds = .Parameters.AddTimeToDrain
            End If
            Status = commandName_ & Translate("Clean Mix") & Timer.ToString(1)
            ' Update Restart State
            StateRestart = EState.TransferDrain


          Case EState.TransferDrain
            If .AddLevel > 10 Then Timer.Seconds = .Parameters.AddTimeToDrain
            If Timer.Finished Then
              .AddReady = False
              State = EState.RinseToDrain
              Timer.Seconds = MinMax(.Parameters.AddTimeRinseToDrain, 5, 120)
            End If
            ' Update Restart State
            StateRestart = EState.RinseToDrain
            If .AddLevel > 10 Then
              Status = commandName_ & Translate("Draining") & (" ") & (.AddLevel / 10).ToString("#0.0") & "%"
            Else
              Status = commandName_ & Translate("Draining") & Timer.ToString(1)
            End If

            '********************************************************************************************
            '******   WALL WASHING RINSE TO DRAIN
            '********************************************************************************************
          Case EState.RinseToDrain
            If Timer.Finished Then
              NumberOfRinses -= 1
              State = EState.TransferDrain2
              Timer.Seconds = .Parameters.AddTimeToDrain
            End If
            Status = commandName_ & Translate("Rinse To Drain") & Timer.ToString(1)
            ' Update Restart State
            StateRestart = EState.RinseToDrain


          Case EState.TransferDrain2
            If .AddLevel > 10 Then Timer.Seconds = .Parameters.AddTimeToDrain
            If Timer.Finished Then
              If (NumberOfRinses > 0) Then
                State = EState.RinseToDrain
                Timer.Seconds = MinMax(.Parameters.AddTimeRinseToDrain, 5, 120)
              Else
                State = EState.Done
                Timer.Seconds = 2
              End If
            End If
            Status = commandName_ & Translate("Draining") & Timer.ToString(1)
            If .AddLevel > 10 Then Status = commandName_ & Translate("Draining") & (" ") & (.AddLevel / 10).ToString("#0.0") & "%"
            ' Update Restart State
            StateRestart = EState.RinseToDrain


          Case EState.Done
            If Timer.Finished Then Cancel()
            Status = commandName_ & Translate("Completing") & Timer.ToString(1)

        End Select
      Loop Until (StartState = State)
    End With
  End Function

  Public Sub Cancel() Implements ACCommand.Cancel
    ContactsAdd = 0
    ContactsRemain = 0
    ContactsStart = 0
    LevelFill = 0
    State = EState.Off
    StateRestart = EState.Off
    timerWaitReady_.Stop()
    Timer.Cancel()
    TimerOverrun.Cancel()
  End Sub

  Public ReadOnly Property IsOn() As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property

  Public ReadOnly Property IsActive As Boolean
    Get
      Return (State > EState.WaitReady) AndAlso (State < EState.Done)
    End Get
  End Property

  Public ReadOnly Property IsForeground As Boolean
    Get
      Return (State > EState.Off) AndAlso (State <= EState.TransferEmptyDose)
    End Get
  End Property

  Public ReadOnly Property IsBackground() As Boolean
    Get
      Return (State >= EState.DrainPause) AndAlso (State <= EState.Done)
    End Get
  End Property

  Public ReadOnly Property IsWaitReady As Boolean
    Get
      Return (State = EState.WaitReady)
    End Get
  End Property

  Public ReadOnly Property IsOverrun() As Boolean
    Get
      Return (State = EState.WaitReady) AndAlso TimerOverrun.Finished
    End Get
  End Property

  Public ReadOnly Property IsDosing() As Boolean
    Get
      Return (State = EState.Dose) OrElse (State = EState.DoseMix)
    End Get
  End Property

  Public ReadOnly Property IsLowPressure() As Boolean
    Get
      Return (State >= EState.Start) AndAlso (State <= EState.TransferEmptyDose)
    End Get
  End Property


#Region " IO PROPERTIES "

  Public ReadOnly Property IoAddLamp As Boolean
    Get
      Return (State = EState.WaitReady)
    End Get
  End Property

  Public ReadOnly Property IoAddFill As Boolean
    Get
      Return (State = EState.RinseToTransfer) OrElse (State = EState.RinseToTransferDose) OrElse
             (State = EState.FillToMix) OrElse (State = EState.RinseToDrain)
    End Get
  End Property

  Public ReadOnly Property IoAddTransfer As Boolean
    Get
      Return (State = EState.Dose) OrElse
             ((State >= EState.TransferEmpty1) AndAlso (State <= EState.TransferEmptyDose))
    End Get
  End Property

  Public ReadOnly Property IoAddMix As Boolean
    Get
      Return (State = EState.DoseMix) OrElse (State = EState.FillToMix) OrElse (State = EState.MixForTime)
    End Get
  End Property

  Public ReadOnly Property IoAddDrain() As Boolean
    Get
      Return (State = EState.TransferDrain) OrElse (State = EState.RinseToDrain) OrElse (State = EState.TransferDrain2)
    End Get
  End Property

  Public ReadOnly Property IoAddPump() As Boolean
    Get
      Return ((State >= EState.Dose) AndAlso (State <= EState.TransferEmptyDose)) OrElse
             ((State >= EState.MixForTime) AndAlso (State <= EState.TransferDrain))
    End Get
  End Property

#End Region

  'Function copied from AT command, but without the curve....dosed linear over the number of contacts
  Private Function Setpoint(NumberOfContacts As Integer) As Integer
    'If timer has finished exit function
    If ContactsRemain = 0 Then
      Setpoint = 0
      Exit Function
    End If

    'Amount we should have transferred so far
    Dim dblPercentComplete As Double, DesiredAmount As Double
    If (ContactsAdd <> 0) Then
      dblPercentComplete = ContactsRemain / ContactsAdd
      DesiredAmount = LevelStart * dblPercentComplete
    Else
      DesiredAmount = 0
    End If

    'Calculate setpoint and limit to 0-1000
    Setpoint = CInt(DesiredAmount)
    If Setpoint < 0 Then Setpoint = 0
    If Setpoint > 1000 Then Setpoint = 1000

    'Global variable for display purposes - yuk!!
    LevelDesired = Setpoint
  End Function

End Class
