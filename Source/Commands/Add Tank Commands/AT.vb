'American & Efird - Mt. Holly Then Platform
' Version 2024-08-25

Imports Utilities.Translations

<Command("Add Transfer", "|0-99|mins Curve:|0-9|", "", "", "'1 + 5"),
TranslateCommand("es", "Agregar transferencia", "|0-99|mins Curve:|0-9|"),
Description("Transfer the add tank to the machine over the specified time and curve."),
TranslateDescription("es", "Transferir el deposito de complemento a la m�quina durante el tiempo especificado y la curva."),
Category("Add Tank Commands"), TranslateCategory("es", "Add Tank Commands")>
Public Class AT : Inherits MarshalByRefObject : Implements ACCommand
  Private Const commandName_ As String = ("AT: ")
  Private ReadOnly controlCode As ControlCode

  Public Enum EState
    Off
    WaitReady
    WaitForSafe

    Start
    Pause
    Restart

    ' Dosing
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
  Public AddTime As Integer
  Public AddCurve As Integer

  Public LevelActual As Integer
  Public LevelDesired As Integer
  Public LevelFill As Integer
  Public LevelStart As Integer
  Public NumberOfRinses As Integer

  Public DosePausedTimeRemaining As Integer
  Friend Property DosePauseTimer As New Timer

  Property Timer As New Timer
  Friend Property TimerMix As New Timer
  Public ReadOnly Property TimeMix As Integer
    Get
      Return TimerMix.Seconds
    End Get
  End Property
  Friend Property TimerOverrun As New Timer
  Public ReadOnly Property TimeOverrun As Integer
    Get
      Return TimerOverrun.Seconds
    End Get
  End Property
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
      .AC.Cancel() ': .AT.Cancel()
      .KA.Cancel() : .WK.Cancel()
      .DR.Cancel() : .FI.Cancel() : .HC.Cancel() : .HD.Cancel() : .PR.Cancel()
      .RI.Cancel() : .RC.Cancel() : .RH.Cancel() : .TM.Cancel()
      .LD.Cancel() : .PH.Cancel() : .SA.Cancel() : .UL.Cancel()
      If .RT.IsForeground Then .RT.Cancel()
      .RW.Cancel()
      '  .CO.Cancel() : .HC.Cancel() : .HE.Cancel() : 
      .WT.Cancel()


      'Check array bounds just to be on the safe side
      If param.GetUpperBound(0) >= 1 Then AddTime = param(1) * 60
      If param.GetUpperBound(0) >= 2 Then AddCurve = param(2)

      'Set the default state
      State = EState.WaitReady
      NumberOfRinses = MinMax(.Parameters.AddRinses, 0, 10)
      Timer.Seconds = 2

      ' Update Drugroom Preview
      timerWaitReady_.Start()

      'Set default time
      TimerOverrun.Seconds = (.Parameters.StandardTimeAddPrepare * 60) + AddTime

      'This is a foreground command so do not step on
      Return False
    End With
  End Function

  Public Function Run() As Boolean Implements ACCommand.Run
    With controlCode
      'In this command we will remember the state of each state machine when we start
      '  we can then loop through the state machine until all state changes have completed 
      '  this prevents dwelling on an intermediate state between control code runs which can occasionally cause IO flicker

      'Remember the volume at the beginning of the transfer so we can update the machine transfer volume
      'Static av1TransferVolume As Integer

      'Pause the state machine under these conditions
      Dim pauseCommand As Boolean = .Parent.IsPaused OrElse .EStop OrElse (Not .PumpControl.PumpRunning) OrElse (Not .TempSafe) OrElse (Not .MachineClosed)
      Dim pauseDose As Boolean = pauseCommand

      Static StartState As EState
      Do
        'Remember state and loop until state does not change
        'This makes sure we go through all the state changes before proceeding and setting IO
        StartState = State

        ' Pause the state machine if necessary
        If (State > EState.Restart) AndAlso (State < EState.DrainPause) AndAlso pauseCommand Then
          DosePausedTimeRemaining = Timer.Seconds
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
            ' Restart Drugroom Preview
            If Not .Parent.IsPaused AndAlso timerWaitReady_.IsPaused Then timerWaitReady_.Restart()
            If .AP.IsOn Then
              Status = commandName_ & .AP.Status
              If Not .AddReady Then Timer.Seconds = 10
            End If
            If .AF.IsOn Then
              Status = commandName_ & .AF.Status
              Timer.Seconds = 10
            End If
            If (.KP.KP1.IsOn AndAlso (.KP.KP1.DispenseTank = 1)) Then
              Status = commandName_ & "Drugroom " & .KP.KP1.Status
              If Not .AddReady Then Timer.Seconds = 10
            End If
            If (.LA.KP1.IsOn AndAlso (.LA.KP1.DispenseTank = 1)) Then
              Status = commandName_ & "Drugroom " & .LA.KP1.Status
              If Not .AddReady Then Timer.Seconds = 10
            End If
            If .Parent.IsPaused Then
              .DrugroomPreview.WaitingAtTransferTimer.Pause()
            End If
            If (Not .Parent.IsPaused) And .DrugroomPreview.WaitingAtTransferTimer.IsPaused Then
              .DrugroomPreview.WaitingAtTransferTimer.Restart()
            End If
            If Timer.Finished Then
              .DrugroomPreview.WaitingAtTransferTimer.Pause()
              .AD.Cancel() : .AF.Cancel() : .AP.Cancel()
              ' TODO add low pressure request to PressureControl (Help with addpump transfer difficulty)
              State = EState.WaitForSafe
              timerWaitReady_.Pause()
              Timer.Seconds = 2
            End If
            ' Update Status
            Status = Translate("Wait Ready")
            If .AP.IsWaitReady Then Status = .AP.Status
            If .AF.IsActive Then Status = .AF.Status
            ' Kitchen Tank
            If .KA.IsActive AndAlso (.KA.Destination = EKitchenDestination.Add) Then Status = .KA.Status
            If .KP.KP1.IsForeground AndAlso (.KP.KP1.Param_Destination = EKitchenDestination.Add) Then Status = .KP.KP1.DrugroomDisplay
            If .LA.KP1.IsForeground AndAlso (.LA.KP1.Param_Destination = EKitchenDestination.Add) Then Status = .LA.KP1.DrugroomDisplay
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
            .AP.Cancel() : .AF.Cancel() : .AddControl.Cancel()
            If AddTime > 0 Then
              State = EState.DoseMix
              DosePauseTimer.Milliseconds = MinMax(.Parameters.AddTransferDoseMinTimeMs, 1000, 10000)
              Timer.Seconds = AddTime
              LevelStart = .AddLevel
            Else
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
            LevelDesired = Setpoint()
            If Timer.Finished Then
              State = StateRestart
              Timer.Seconds = 5
            End If
            Status = commandName_ & Translate("Restarting") & Timer.ToString(1)


            '********************************************************************************************
            '******   DOSING THROUGH ADD PUMP
            '********************************************************************************************

          Case EState.Dose
            LevelDesired = Setpoint()
            ' Check level - if level low switch to circulate
            If (.AddLevel < LevelDesired) AndAlso DosePauseTimer.Finished Then
              State = EState.DoseMix
              DosePauseTimer.Milliseconds = MinMax(.Parameters.AddTransferDoseMinTimeMs, 500, 10000)
            End If
            ' If we're finished transfer empty
            If Timer.Finished OrElse (.AddLevel <= 10) Then
              State = EState.TransferEmpty1
              Timer.Seconds = .Parameters.AddTransferTimeBeforeRinse
            End If
            ' Update Restart State
            StateRestart = EState.DoseMix
            Status = commandName_ & Translate("Dosing") & (" ") & (.AddLevel / 10).ToString("#0.0") & " / " & (LevelDesired / 10).ToString("#0.0") & "%"


          Case EState.DoseMix
            LevelDesired = Setpoint()
            ' Check level - if high switch to dose
            If (.AddLevel >= LevelDesired) AndAlso DosePauseTimer.Finished Then
              State = EState.Dose
              DosePauseTimer.Milliseconds = MinMax(.Parameters.AddTransferDoseMinTimeMs, 500, 10000)
            End If
            ' Finished dosing, transfer empty
            If Timer.Finished OrElse (.AddLevel < 10) Then
              State = EState.TransferEmpty1
              Timer.Seconds = .Parameters.AddTransferTimeBeforeRinse
            End If
            ' Update Restart State
            StateRestart = EState.DoseMix
            Status = commandName_ & Translate("Mixing") & (" ") & (.AddLevel / 10).ToString("#0.0") & " / " & (LevelDesired / 10).ToString("#0.0") & "%"



            '********************************************************************************************
            '******   TRANSFERRING EMPTY THROUGH MAIN ADD TRANSFER LINE
            '********************************************************************************************
          Case EState.TransferEmpty1
            If .AddLevel > 10 Then Timer.Seconds = .Parameters.AddTransferTimeBeforeRinse
            If Timer.Finished Then
              State = EState.RinseToTransfer
              Timer.Seconds = MinMax(.Parameters.AddTimeToRinseToMachine, 5, 120)
            End If
            ' Update Restart State
            StateRestart = EState.TransferEmpty1
            If .AddLevel > 10 Then
              Status = commandName_ & Translate("Transferring") & (" ") & (.AddLevel / 10).ToString("#0.0") & "%"
            Else Status = commandName_ & Translate("Transferring") & Timer.ToString(1)
            End If


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
            ' Update Restart State
            StateRestart = EState.RinseToDrain
            Status = commandName_ & Translate("Draining") & Timer.ToString(1)
            If .AddLevel > 10 Then Status = commandName_ & Translate("Draining") & (" ") & (.AddLevel / 10).ToString("#0.0") & "%"


          Case EState.Done
            If Timer.Finished Then Cancel()
            Status = commandName_ & Translate("Completing") & Timer.ToString(1)

        End Select
      Loop Until (StartState = State)
    End With
  End Function

  Public Sub Cancel() Implements ACCommand.Cancel
    AddTime = 0
    AddCurve = 0
    LevelFill = 0
    State = EState.Off
    StateRestart = EState.Off
    Status = ""
    Timer.Cancel()
    TimerOverrun.Cancel()
    timerWaitReady_.Stop()
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

#Region " Method "
  'Function uses the expressions (y = x * x) and (y = sqr(x)) to generate curves for
  'dosing control.
  'x = A fraction between 0 and 1 (I know) representing elapsed time
  'y = A fraction between 0 and 1 representing the amount we should have transferred so far
  '
  'The curves are scaled by adding a percent of the difference between the curve value and
  'the equivalent linear value to the original curve value (??).
  Private Function Setpoint() As Integer
    'If timer has finished exit function
    If Timer.Finished Then
      Setpoint = 0
      Exit Function
    End If

    'Amount we should have transferred so far
    Dim ElapsedTime As Double, LinearTerm As Double, TransferAmount As Double
    If AddTime = 0 Then AddTime = 1
    ElapsedTime = (AddTime - Timer.Seconds) / AddTime
    LinearTerm = ElapsedTime
    TransferAmount = LevelStart * LinearTerm

    'Calculate scaling factor (0-1) for progressive and digressive curves
    If AddCurve > 0 Then
      Dim ScalingFactor As Double
      ScalingFactor = (10 - AddCurve) / 10
      'Calculate term for progressive transfer (0-1) if odd curve
      If (AddCurve Mod 2) = 1 Then
        Dim MaxOddCurve As Double, OddTerm As Double
        MaxOddCurve = (ElapsedTime * ElapsedTime)
        OddTerm = MaxOddCurve + ((LinearTerm - MaxOddCurve) * ScalingFactor)
        TransferAmount = LevelStart * OddTerm
      Else
        'Calculate term for digressive transfer (0-1) if even curve
        Dim MaxEvenCurve As Double, EvenTerm As Double
        MaxEvenCurve = Math.Sqrt(ElapsedTime)
        EvenTerm = MaxEvenCurve - ((MaxEvenCurve - LinearTerm) * ScalingFactor)
        TransferAmount = LevelStart * EvenTerm
      End If
    End If

    'Calculate setpoint and limit to 0-1000
    Setpoint = LevelStart - CInt(TransferAmount)
    If Setpoint < 0 Then Setpoint = 0
    If Setpoint > 1000 Then Setpoint = 1000

    'Global variable for display purposes - yuk!!
    LevelDesired = Setpoint
  End Function

#End Region


#If 0 Then

  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Implements ACCommand
Public Enum ATState
  ATOff
  ATWaitReady
  ATPause
  ATDose
  ATDosePause
  ATTransfer
  ATTransferEmpty1
  ATRinse
  ATSettleTank
  ATTransferEmpty2
  ATFillForMix
  ATMixForTime
  ATCloseMixForTime
  ATTransferToDrain
End Enum
Public State As ATState
Public StateString As String
Public Timer As New acTimer
Public MixTimer As New acTimer
Public AddTime As Long
Public AddCurve As Long
Public StartLevel As Long
Public DesiredLevel As Long
Public OverrunTimer As New acTimer

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Parameters=|0-99|m Curve:|0-9|\r\nName=Add Transfer\r\nMinutes=(3+'1)\r\nHelp=Doses the contents of the side tank over the time specified, using one of ten curves. Curve 0 is linear, odd numbers are progressive adds and even numbers are regressive adds "
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  
  If Param(1) > 0 Then
    AddTime = Param(1) * 60
  Else: AddTime = 0
  End If
  AddCurve = Param(2)
  
  With ControlCode
    .AC.ACCommand_Cancel
    .DR.ACCommand_Cancel: .FI.ACCommand_Cancel: .HD.ACCommand_Cancel
    .HC.ACCommand_Cancel: .LD.ACCommand_Cancel: .PH.ACCommand_Cancel
    .RC.ACCommand_Cancel: .RH.ACCommand_Cancel: .RI.ACCommand_Cancel
    .RT.ACCommand_Cancel: .RW.ACCommand_Cancel: .SA.ACCommand_Cancel
    .TM.ACCommand_Cancel: .UL.ACCommand_Cancel: .WK.ACCommand_Cancel
    .WT.ACCommand_Cancel
    .DrugroomPreview.WaitingAtTransferTimer.Start
  End With
  
  Timer = 10
  OverrunTimer = 300
  State = ATWaitReady
  
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  
  With ControlCode
    Select Case State
      
      Case ATOff
        StateString = ""
        
      Case ATWaitReady
        'Local Add
        If .AP.IsOn Then
          StateString = "AT: " & .AP.StateString
          If Not .AddReady Then Timer = 10
        End If
        If .AF.IsOn Then
          StateString = "AT: " & .AF.StateString
          Timer = 10
        End If
        'Kitchen Prepare
        If (.KA.IsActive And (.KA.Destination = 65)) Then
          StateString = "AT: Drugroom " & .KA.StateString
          If Not .AddReady Then Timer = 10
        End If
        If (.KP.KP1.IsOn And (.KP.KP1.DispenseTank = 1)) Then
          StateString = "AT: Drugroom " & .KP.KP1.StateString
          If Not .AddReady Then Timer = 10
        End If
        If (.LA.KP1.IsOn And (.LA.KP1.DispenseTank = 1)) Then
          StateString = "AT: Drugroom " & .LA.KP1.StateString
          If Not .AddReady Then Timer = 10
        End If
        If .Parent.IsPaused Then
          .DrugroomPreview.WaitingAtTransferTimer.Pause
        End If
        If (Not .Parent.IsPaused) And .DrugroomPreview.WaitingAtTransferTimer.IsPaused Then
          .DrugroomPreview.WaitingAtTransferTimer.Restart
        End If
        If Timer.Finished Then
          .DrugroomPreview.WaitingAtTransferTimer.Pause
          If AddTime = 0 Then
            .AF.AddMixing = False
            State = ATTransferEmpty1
            Timer = .Parameters_AddTransferTimeBeforeRinse
          Else
            .AF.AddMixing = False
            State = ATDosePause
            Timer = AddTime
            StartLevel = .AdditionLevel
          End If
        End If
        
      Case ATPause
        If Not .LidLocked Then
          StateString = "AT: Lid Not Locked "
        ElseIf Not .IO_MainPumpRunning Then
          StateString = "AT: Pump Not Running "
        Else: StateString = "AT: Paused "
        End If
        Timer.Pause
        If (Not .Parent.IsPaused) And (Not .IO_EStop_PB) And .IO_MainPumpRunning And .LidLocked Then
          If AddTime = 0 Then
            .AF.AddMixing = False
            State = ATTransferEmpty1
            Timer = .Parameters_AddTransferTimeBeforeRinse
          Else
            State = ATDosePause
            Timer.Restart
          End If
        End If
        
      Case ATDose
        StateString = "AT: Dosing Addition " & Pad(.AdditionLevel / 10, "0", 3) & "% / " & Pad(DesiredLevel / 10, "0", 3) & "%"
        'Check level - if level low switch to circulate
        If .AdditionLevel < Setpoint() Then State = ATDosePause
        'Check level - if level high then switch to transfer
        If .AdditionLevel > (Setpoint() + 30) Then State = ATTransfer
        'If pump not running or lid not locked go to pause state
        If .Parent.IsPaused Or .IO_EStop_PB Or Not (.IO_MainPumpRunning And .LidLocked) Then State = ATPause
        'If we're finished transfer empty
        If Timer.Finished Then
          DesiredLevel = 0
          State = ATTransferEmpty1
          Timer = .Parameters_AddTransferTimeBeforeRinse
        End If
      
      Case ATTransfer
        StateString = "AT: Dosing Addition " & Pad(.AdditionLevel / 10, "0", 3) & "% / " & Pad(DesiredLevel / 10, "0", 3) & "%"
        'Check level - if level low switch to circulate
        If .AdditionLevel < Setpoint() Then State = ATDosePause
        'If pump not running go to pause state
        If Not (.IO_MainPumpRunning And .LidLocked) Then State = ATPause
        'If we're finished transfer empty
        If Timer.Finished Then
          DesiredLevel = 0
          State = ATTransferEmpty1
          Timer = .Parameters_AddTransferTimeBeforeRinse
        End If

      Case ATDosePause
        StateString = "AT: Dose Paused " & Pad(.AdditionLevel / 10, "0", 3) & "% / " & Pad(DesiredLevel / 10, "0", 3) & "%"
        'Check level - if high switch to dose
        If .AdditionLevel > (Setpoint() + 30) Then State = ATDose
        'If pump not running go to pause state
        If Not (.IO_MainPumpRunning And .LidLocked) Then State = ATPause
        'If we're finished transfer empty
        If Timer.Finished Then
          State = ATTransferEmpty1
          Timer = .Parameters_AddTransferTimeBeforeRinse
        End If
      
      Case ATTransferEmpty1
        If .AdditionLevel > 10 Then
          StateString = "AT: Transfer To Machine " & Pad(.AdditionLevel / 10, "0", 3) & "% "
          Timer = .Parameters_AddTransferTimeBeforeRinse
        Else: StateString = "AT: Tranfering to Machine " & TimerString(Timer.TimeRemaining)
        End If
        'If pump not running go to pause state
        If Not (.IO_MainPumpRunning And .LidLocked) Then State = ATPause
        If Timer.Finished Then
          State = ATRinse
          Timer = .Parameters_AddTransferRinseTime
        End If
        
      Case ATRinse
        StateString = "AT: Rinsing Addition " & TimerString(Timer.TimeRemaining)
        If Timer.Finished Then
          State = ATSettleTank
          Timer = .Parameters_AddTransferSettleTime
        End If
         
      Case ATSettleTank
        StateString = "AT: Settle Addition " & TimerString(Timer.TimeRemaining)
        If Timer.Finished Then
          State = ATTransferEmpty2
          Timer = .Parameters_AddTransferTimeAfterRinse
        End If
         
      Case ATTransferEmpty2
        If .AdditionLevel > 10 Then
          StateString = "AT: Transfer Rinse To Machine " & Pad(.AdditionLevel / 10, "0", 3) & "% "
          Timer = .Parameters_AddTransferTimeAfterRinse
        Else: StateString = "AT: Tranfer Rinse to Machine " & TimerString(Timer.TimeRemaining)
        End If
        'If pump not running go to pause state
        If Not (.IO_MainPumpRunning Or .LidLocked) Then State = ATPause
        If Timer.Finished Then
          .AddReady = False
          .AF.ACCommand_Cancel
          State = ATFillForMix
        End If
      
      Case ATFillForMix
        StateString = "AT: Fill To Clean " & Pad(.AdditionLevel / 10, "0", 3) & "% / " & Pad(.Parameters_AddRinseFillLevel / 10, "0", 3) & "% "
        If .AdditionLevel > .Parameters_AddRinseFillLevel Then
          State = ATMixForTime
          Timer = .Parameters_AddRinseMixTime
          MixTimer = .Parameters_AddRinseMixPulseTime
        End If
        
      Case ATMixForTime
        StateString = "AT: Clean Mix " & TimerString(Timer.TimeRemaining)
        If MixTimer.Finished Then
          MixTimer = 1
          State = ATCloseMixForTime
        End If
        If Timer.Finished Then
          Timer = .Parameters_AddTransferToDrainTime
          State = ATTransferToDrain
        End If
      
      Case ATCloseMixForTime
        StateString = "AT: Clean Mix " & TimerString(Timer.TimeRemaining)
        If MixTimer.Finished Then
          MixTimer = .Parameters_AddRinseMixPulseTime
          State = ATMixForTime
        End If
        If Timer.Finished Then
          Timer = .Parameters_AddTransferToDrainTime
          State = ATTransferToDrain
        End If
      
      Case ATTransferToDrain
        If .AdditionLevel > 10 Then
          StateString = "AT: Transfer To Drain " & Pad(.AdditionLevel / 10, "0", 3) & "% "
          Timer = .Parameters_AddTransferToDrainTime
        Else:  StateString = "AT: Transfer To Drain " & TimerString(Timer.TimeRemaining)
        End If
        If Timer.Finished Then
          .AddReady = False
          .AddReadyFromPB = False
          .AP.ACCommand_Cancel
          State = ATOff
        End If
        
    End Select
  End With

End Sub
Friend Sub ACCommand_Cancel()
  Timer = 0
  State = ATOff
End Sub
Friend Property Get IsOn() As Boolean
  IsOn = (State <> ATOff)
End Property
Friend Property Get IsActive() As Boolean
  If IsTransfer Or (State = ATDosePause) Or IsRinse Or IsFillForMix Or IsTransferToDrain Then IsActive = True
End Property
Friend Property Get IsWaitReady() As Boolean
  If (State = ATWaitReady) Then IsWaitReady = True
End Property
Friend Property Get IsDelayedWaitReady() As Boolean
  IsDelayedWaitReady = OverrunTimer.Finished And IsWaitReady
End Property
Friend Property Get IsTransfer() As Boolean
  If (State = ATDose) Or (State = ATTransfer) Or (State = ATSettleTank) Or _
     (State = ATTransferEmpty1) Or (State = ATTransferEmpty2) Then IsTransfer = True
End Property
Friend Property Get IsDosing() As Boolean
  If (State = ATDose) Or (State = ATTransfer) Then IsDosing = True
End Property
Friend Property Get IsDosePaused() As Boolean
  If (State = ATDosePause) Then IsDosePaused = True
End Property
Friend Property Get IsRinse() As Boolean
  If (State = ATRinse) Then IsRinse = True
End Property
Friend Property Get IsFillForMix() As Boolean
  If (State = ATFillForMix) Then IsFillForMix = True
End Property
Friend Property Get IsMixForTime() As Boolean
  If (State = ATMixForTime) Then IsMixForTime = True
End Property
Friend Property Get IsCleaning() As Boolean
  If IsFillForMix Or IsMixForTime Or (State = ATCloseMixForTime) Then IsCleaning = True
End Property
Friend Property Get IsMixCloseMix() As Boolean
  If (State = ATCloseMixForTime) Then IsMixCloseMix = True
End Property
Friend Property Get IsAddPumping() As Boolean
  If IsTransfer Or IsMixForTime Or IsMixCloseMix Or IsDosePaused Then IsAddPumping = True
End Property
Friend Property Get IsTransferToMachine() As Boolean
  If (State = ATTransferEmpty1) Or (State = ATSettleTank) Or (State = ATTransferEmpty2) Then IsTransferToMachine = True
End Property
Friend Property Get IsTransferToDrain() As Boolean
  If (State = ATTransferToDrain) Then IsTransferToDrain = True
End Property
Friend Property Get IsDelayed() As Boolean
  If (State = ATWaitReady) Then IsDelayed = True
End Property
Friend Property Get IsPaused() As Boolean
  If (State = ATPause) Then IsPaused = True
End Property
'Function uses the expressions (y = x * x) and (y = sqr(x)) to generate curves for
'dosing control.
'x = A fraction between 0 and 1 (I know) representing elapsed time
'y = A fraction between 0 and 1 representing the amount we should have transferred so far
'
'The curves are scaled by adding a percent of the difference between the curve value and
'the equivalent linear value to the original curve value (??).
Private Function Setpoint() As Long
'If timer has finished exit function
  If Timer.Finished Then
    Setpoint = 0
    Exit Function
  End If
  
'Amount we should have transferred so far
  Dim ElapsedTime As Double, LinearTerm As Double, TransferAmount As Double
  ElapsedTime = (AddTime - Timer) / AddTime
  LinearTerm = ElapsedTime
  TransferAmount = StartLevel * LinearTerm
  
'Calculate scaling factor (0-1) for progressive and digressive curves
  If AddCurve > 0 Then
    Dim ScalingFactor As Double
    ScalingFactor = (10 - AddCurve) / 10
'Calculate term for progressive transfer (0-1) if odd curve
    If (AddCurve Mod 2) = 1 Then
      Dim MaxOddCurve As Double, OddTerm As Double
      MaxOddCurve = (ElapsedTime * ElapsedTime)
      OddTerm = MaxOddCurve + ((LinearTerm - MaxOddCurve) * ScalingFactor)
      TransferAmount = StartLevel * OddTerm
    Else
'Calculate term for digressive transfer (0-1) if even curve
      Dim MaxEvenCurve As Double, EvenTerm As Double
      MaxEvenCurve = Sqr(ElapsedTime)
      EvenTerm = MaxEvenCurve - ((MaxEvenCurve - LinearTerm) * ScalingFactor)
      TransferAmount = StartLevel * EvenTerm
    End If
  End If
   
'Calculate setpoint and limit to 0-1000
  Setpoint = StartLevel - CLng(TransferAmount)
  If Setpoint < 0 Then Setpoint = 0
  If Setpoint > 1000 Then Setpoint = 1000
  
'Global variable for display purposes - yuk!!
  DesiredLevel = Setpoint
End Function
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property

#End If
End Class
