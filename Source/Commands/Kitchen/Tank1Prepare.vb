'American & Efird - Mt Holly Then Platform
' Version 2024-09-26
' VB6 copy at bottom of Version 2.4 [20220214 1544] DH

Imports Utilities.Translations

Public Class Tank1Prepare : Inherits MarshalByRefObject
  Private ReadOnly controlCode As ControlCode

  Public Enum EState
    Off
    WaitIdle
    PreFill
    DispenseWaitTurn
    DispenseWaitReady
    DispenseWaitProducts
    DispenseWaitResponse
    Fill
    FillOff
    Heat
    Slow
    Fast
    FillToMixLevel
    MixForTime
    Ready

    ' Transferring
    TransferInterlock
    TransferPause
    Transfer1Empty
    Transfer1Delay
    Transfer1Rinse
    Transfer2Empty
    Transfer2Delay

    ' Background Rinsing to Drain
    TransferDrainRinse
    TransferDrainEmpty
    TransferDrainDelay

    InManual
  End Enum
  Public State As EState
  Private StateRestart As EState
  Public StatePrevious As EState
  Public Status As String

  Property Param_Destination As EKitchenDestination
  Property Param_Calloff As Integer
  Property Param_LevelFill As Integer
  Property Param_TempDesired As Integer
  Property Param_MixTime As Integer
  Property Param_MixRequest As Integer
  Property Param_OverrunTime As Integer
  Property RinseNumber As Integer

  Public HeatOn As Boolean
  Public HeatPrepTimer As New Timer
  Public FillOn As Boolean
  Public MixOn As Boolean

  Public ReadOnly Timer As New Timer
  Friend ReadOnly TimerAdvance As New Timer
  Friend ReadOnly TimerLevelDisregard As New Timer
  Public ReadOnly Property TimeLevelDisregard As Integer
    Get
      Return TimerLevelDisregard.Seconds
    End Get
  End Property

  Friend ReadOnly TimerMix As New Timer
  Friend ReadOnly TimerOverrun As New Timer

  Friend ReadOnly TimerUpWaitReady As New TimerUp

  'variables for alarms on manual dispense or error. does nothing
  Public ManualAdd As Boolean
  Public AddDispenseError As Boolean
  Public AlarmRedyeIssue As Boolean

  ' For display only
  Public AddRecipeStep As String
  Private ReadOnly RecipeSteps(64, 8) As String   ' TODO we update this   'vb6: Private RecipeSteps(1 To 64, 1 To 8) As String

  ' For Display in drugroom preview
  Public DrugroomDisplay As String

  ' Dispense States
  Public DispenseCalloff As Integer          'This filled in by us
  Public DispenseTank As Integer             'This filled in by us
  Public DispenseState As Integer            'This filled in by AutoDispenser
  Public DispenseProducts As String          'This filled in by Autodispenser

  Public DispenseDyesOnly As Boolean         'Used to determine dispenser delay type
  Public DispenseChemsOnly As Boolean
  Public DispenseDyesChems As Boolean
  Private DispenseDyes As Boolean
  Private DispenseChems As Boolean
  Friend DispenseTimer As New Timer

  'Dispense States
  Private Const DispenseReady As Integer = 101
  Private Const DispenseBusy As Integer = 102
  Private Const DispenseAuto As Integer = 201
  Private Const DispenseScheduled As Integer = 202
  Private Const DispenseComplete As Integer = 301
  Private Const DispenseManual As Integer = 302
  Private Const DispenseError As Integer = 309

  Public Sub New(ByVal controlCode As ControlCode)
    Me.controlCode = controlCode
  End Sub

  Public Sub Start(CallOff As Integer, Destination As EKitchenDestination, LevelDesired As Integer, TempDesired As Integer, MixTime As Integer, MixRequest As Integer, StandardPrepareTime As Integer)
    ' Set command parameter values 
    Param_Calloff = CallOff
    Param_Destination = Destination
    DispenseTank = CInt(Destination) ' 1=Add, 2=Reserve
    Param_LevelFill = LevelDesired
    Param_TempDesired = MinMax(TempDesired, 0, 1800)
    Param_MixTime = MixTime
    Param_MixRequest = MixRequest
    Param_OverrunTime = MinMax(StandardPrepareTime, 5, 120)

    ' Reset Recipe Details
    AddRecipeStep = ""
    DispenseDyes = False
    DispenseChems = False
    DispenseDyesOnly = False
    DispenseChemsOnly = False
    DispenseDyesChems = False

    AddDispenseError = False

    'Set state
    State = EState.WaitIdle
    TimerUpWaitReady.Pause()
    Timer.Seconds = 5

  End Sub

  Public Sub Run()
    With ControlCode
      DispenseState = .DispenseState
      DispenseProducts = .DispenseProducts

      ' Run an advance timer to reset alarms, where necessary
      If Not .AdvancePb Then TimerAdvance.TimeRemaining = 2
      If TimerAdvance.Finished Then AlarmRedyeIssue = False

      Dim pauseControl As Boolean = .Parent.IsPaused OrElse .EStop OrElse (Not .MachineClosed)

      Do
        ' Remember state and loop until state does not change
        ' This makes sure we go through all the state changes before proceeding and setting IO
        StatePrevious = State

        ' Pause tank control
        If pauseControl Then
          If (Not State = EState.TransferPause) AndAlso (State > EState.TransferPause) AndAlso (State < EState.InManual) Then
            StateRestart = State
            State = EState.TransferPause
            Timer.Pause()
          End If
        End If

        '********************************************************************************************
        '******   STATE LOGIC
        '********************************************************************************************
        Select Case State

          Case EState.Off
            Status = (" ")
            DrugroomDisplay = Translate("Tank 1") & (" ") & Translate("Idle")


          Case EState.WaitIdle
            Status = Translate("Wait Idle") & Timer.ToString(1)
            DrugroomDisplay = Translate("Tank 1") & (" ") & Translate("Wait Idle") & Timer.ToString(1)
            If AlarmRedyeIssue AndAlso (.Parameters.EnableRedyeIssueAlarm = 1) Then
              Timer.Seconds = 5
              Status = "Check Tank (Redye Issue), Hold Run to reset" & TimerAdvance.ToString(1)
              DrugroomDisplay = Translate("Tank 1") & (" ") & "Check Tank (Redye Issue), Hold Run to Reset" & TimerAdvance.ToString(1)
            ElseIf (Param_Destination = EKitchenDestination.Add) AndAlso (.AD.IsOn OrElse .AddControl.IsActive) Then
              Timer.Seconds = 5
              If .AD.IsOn Then Status = "Wait to Add - " & .AD.StateString
              If .AddControl.IsActive Then Status = "Wait to Add - " ' & .AddControl.Status
            ElseIf (Param_Destination = EKitchenDestination.Reserve) AndAlso (.RD.IsOn OrElse .ReserveControl.IsActive) Then
              Timer.Seconds = 5
              Status = "Wait to Reserve - " & .RD.Status
            End If
            If .CK.IsOn Then
              Status = .CK.Status
              Timer.Seconds = 2
            End If
            If Timer.Finished Then
              RinseNumber = .Parameters.Tank1RinseNumber
              State = EState.PreFill
              ' Delay timer just to account for bogus readings from tanklevel transmitter
              Timer.Seconds = 1
              TimerLevelDisregard.Seconds = MinMax(.Parameters.Tank1LevelDisregardTimeFill, 0, 15)
            End If



          Case EState.PreFill
            If .Parameters.Tank1PreFillLevel > 0 Then
              ' Use Timer for bogus tank readings
              If (.Tank1Level < .Parameters.Tank1PreFillLevel) Then
                Timer.Seconds = 1
              End If
              If (.Tank1Level >= .Parameters.Tank1PreFillLevel) AndAlso Timer.Finished Then
                State = EState.DispenseWaitTurn
              End If
              Status = Translate("PreFilling") & (" ") & (.Tank1Level / 10).ToString("#0.0") & " / " & (.Parameters.Tank1PreFillLevel / 10).ToString("#0.0") & "%" & Timer.ToString(1)
            Else
              DispenseTimer.Seconds = .Parameters.DispenseReadyDelayTime * 60
              State = EState.DispenseWaitTurn
              Status = Translate("PreFilling") & (" ") & (.Tank1Level / 10).ToString("#0.0") & " / " & (.Parameters.Tank1PreFillLevel / 10).ToString("#0.0") & "%"
            End If
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status



          Case EState.DispenseWaitTurn
            'Wait for other tank to be finished with dispensing before new dispense
            If (Param_Calloff = 0) OrElse (.Parameters.DispenseEnabled <> 1) Then
              HeatPrepTimer.Seconds = .Parameters.Tank1HeatPrepTime
              State = EState.Fill
              ManualAdd = True
            Else
              DispenseTimer.Minutes = .Parameters.DispenseReadyDelayTime
              State = EState.DispenseWaitReady
              DispenseCalloff = 0
            End If
            Status = "Wait for Turn "
            DrugroomDisplay = Status


          Case EState.DispenseWaitReady
            'Wait to make sure dispenser has completed previous dispense and is ready
            If DispenseState = EDispenseResult.Ready Then     ' 101
              If .Parameters.DispenseTestState > 0 Then .Parameters.DispenseTestState = 0
              State = EState.DispenseWaitProducts
              DispenseTimer.Minutes = .Parameters.DispenseResponseDelayTime
              DispenseCalloff = Param_Calloff
              .DispenseTank = CInt(Param_Destination) ' paramDispenseTank
            End If
            If (Param_Calloff = 0) OrElse (.Parameters.DispenseEnabled <> 1) Then
              HeatPrepTimer.Seconds = .Parameters.Tank1HeatPrepTime
              State = EState.Fill
              DispenseCalloff = 0
              ManualAdd = True
            End If
            Status = "Wait for Dispenser Ready "
            DrugroomDisplay = Status


          Case EState.DispenseWaitProducts
            ' Wait for Dispensestate & DispenseProducts String to be set to determine how to signal delays
            Status = "Wait for Dispense Products "
            DrugroomDisplay = Status
            DispenseTimer.Minutes = .Parameters.DispenseResponseDelayTime
            If DispenseState <> DispenseReady Then                '(DispenseBusy = 102) but if these no recipe, we'll get a (DispenseManual = 302)
              If DispenseProducts <> "" Then
                'split the products
                'products() = "Step <calloff>| <Ingredient_id> : <Amount> <Units> <DResult> | <Ingredient_Desc> "
                Dim ProductsArray() As String
                ' ProductsArray = Split(DispenseProducts, "|")
                ProductsArray = DispenseProducts.Split(CType("|", Char()))
                If ProductsArray.Length > 1 Then
                  For i As Integer = 1 To ProductsArray.Length - 1            'Disregard the 0 row "Step <calloff>

                    If ProductsArray(i) = ":" Then
                      'Dim test As String = ProductsArray(i, 5)

                    End If


                  Next i
                End If

                'For i = 1 To UBound(ProductsArray)                'Disregard the 0 row "Step <calloff>
                '  Dim position As Long
                '  position = InStr(ProductsArray(i), ":")         'Need to verify ":" is in the current row due to second split "|"
                '  If position > 0 Then
                '    Dim test As String
                '    test = Mid(ProductsArray(i), 6, 1)
                '    If test = ":" Then                            'If ":" is at 5th position, we have a chemical => "| 1004: ..."
                '      DispenseChems = True
                '    Else
                '      DispenseDyes = True
                '    End If
                '  End If
                'Next i
                ' Set Flags
                If DispenseChems And DispenseDyes Then
                  DispenseDyesChems = True
                Else
                  If DispenseChems Then
                    DispenseChemsOnly = True
                  Else : DispenseDyesOnly = True
                  End If
                End If
              End If
              'Proceed to the next state
              State = EState.DispenseWaitResponse
              DispenseCalloff = Param_Calloff ' AddCallOff
              .DispenseTank = DispenseTank
            End If
            If (Param_Calloff = 0) Or (.Parameters.DispenseEnabled <> 1) Then
              HeatPrepTimer.Seconds = .Parameters.Tank1HeatPrepTime
              State = EState.Fill
              DispenseCalloff = 0
              ManualAdd = True
            End If


          Case EState.DispenseWaitResponse
            ' Wait for response from Dispenser
            AddRecipeStep = DispenseProducts
            Select Case DispenseState                       ' 301, 302, 309
              Case EDispenseResult.Complete
                .Parameters.DispenseTestState = 1
                If Not .WK.IsOn Then
                  If DispenseTank = 2 Then
                    .ReserveReady = True
                  ElseIf DispenseTank = 1 Then
                    .AddReady = True
                  End If
                  .KR.Cancel()
                  If .LA.KP1.IsOn Then .LAActive = True
                  State = EState.Off
                  .DispenseTank = 0
                  DispenseTank = 0
                  DispenseCalloff = 0
                End If

              Case EDispenseResult.Manual
                .Parameters.DispenseTestState = 1
                HeatPrepTimer.Seconds = .Parameters.Tank1HeatPrepTime
                State = EState.Fill
                DispenseCalloff = 0
                ManualAdd = True

              Case EDispenseResult.Error
                .Parameters.DispenseTestState = 1
                AddDispenseError = True
                HeatPrepTimer.Seconds = .Parameters.Tank1HeatPrepTime
                State = EState.Fill
                DispenseCalloff = 0

            End Select
            Status = "Wait For Reponse From Dispenser "
            DrugroomDisplay = "Wait for Response From Dispenser "


          Case EState.Fill
            ' Monitor tank level switch when complete
            If (.Tank1Level > Param_LevelFill) AndAlso Timer.Finished Then
              State = EState.FillOff
              Timer.Seconds = 2
            End If
            ' Disregard level if parameter set, due to flaky UV level transmitters
            If TimerLevelDisregard.Finished AndAlso (.Parameters.Tank1LevelDisregardTimeFill > 0) Then
              State = EState.FillOff
              Timer.Seconds = 2
            End If
            ' No level requested
            If (Param_LevelFill = 0) Then
              TimerUpWaitReady.Start()
              State = EState.Slow
            End If
            ' Switch in manual
            If .IO.Tank1ManualSw Then State = EState.InManual
            Status = Translate("Filling") & (" ") & (.Tank1Level / 10).ToString("#0.0") & " / " & (Param_LevelFill / 10).ToString("#0.0") & "%"
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status


          Case EState.FillOff
            Status = Translate("Fill Stop") & Timer.ToString(1)
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status
            If Timer.Finished Then
              If Settings.Tank1HeatEnabled = 1 Then
                State = EState.Heat
              Else
                ' No Heat available in drugroom, step over heat
                HeatOn = False
                If .Tank1MixerEnable Then MixOn = True
                State = EState.Slow
                TimerUpWaitReady.Start()
                TimerOverrun.Minutes = Param_OverrunTime
              End If
            End If
            ' Switch in Manual
            If .IO.Tank1ManualSw Then State = EState.InManual


          Case EState.Heat
            If Not .Tank1MixerEnable Then
              Status = Translate("Level Low") & (" ") & (.Tank1Level / 10).ToString("#0.0") & " < " & (.Parameters.Tank1MixerOnLevel / 10).ToString("#0.0") & "%"
              MixOn = False
            Else
              Status = Translate("Heating") & (" ") & (.IO.Tank1Temp / 10).ToString("#0.0") & " / " & (Param_TempDesired / 10).ToString("#0.0") & "F"
              MixOn = True
            End If
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status
            ' Need to continue heating
            If (.IO.Tank1Temp < Param_TempDesired) AndAlso MixOn Then HeatOn = True
            If (.IO.Tank1Temp > Param_TempDesired) OrElse (Param_LevelFill = 0) OrElse (Param_TempDesired = 0) Then
              HeatOn = False
              State = EState.Slow
              TimerUpWaitReady.Start()
              TimerOverrun.Minutes = Param_OverrunTime
            End If
            ' Switch in Manual
            If .IO.Tank1ManualSw Then State = EState.InManual



          Case EState.Slow
            Status = Translate("Tank 1") & (" ") & Translate("Wait Ready") & (" ") & TimerUpWaitReady.ToString
            DrugroomDisplay = Status
            ' Is tank1heating enabled/available
            If Settings.Tank1HeatEnabled = 1 Then
              If .IO.Tank1Temp < Param_TempDesired Then
                HeatOn = True
                MixOn = True
              End If
            Else
              HeatOn = False
            End If
            If .IO.Tank1Temp > Param_TempDesired Then HeatOn = False
            ' Tank 1 level monitoring
            If Not .Tank1MixerEnable Then
              HeatOn = False
              MixOn = False
            Else : MixOn = True
            End If
            ' Destination - continue
            If Param_Destination = EKitchenDestination.Reserve Then
              If .RT.IsWaitReady Then State = EState.Fast
            ElseIf Param_Destination = EKitchenDestination.Add Then
              If .AC.IsWaitReady Then State = EState.Fast
              If .AT.IsWaitReady Then State = EState.Fast
            End If
            ' Waiting too long
            If TimerOverrun.Finished Then State = EState.Fast
            ' Ready Pressed
            If .Tank1Ready Then
              TimerUpWaitReady.Pause()
              ' Mixing requested, already at level, no need to fill [Issue with overfilling due to UV level transmitters]
              If (Param_MixRequest > 0) AndAlso (Param_MixTime > 0) Then
                If (Not .Tank1MixerEnable) Then
                  ' Need to Mix, Fill to level
                  MixOn = False
                  HeatOn = False
                  State = EState.FillToMixLevel
                  ' Delay timer just to account for bogus readings from tanklevel transmitter
                  Timer.Seconds = 2
                  TimerLevelDisregard.Seconds = MinMax(.Parameters.Tank1LevelDisregardTimeFill, 0, 15)
                Else
                  MixOn = True
                  State = EState.MixForTime
                  TimerMix.Seconds = Param_MixTime
                End If
              Else
                ' Step to Ready State
                State = EState.Ready
              End If
            End If
            ' Restart Waiting Timer - should copy over LA to KP
            If TimerUpWaitReady.IsPaused Then TimerUpWaitReady.Restart()
            ' Switch in Manual
            If .IO.Tank1ManualSw Then State = EState.InManual



          Case EState.Fast
            Status = Translate("Tank 1") & (" ") & Translate("Wait Ready") & (" ") & TimerUpWaitReady.ToString
            DrugroomDisplay = Status
            ' Is tank1heating enabled/available
            If Settings.Tank1HeatEnabled = 1 Then
              If .IO.Tank1Temp < Param_TempDesired Then HeatOn = True
              If .IO.Tank1Temp > Param_TempDesired Then HeatOn = False
            Else
              HeatOn = False
            End If
            ' Tank 1 level monitoring
            If Not .Tank1MixerEnable Then
              HeatOn = False
              MixOn = False
            Else : MixOn = True
            End If
            ' Ready Pressed
            If .Tank1Ready Then
              TimerUpWaitReady.Pause()
              ' Mixing requested, already at level, no need to fill [Issue with overfilling due to UV level transmitters]
              If (Param_MixRequest > 0) AndAlso (Param_MixTime > 0) Then
                If (Not .Tank1MixerEnable) Then
                  ' Need to Mix, Fill to level
                  MixOn = False
                  HeatOn = False
                  State = EState.FillToMixLevel
                  ' Delay timer just to account for bogus readings from tanklevel transmitter
                  Timer.Seconds = 2
                  TimerLevelDisregard.Seconds = MinMax(.Parameters.Tank1LevelDisregardTimeFill, 0, 15)
                Else
                  MixOn = True
                  State = EState.MixForTime
                  TimerMix.Seconds = Param_MixTime
                End If
              Else
                ' Step to Ready State
                State = EState.Ready
              End If
            End If
            ' Restart Waiting Timer - should copy over LA to KP
            If TimerUpWaitReady.IsPaused Then TimerUpWaitReady.Restart()
            ' Switch in Manual
            If .IO.Tank1ManualSw Then State = EState.InManual


          Case EState.FillToMixLevel
            Status = Translate("Filling")
            If .Tank1Level < (.Parameters.Tank1MixerOnLevel) Then
              Status = Status & (" ") & (.Tank1Level / 10).ToString("#0.0") & " / " & (.Parameters.Tank1MixerOnLevel / 10).ToString("#0.0") & "%"
              Timer.Seconds = 2
            Else
              Status = Status & Timer.ToString(1)
            End If
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status
            ' Monitor tank level switch when complete
            If (.Tank1Level > .Parameters.Tank1MixerOnLevel) AndAlso Timer.Finished Then
              MixOn = True
              State = EState.MixForTime
              TimerMix.Seconds = Param_MixTime
            End If
            ' Disregard level if parameter set, due to flaky UV level transmitters
            If TimerLevelDisregard.Finished AndAlso (.Parameters.Tank1LevelDisregardTimeFill > 0) Then
              MixOn = True
              State = EState.MixForTime
              TimerMix.Seconds = Param_MixTime
            End If
            ' Mix Not Requested
            If (Param_MixRequest = 0) OrElse (Param_MixTime = 0) Then
              State = EState.Ready
            End If
            ' Tank Ready lost, return to prepare
            If Not .Tank1Ready Then
              State = EState.Fast
              TimerUpWaitReady.Restart()
            End If
            ' Switch in manual
            If .IO.Tank1ManualSw Then State = EState.InManual


          Case EState.MixForTime
            If (Param_MixRequest > 0) AndAlso (Param_MixTime > 0) Then
              If Not .Tank1MixerRequest Then .Tank1MixerRequest = True
              If TimerMix.Finished Then State = EState.Ready
            Else
              .Tank1MixerRequest = False
              State = EState.Ready
            End If
            If .IO.Tank1Temp < Param_TempDesired Then HeatOn = True
            If .IO.Tank1Temp > Param_TempDesired Then HeatOn = False
            ' Tank Ready lost, return to prepare
            If Not .Tank1Ready Then
              State = EState.Fast
              TimerUpWaitReady.Restart()
            End If
            ' Switch in manual
            If .IO.Tank1ManualSw Then State = EState.InManual
            Status = Translate("Mixing") & (" ") & TimerMix.ToString
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status


          Case EState.Ready
            Status = Translate("Ready")
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status
            If .IO.Tank1ManualSw Then State = EState.InManual
            If Param_Destination = EKitchenDestination.Add And (.AddLevel > .Parameters.AdditionMaxTransferLevel) Then
              State = EState.TransferInterlock
            Else
              .Tank1MixerRequest = False
              .Tank1Ready = False
              FillOn = False
              HeatOn = False
              MixOn = False
              AddDispenseError = False
              ManualAdd = False
              State = EState.Transfer1Empty
              Timer.Seconds = 2
              TimerLevelDisregard.Seconds = MinMax(.Parameters.Tank1LevelDisregardTimeTransfer, 10, 100)
            End If


            '********************************************************************************************
            '******   TRANSFER & RINSE TO MACHINE
            '********************************************************************************************

          Case EState.TransferInterlock
            Status = Translate("Add Level Too High") & (" ") & (.AddLevel / 10).ToString("#0.0") & " > " & (.Parameters.AdditionMaxTransferLevel / 10).ToString("#0.0") & "%"
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status
            If .IO.Tank1ManualSw Then State = EState.InManual
            If (.AddLevel <= .Parameters.AdditionMaxTransferLevel) Then
              .Tank1Ready = False
              FillOn = False
              HeatOn = False
              MixOn = False
              State = EState.Transfer1Empty
              Timer.Seconds = 2
              TimerLevelDisregard.Seconds = MinMax(.Parameters.Tank1LevelDisregardTimeTransfer, 10, 100)
            End If

          Case EState.TransferPause
            Status = Translate("Transfer Paused")
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status
            If .EStop Then Status = Translate("Transfer Paused") & (" ") & Translate("EStop")
            If Not pauseControl Then
              State = StateRestart
              Timer.Restart()
            End If

          Case EState.Transfer1Empty
            Status = Translate("Transferring")
            If (.Tank1Level > 50) Then
              Status = Status & (" ") & (.Tank1Level / 10).ToString("#0.0") & "%"
              ' Reset Level Timer to 2 seconds - when using level, jump 
              Timer.Seconds = .Parameters.Tank1TimeBeforeRinse
            Else
              Status = Status & Timer.ToString(1)
            End If
            If .AdvancePb Then Status = "Transfer Advancing " & TimerAdvance.ToString
            If (.Parameters.Tank1LevelDisregardTimeTransfer > 0) Then
              ' Disregard tank level input due to signal bouncing - use delay timer
              Status = Translate("Transfer") & (" ") & TimerLevelDisregard.ToString
              If TimerLevelDisregard.Finished Then
                State = EState.Transfer1Delay
                Timer.Seconds = .Parameters.Tank1TimeBeforeRinse
              End If
            Else
              ' Use level timer
              If Timer.Finished Then
                State = EState.Transfer1Delay
                Timer.Seconds = .Parameters.Tank1TimeBeforeRinse
              End If
            End If
            If TimerAdvance.Finished Then
              State = EState.Transfer1Delay
              Timer.Seconds = .Parameters.Tank1TimeBeforeRinse
            End If
            ' Switch to manual mode
            If .IO.Tank1ManualSw Then State = EState.InManual
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status


          Case EState.Transfer1Delay
            ' Once level drops below 1.0% for 2 seconds, disregard level entireley 
            '   UltraSonic kitchen tank level transmitters bounce around when tank is empty, this is a workaround
            Status = Translate("Transfer") & Timer.ToString(1)
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status
            If .IO.Tank1ManualSw Then State = EState.InManual
            If Timer.Finished Then
              ' Kitchen Rinse command active and set to '0'
              If .KR.IsOn Then
                If (.KR.RinseMachine = 0) And (.KR.RinseDrain = 0) Then
                  If Not .WK.IsOn Then
                    If Param_Destination = EKitchenDestination.Add Then .AddReady = True
                    If Param_Destination = EKitchenDestination.Reserve Then .ReserveReady = True
                  End If
                  If .LA.KP1.IsOn Then .LAActive = True
                  .Tank1Ready = False
                  .KR.Cancel()
                  Cancel()
                ElseIf .KR.RinseMachine = 0 Then
                  State = EState.TransferDrainRinse
                  Timer.Seconds = MinMax(.Parameters.Tank1RinseToDrainTime, 5, 300)
                End If
              Else
                State = EState.Transfer1Rinse
                Timer.Seconds = .Parameters.Tank1RinseTime
              End If
            End If
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status


          Case EState.Transfer1Rinse
            Status = Translate("Rinsing") & Timer.ToString(1)
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status
            If Timer.Finished Then
              State = EState.Transfer2Empty
              Timer.Seconds = 2
              TimerLevelDisregard.Seconds = MinMax(.Parameters.Tank1LevelDisregardTimeTransfer, 10, 100)
            End If
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status
            ' Switch to manual mode
            If .IO.Tank1ManualSw Then State = EState.InManual


          Case EState.Transfer2Empty
            Status = Translate("Transferring")
            If (.Tank1Level > 50) Then
              Status = Status & (" ") & (.Tank1Level / 10).ToString("#0.0") & "%"
              ' Reset Level Timer to 2 seconds - when using level, jump 
              Timer.Seconds = 2
            Else
              Status = Status & Timer.ToString(1)
            End If
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status
            If .AdvancePb Then Status = "Transfer Advancing " & TimerAdvance.ToString
            If (.Parameters.Tank1LevelDisregardTimeTransfer > 0) Then
              ' Disregard tank level input due to signal bouncing - use delay timer
              Status = Translate("Transfer") & (" ") & TimerLevelDisregard.ToString
              If TimerLevelDisregard.Finished Then
                State = EState.Transfer2Delay
                Timer.Seconds = .Parameters.Tank1TimeAfterRinse
              End If
            Else
              ' Use level timer
              If Timer.Finished Then
                State = EState.Transfer2Delay
                Timer.Seconds = .Parameters.Tank1TimeAfterRinse
              End If
            End If
            If TimerAdvance.Finished Then
              State = EState.Transfer2Delay
              Timer.Seconds = .Parameters.Tank1TimeAfterRinse
            End If
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status
            ' Switch to manual mode
            If .IO.Tank1ManualSw Then State = EState.InManual


          Case EState.Transfer2Delay
            ' Once level drops below 1.0% for 2 seconds, disregard level entireley 
            '   UltraSonic kitchen tank level transmitters bounce around when tank is empty, this is a workaround
            Status = Translate("Transfer") & Timer.ToString(1)
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status
            If Timer.Finished Then
              If Not .WK.IsOn Then
                If Param_Destination = EKitchenDestination.Add Then .AddReady = True
                If Param_Destination = EKitchenDestination.Reserve Then .ReserveReady = True
              End If
              If .LA.KP1.IsOn Then .LAActive = True
              .Tank1Ready = False
              State = EState.TransferDrainRinse
              Timer.Seconds = MinMax(.Parameters.Tank1RinseToDrainTime, 5, 300)
            End If
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status
            If .IO.Tank1ManualSw Then State = EState.InManual


          Case EState.TransferDrainRinse
            Status = Translate("Rinsing to Drain") & Timer.ToString(1)
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status
            If (.Parameters.Tank1RinseToDrainLevel > 0) Then
              ' Use Rinse Level
              Status = Translate("Fill to Rinse") & (" ") & (.Tank1Level / 10).ToString("#0.0") & (" / ") & (.Parameters.Tank1RinseToDrainLevel / 10).ToString("#0.0") & "%"
              If (.Tank1Level > .Parameters.Tank1RinseToDrainLevel) Then
                RinseNumber = RinseNumber - 1
                State = EState.TransferDrainEmpty
                Timer.Seconds = 2
                TimerLevelDisregard.Seconds = MinMax(.Parameters.Tank1LevelDisregardTimeTransfer, 10, 100)
              End If
            Else
              ' Use Rinse Time
              If Timer.Finished Then
                RinseNumber = RinseNumber - 1
                State = EState.TransferDrainEmpty
                Timer.Seconds = .Parameters.Tank1DrainTime
                TimerLevelDisregard.Seconds = MinMax(.Parameters.Tank1LevelDisregardTimeTransfer, 10, 100)
              End If
            End If
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status


          Case EState.TransferDrainEmpty
            Status = Translate("Transfer to Drain")
            If (.Tank1Level > 50) Then
              Status = Status & (" ") & (.Tank1Level / 10).ToString("#0.0") & "%"
              ' Reset Level Timer to 2 seconds - when using level, jump 
              Timer.Seconds = .Parameters.Tank1DrainTime
            Else
              Status = Status & Timer.ToString(1)
            End If
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status
            If .AdvancePb Then Status = "Transfer Advancing " & TimerAdvance.ToString
            If (.Parameters.Tank1LevelDisregardTimeTransfer > 0) Then
              ' Disregard tank level input due to signal bouncing - use delay timer
              Status = Translate("Transfer") & (" ") & TimerLevelDisregard.ToString
              If TimerLevelDisregard.Finished Then
                State = EState.TransferDrainDelay
                Timer.Seconds = .Parameters.Tank1DrainTime
              End If
            Else
              ' Use level timer
              If Timer.Finished Then
                State = EState.TransferDrainDelay
                Timer.Seconds = .Parameters.Tank1DrainTime
              End If
            End If
            If TimerAdvance.Finished Then
              State = EState.TransferDrainDelay
              Timer.Seconds = .Parameters.Tank1DrainTime
            End If
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status


          Case EState.TransferDrainDelay
            ' Once level drops below 1.0% for 2 seconds, disregard level entireley 
            '   UltraSonic kitchen tank level transmitters bounce around when tank is empty, this is a workaround
            Status = Translate("Draining") & Timer.ToString(1)
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status
            If .IO.Tank1ManualSw Then State = EState.InManual
            If Timer.Finished Then
              If RinseNumber > 0 Then
                State = EState.TransferDrainRinse
                Timer.Seconds = MinMax(.Parameters.Tank1RinseToDrainTime, 5, 300)
              Else
                If Not .WK.IsOn Then
                  If (DispenseTank = 1) Then .AddReady = True
                  If (DispenseTank = 2) Then .ReserveReady = True
                End If
                If .LA.KP1.IsOn Then .LAActive = True
                DispenseTank = 0
                .DispenseTank = 0
                .Tank1Ready = False
                .KR.Cancel()
                Cancel()
              End If
            End If
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status


          Case EState.InManual
            Status = Translate("In Manual") & (" ") & (.Tank1Level / 10).ToString("#0.0") & "%"
            DrugroomDisplay = Translate("Tank 1") & (" ") & Status
            .Tank1Ready = False
            HeatPrepTimer.Seconds = .Parameters.Tank1HeatPrepTime
            If Not .IO.Tank1ManualSw Then State = EState.Fill


        End Select

      Loop Until (StatePrevious = State) 'Loop until state does not change
    End With
  End Sub

  Public Sub Cancel()
    Param_Destination = 0
    Param_Calloff = 0
    Param_MixRequest = 0
    Param_MixTime = 0
    Param_LevelFill = 0
    Param_TempDesired = 0

    FillOn = False
    HeatOn = False
    MixOn = False

    AddRecipeStep = ""

    DispenseDyes = False
    DispenseChems = False
    DispenseDyesOnly = False
    DispenseChemsOnly = False
    DispenseDyesChems = False

    State = EState.Off

    TimerUpWaitReady.Pause()

  End Sub

  Public Sub CopyTo(ByRef Target As Tank1Prepare)
    With Target
      .State = State

      .Param_Calloff = Param_Calloff
      .Param_Destination = Param_Destination
      .Param_LevelFill = Param_LevelFill
      .Param_TempDesired = Param_TempDesired
      .Param_MixTime = Param_MixTime
      .Param_MixRequest = Param_MixRequest
      .Param_OverrunTime = Param_OverrunTime

      .AddRecipeStep = AddRecipeStep

      .Timer.Seconds = Timer.Seconds
      .TimerOverrun.Seconds = TimerOverrun.Seconds

      ' TODO - add a method to copy up timer value 
      '  .TimerUpWaitReady.TimeElapsed
      '  .TimerUpWaitReady.StartTimeUtc = TimerUpWaitReady.StartTimeUtc

    End With
  End Sub

  Public Sub ProgramStart()
    On Error Resume Next

    Param_Calloff = 0
    AddRecipeStep = ""
    Dim i1, i2 As Integer
    For i1 = 1 To 64
      For i2 = 1 To 8
        RecipeSteps(i1, i2) = ""
      Next i2
    Next i1
  End Sub

#Region " PROPERTIES "

  ReadOnly Property IsOn As Boolean
    Get
      Return (State <> EState.Off)
    End Get
  End Property

  Public ReadOnly Property IsReady() As Boolean
    Get
      Return (State = EState.Ready)
    End Get
  End Property
  Public ReadOnly Property IsManualPrepareOverrun() As Boolean
    Get
      Return (IsFast AndAlso TimerOverrun.Finished)
    End Get
  End Property
  Public ReadOnly Property IsSlow() As Boolean
    Get
      Return (State = EState.Slow)
    End Get
  End Property
  Public ReadOnly Property IsFast() As Boolean
    Get
      Return (State = EState.Fast)
    End Get
  End Property
  Public ReadOnly Property IsWaitReady As Boolean
    Get
      Return IsSlow OrElse IsFast OrElse (State = EState.MixForTime)
    End Get
  End Property
  Public ReadOnly Property IsWaitingToDispense As Boolean
    Get
      Return False ' TODO
    End Get
  End Property
  Public ReadOnly Property IsDispensing As Boolean
    Get
      Return False ' TODO
    End Get
  End Property


  Public ReadOnly Property IsForeground As Boolean
    Get
      Return ((State >= EState.WaitIdle) AndAlso (State <= EState.Transfer2Delay))

    End Get
  End Property
  Public ReadOnly Property IsBackground As Boolean
    Get
      Return (State >= EState.TransferDrainRinse) AndAlso (State <= EState.TransferDrainDelay)
    End Get
  End Property
  Public ReadOnly Property IsInManual As Boolean
    Get
      Return (State = EState.InManual)
    End Get
  End Property
  Public ReadOnly Property IsOverrun As Boolean
    Get
      Return TimerOverrun.Finished AndAlso IsWaitReady
    End Get
  End Property

#End Region

#Region " I/O PROPERTIES "

  ReadOnly Property IoTank1Lamp As Boolean
    Get
      Return State = EState.Slow OrElse State = EState.Fast
      ' TODO - Include HeatTank Overrun
    End Get
  End Property

  Public ReadOnly Property IoFillCold As Boolean
    Get
      Return (State = EState.PreFill) OrElse (State = EState.Fill) OrElse (State = EState.FillToMixLevel) OrElse (State = EState.Transfer1Rinse) OrElse (State = EState.TransferDrainRinse)
      'Return (State = EState.Fill) OrElse (State = EState.FillToMixLevel) OrElse (State = EState.Transfer1Rinse) OrElse (State = EState.TransferDrainRinse)
    End Get
  End Property

  Public ReadOnly Property IoFillHot As Boolean
    Get
      Return False
    End Get
  End Property

  Public ReadOnly Property IoMixerOn As Boolean
    Get
      Return MixOn AndAlso ((State = EState.Heat) OrElse (State = EState.Slow) OrElse (State = EState.Fast) OrElse (State = EState.MixForTime) OrElse (State = EState.Ready) OrElse (State = EState.TransferInterlock))
    End Get
  End Property

  Public ReadOnly Property IoTransfer As Boolean
    Get
      Return (State = EState.Transfer1Empty) OrElse (State = EState.Transfer1Delay) OrElse (State = EState.Transfer1Rinse) OrElse
             (State = EState.Transfer2Empty) OrElse (State = EState.Transfer2Delay) OrElse
             (State = EState.TransferDrainEmpty) OrElse (State = EState.TransferDrainDelay)

      'Return (State = EState.Transfer1Empty) OrElse (State = EState.Transfer1Delay) OrElse (State = EState.Transfer1Rinse) OrElse
      '       (State = EState.Transfer2Empty) OrElse (State = EState.Transfer2Delay) OrElse
      '       (State = EState.TransferDrainEmpty) OrElse (State = EState.TransferDrainDelay)
    End Get
  End Property

  Public ReadOnly Property IoToAddTank As Boolean
    Get
      Return Param_Destination = EKitchenDestination.Add AndAlso ((State = EState.Transfer1Empty) OrElse (State = EState.Transfer1Delay) OrElse (State = EState.Transfer1Rinse) OrElse (State = EState.Transfer2Empty) OrElse (State = EState.Transfer2Delay))
      '    Return Param_Destination = EKitchenDestination.Add AndAlso ((State = EState.Transfer1Empty) OrElse (State = EState.Transfer1Delay) OrElse (State = EState.Transfer1Rinse) OrElse (State = EState.Transfer2Empty) OrElse (State = EState.Transfer2Delay))
    End Get
  End Property

  Public ReadOnly Property IoToReserveTank As Boolean
    Get
      Return Param_Destination = EKitchenDestination.Reserve AndAlso ((State = EState.Transfer1Empty) OrElse (State = EState.Transfer1Delay) OrElse (State = EState.Transfer1Rinse) OrElse (State = EState.Transfer2Empty) OrElse (State = EState.Transfer2Delay))
      '  Return Param_Destination = EKitchenDestination.Reserve AndAlso ((State = EState.Transfer1Empty) OrElse (State = EState.Transfer1Delay) OrElse (State = EState.Transfer1Rinse) OrElse (State = EState.Transfer2Empty) OrElse (State = EState.Transfer2Delay))
    End Get
  End Property

  Public ReadOnly Property IoDrain As Boolean
    Get
      Return (State = EState.TransferDrainRinse) OrElse (State = EState.TransferDrainEmpty) OrElse (State = EState.TransferDrainDelay)
      ' Return (State = EState.TransferDrainRinse) OrElse (State = EState.TransferDrainEmpty) OrElse (State = EState.TransferDrainDelay)
    End Get
  End Property

  Public ReadOnly Property IoSteam As Boolean
    Get
      Return HeatOn AndAlso ((State = EState.Heat) OrElse (State = EState.Slow) OrElse (State = EState.Fast) OrElse (State = EState.MixForTime) OrElse (State = EState.Ready) OrElse (State = EState.TransferInterlock))
    End Get
  End Property

#End Region


  ' VB6 from latest AEThenPlatform project
#If 0 Then
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acAddPrepare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===============================================================================================
'General purpose add tank prepare command
'===============================================================================================
Option Explicit
Public Enum AddPrepareState
  Off
  WaitIdle
  PreFill
  DispenseWaitTurn
  DispenseWaitReady
  DispenseWaitProducts
  DispenseWaitResponse
  Fill
  Heat
  Slow
  Fast
  MixForTime
  Ready
  KAInterlock
  KATransfer1
  KARinse
  KATransfer2
  KARinseToDrain
  KATransferToDrain
  InManual
End Enum

Public State As AddPrepareState
Public StatePrev As AddPrepareState
Public StateString As String

Public AddPreFillLevel As Long
Public AddFillLevel As Long
Public DesiredTemperature As Long
Public AddMixTime As Long
Public AddMixing As Boolean
Public OverrunTime As Long
Public OverrunTimer As New acTimer
Public MixTimer As New acTimer
Public HeatOn As Boolean, FillOn As Boolean
Public AddCallOff As Long
Public Timer As New acTimer
Public AdvanceTimer As New acTimer
Public HeatPrepTimer As New acTimer
Public DispenseTimer As New acTimer
Private NumberOfRinses As Long
Public WaitReadyTimer As New acTimerUp
 
'variables for alarms on manual dispense or error. does nothing
Public ManualAdd As Boolean
Public AddDispenseError As Boolean
Public AlarmRedyeIssue As Boolean

'For display only
Public AddRecipeStep As String
Private RecipeSteps(1 To 64, 1 To 8) As String

'For Display in drugroom preview
Public DrugroomDisplay As String

'Dispense States
Public DispenseCalloff As Long            'This filled in by us
Public DispenseTank As Long               'This filled in by us
Public Dispensestate As Long              'This filled in by AutoDispenser
Public Dispenseproducts As String         'This filled in by Autodispenser

Public DispenseDyesOnly As Boolean        'Used to determine dispenser delay type
Public DispenseChemsOnly As Boolean
Public DispenseDyesChems As Boolean
Private DispenseDyes As Boolean
Private DispenseChems As Boolean

'Dispense States
Private Const DispenseReady As Long = 101
Private Const DispenseBusy As Long = 102
Private Const DispenseAuto As Long = 201
Private Const DispenseScheduled As Long = 202
Private Const DispenseComplete As Long = 301
Private Const DispenseManual As Long = 302
Private Const DispenseError As Long = 309

Public Sub Start(FillLevel As Long, DesiredTemp As Long, MixTime As Long, MixingOn As Long, Calloff As Long, StandardTime As Long, DTank As Long)
  AddFillLevel = FillLevel
  DesiredTemperature = DesiredTemp
  If DesiredTemperature > 1800 Then DesiredTemperature = 1800
  AddMixTime = MixTime
  AddMixing = MixingOn
  AddCallOff = Calloff
  OverrunTime = StandardTime
  DispenseTank = DTank
  
  State = WaitIdle
  WaitReadyTimer.Pause
  Timer = 5
  AddDispenseError = False
  
  DispenseDyes = False
  DispenseChems = False
  DispenseDyesOnly = False
  DispenseChemsOnly = False
  DispenseDyesChems = False
  
End Sub

Public Sub Run(ByVal ControlObject As Object)
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode
    Dispensestate = .Dispensestate
    Dispenseproducts = .Dispenseproducts
    
    'Run an advance timer to reset alarms, where necessary
    If Not .IO_RemoteRun Then AdvanceTimer.TimeRemaining = 2
    If AdvanceTimer.Finished Then AlarmRedyeIssue = False
  
  Do
    'Remember state and loop until state does not change
    'This makes sure we go through all the state changes before proceeding and setting IO
     StatePrev = State
  
  Select Case State
  
    Case Off
      StateString = ""
      DrugroomDisplay = "Tank 1 Idle"
      
    Case WaitIdle    'Wait for destination tank to be idle - no active drains...
      DrugroomDisplay = "Wait For Destination Idle "
      If AlarmRedyeIssue And (.Parameters_EnableRedyeIssueAlarm = 1) Then
        Timer = 5
        StateString = "Check Tank (Redye Issue), Hold Run to reset " & TimerString(AdvanceTimer.TimeRemaining)
        DrugroomDisplay = "Check Tank (Redye Issue), Hold Run to Reset " & TimerString(AdvanceTimer.TimeRemaining)
      ElseIf (DispenseTank = 1) And (.AD.IsOn Or .ManualAddDrain.IsActive) Then
        Timer = 5
        StateString = "Wait to Add - " & .AD.StateString
      ElseIf (DispenseTank = 2) And (.RD.IsOn Or .ManualReserveDrain.IsActive) Then
        Timer = 5
        StateString = "Wait to Reserve - " & .RD.StateString
      Else
        StateString = "Wait for Destination Idle "
      End If
      If Timer.Finished Then State = PreFill

    Case PreFill
      StateString = "Tank 1 PreFill to " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters_Tank1PreFillLevel / 10, "0", 3) & "%"
      DrugroomDisplay = "Tank 1 PreFill to " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters_Tank1PreFillLevel / 10, "0", 3) & "%"
      If (.Tank1Level >= .Parameters_Tank1PreFillLevel) Then
         State = DispenseWaitTurn
      End If
      
    Case DispenseWaitTurn    'Wait for other tank to be finished with dispensing before new dispense
      StateString = "Wait for Turn "
      DrugroomDisplay = "Wait for Turn"
      If (AddCallOff = 0) Or (.Parameters_DispenseEnabled <> 1) Then
        HeatPrepTimer = .Parameters_Tank1HeatPrepTimer
        State = Fill
        ManualAdd = True
      Else
        DispenseTimer = .Parameters_DispenseReadyDelayTime * 60
        State = DispenseWaitReady
        DispenseCalloff = 0
      End If
      
    Case DispenseWaitReady    'Wait to make sure dispenser has completed previous dispense and is ready
      StateString = "Wait for Dispenser Ready "
      DrugroomDisplay = "Wait for Dispenser Ready "
      If Dispensestate = DispenseReady Then                                   '101
         State = DispenseWaitProducts
         DispenseTimer = .Parameters_DispenseResponseDelayTime * 60
         DispenseCalloff = AddCallOff
         .DispenseTank = DispenseTank
      End If
      If (AddCallOff = 0) Or (.Parameters_DispenseEnabled <> 1) Then
        HeatPrepTimer = .Parameters_Tank1HeatPrepTimer
        State = Fill
        DispenseCalloff = 0
        ManualAdd = True
      End If
          
    Case DispenseWaitProducts     'Wait for Dispensestate & DispenseProducts String to be set to determine how to signal delays
      StateString = "Wait for Dispense Products "
      DrugroomDisplay = "Wait for Dispense Products "
      DispenseTimer = .Parameters_DispenseResponseDelayTime * 60
      If Dispensestate <> DispenseReady Then                '(DispenseBusy = 102) but if these no recipe, we'll get a (DispenseManual = 302)
        If Dispenseproducts <> "" Then
          'split the products
          'products() = "Step <calloff>| <Ingredient_id> : <Amount> <Units> <DResult> | <Ingredient_Desc> "
          Dim ProductsArray() As String, i As Long
          ProductsArray = Split(Dispenseproducts, "|")
          For i = 1 To UBound(ProductsArray)                'Disregard the 0 row "Step <calloff>
            Dim position As Long
            position = InStr(ProductsArray(i), ":")         'Need to verify ":" is in the current row due to second split "|"
            If position > 0 Then
              Dim test As String
              test = Mid(ProductsArray(i), 6, 1)
              If test = ":" Then                            'If ":" is at 5th position, we have a chemical => "| 1004: ..."
                DispenseChems = True
              Else
                DispenseDyes = True
              End If
            End If
          Next i
          If DispenseChems And DispenseDyes Then
            DispenseDyesChems = True
          Else
            If DispenseChems Then
              DispenseChemsOnly = True
            Else: DispenseDyesOnly = True
            End If
          End If
        End If
        'Proceed to the next state
        State = DispenseWaitResponse
        DispenseCalloff = AddCallOff
        .DispenseTank = DispenseTank
      End If
      If (AddCallOff = 0) Or (.Parameters_DispenseEnabled <> 1) Then
        HeatPrepTimer = .Parameters_Tank1HeatPrepTimer
        State = Fill
        DispenseCalloff = 0
        ManualAdd = True
      End If
      
    Case DispenseWaitResponse     'Wait for response from dispenser
      StateString = "Wait For Reponse From Dispenser "
      DrugroomDisplay = "Wait for Response From Dispenser "
      AddRecipeStep = Dispenseproducts
      Select Case Dispensestate
         Case DispenseComplete                                                '301
          If Not .WK.IsOn Then
            If DispenseTank = 2 Then
             .ReserveReady = True
            ElseIf DispenseTank = 1 Then
             .AddReady = True
            End If
          End If
          .KR.ACCommand_Cancel
          If .LA.KP1.IsOn Then .LAActive = True
          State = Off
          Cancel
          .DispenseTank = 0
          DispenseTank = 0
          DispenseCalloff = 0

        Case DispenseManual                                                  '302
          HeatPrepTimer = .Parameters_Tank1HeatPrepTimer
          State = Fill
          DispenseCalloff = 0
          ManualAdd = True

        Case DispenseError                                                   '309
          AddDispenseError = True
          HeatPrepTimer = .Parameters_Tank1HeatPrepTimer
          State = Fill
          DispenseCalloff = 0

      End Select
     
    Case Fill
      StateString = "Tank 1 Fill to " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(AddFillLevel / 10, "0", 3) & "%"
      DrugroomDisplay = "Fill to " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(AddFillLevel / 10, "0", 3) & "%"
      If .IO_Tank1Manual_SW Then State = InManual
      If AddFillLevel = 0 Then
        WaitReadyTimer.Start
        State = Slow
      End If
      If .Tank1Level > AddFillLevel Then State = Heat
      
    Case Heat
      If Not .DrugroomMixerOn Then
        StateString = "Tank 1 Level Too Low to Heat " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters_Tank1MixerOnLevel / 10, "0", 3) & "%"
        DrugroomDisplay = "Level Too Low to Heat " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters_Tank1MixerOnLevel / 10, "0", 3) & "%"
      Else
        StateString = "Tank 1 Heat to " & Pad(.IO_Tank1Temp / 10, "0", 3) & "F / " & Pad(DesiredTemperature / 10, "0", 3)
        DrugroomDisplay = "Heat to " & Pad(.IO_Tank1Temp / 10, "0", 3) & "F / " & Pad(DesiredTemperature / 10, "0", 3)
      End If
      If .IO_Tank1Manual_SW Then State = InManual
      If AddFillLevel = 0 Then State = Slow
      ' Heat Start/Stop
      If .IO_Tank1Temp < DesiredTemperature Then
        HeatOn = True
        AddMixing = True ' Must mix if heating
      End If
      If .IO_Tank1Temp > DesiredTemperature Then
        HeatOn = False
        State = Slow
        WaitReadyTimer.Start
        OverrunTimer = OverrunTime
      End If
      
    Case Slow
      StateString = "Waiting on Tank 1 " & TimerString(WaitReadyTimer.TimeElapsed)
      DrugroomDisplay = "Prepare Tank 1"
      If .IO_Tank1Manual_SW Then State = InManual
      ' Heat Start/Stop
      If .IO_Tank1Temp < DesiredTemperature Then
        HeatOn = True
        AddMixing = True ' Must mix if heating
      End If
      If .IO_Tank1Temp >= DesiredTemperature Then HeatOn = False
      If DispenseTank = 2 Then
        'Reserve Tank
        If .RT.IsWaitReady Then State = Fast
      ElseIf DispenseTank = 1 Then
        'Add Tank
        If .AC.IsWaitReady Then State = Fast
        If .AT.IsWaitReady Then State = Fast
      End If
      If .Tank1Ready Then
        WaitReadyTimer.Pause
        State = MixForTime
        MixTimer = AddMixTime
      End If
      If WaitReadyTimer.IsPaused Then WaitReadyTimer.Restart 'should cover LA copy to KP function
      If OverrunTimer.Finished Then
         State = Fast
      End If
    
    Case Fast
      StateString = "Waiting on Tank 1 " & TimerString(WaitReadyTimer.TimeElapsed)
      DrugroomDisplay = "Prepare Tank 1"
      If .IO_Tank1Manual_SW Then State = InManual
      ' Heat Start/Stop
      If .IO_Tank1Temp < DesiredTemperature Then
        HeatOn = True
        AddMixing = True ' Must mix if heating
      End If
      If .IO_Tank1Temp >= DesiredTemperature Then HeatOn = False
      If WaitReadyTimer.IsPaused Then WaitReadyTimer.Restart 'should cover LA copy to KP function
      If .Tank1Ready Then
        If .Tank1Level Then
          WaitReadyTimer.Pause
          State = MixForTime
          MixTimer = AddMixTime
        Else
          StateString = "Tank 1 Level low " & Pad(.Tank1Level / 10, "0", 3) & "%"
        End If
      End If

      
    Case MixForTime
      If Not .DrugroomMixerOn Then
        StateString = "Tank 1 Level Too Low to Mix " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters_Tank1MixerOnLevel / 10, "0", 3) & "%"
        DrugroomDisplay = "Level Too Low to Mix " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters_Tank1MixerOnLevel / 10, "0", 3) & "%"
      Else
        StateString = "Tank 1 Mix For Time " & TimerString(MixTimer.TimeRemaining)
        DrugroomDisplay = "Mix For Time " & TimerString(MixTimer.TimeRemaining)
      End If
      If .IO_Tank1Manual_SW Then State = InManual
      If AddFillLevel = 0 Then State = Ready
      ' Heat Start/Stop
      If .IO_Tank1Temp < DesiredTemperature Then
        HeatOn = True
        AddMixing = True ' Must mix if heating
      End If
      If .IO_Tank1Temp >= DesiredTemperature Then HeatOn = False
      If MixTimer.Finished Then State = Ready
      If Not .Tank1Ready Then
        State = Fast
        WaitReadyTimer.Restart
      End If
      
    Case Ready
      StateString = "Tank 1 Ready "
      DrugroomDisplay = "Ready "
      If .IO_Tank1Manual_SW Then State = InManual
      If (DispenseTank = 1) And (.AdditionLevel > .Parameters_AdditionMaxTransferLevel) Then
         State = KAInterlock
      Else
        State = KATransfer1
      End If
      FillOn = False
      HeatOn = False
      AddMixing = False
      AddDispenseError = False
      ManualAdd = False

    Case KAInterlock
      StateString = "Add Level High for Tank 1 Transfer " & Pad(.AdditionLevel / 10, "0", 3) & "%"
      DrugroomDisplay = "Add Level High for Tank 1 Transfer " & Pad(.AdditionLevel / 10, "0", 3) & "%"
      If .IO_Tank1Manual_SW Then State = InManual
      If (.AdditionLevel <= .Parameters_AdditionMaxTransferLevel) Then
        State = KATransfer1
      End If
    
    Case KATransfer1
      If (.Tank1Level > 10) Then
        Timer = .Parameters_Tank1TimeBeforeRinse
        StateString = "Tank 1 Transfer Empty " & Pad(.Tank1Level / 10, "0", 3) & "%"
        DrugroomDisplay = "Transfer Empty " & Pad(.Tank1Level / 10, "0", 3) & "%"
      Else
        StateString = "Tank 1 Transfer Empty " & TimerString(Timer.TimeRemaining)
        DrugroomDisplay = "Transfer Empty " & TimerString(Timer.TimeRemaining)
      End If
      If .IO_Tank1Manual_SW Then State = InManual
      If Timer.Finished Then
        NumberOfRinses = .Parameters_DrugroomRinses
        If .KR.IsOn Then
          If (.KR.RinseMachine = 0) And (.KR.RinseDrain = 0) Then
            If Not .WK.IsOn Then
              If (DispenseTank = 1) Then .AddReady = True
              If (DispenseTank = 2) Then .ReserveReady = True
            End If
            If .LA.KP1.IsOn Then .LAActive = True
            DispenseTank = 0
            .DispenseTank = 0
            .Tank1Ready = False
            .KR.ACCommand_Cancel
            Cancel
          ElseIf .KR.RinseMachine = 0 Then
            State = KARinseToDrain
            Timer = .Parameters_Tank1RinseToDrainTime
          End If
        Else
          State = KARinse
          Timer = .Parameters_Tank1RinseTime
        End If
      End If
            
    Case KARinse
      StateString = "Tank 1 Rinse " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters_Tank1RinseLevel / 10, "0", 3) & "%"
      DrugroomDisplay = "Rinse " & Pad(.Tank1Level / 10, "0", 3) & "% / " & Pad(.Parameters_Tank1RinseLevel / 10, "0", 3) & "%"
      If .IO_Tank1Manual_SW Then State = InManual
      If (.Tank1Level > .Parameters_Tank1RinseLevel) Then
        State = KATransfer2
        Timer = .Parameters_Tank1TimeAfterRinse
      End If
      
    Case KATransfer2
      If (.Tank1Level > 10) Then
        Timer = .Parameters_Tank1TimeAfterRinse
        StateString = "Tank 1 Transfer Rinse " & Pad(.Tank1Level / 10, "0", 3) & "%"
        DrugroomDisplay = "Transfer Rinse " & Pad(.Tank1Level / 10, "0", 3) & "%"
      Else
        StateString = "Tank 1 Transfer Rinse " & TimerString(Timer.TimeRemaining)
        DrugroomDisplay = "Transfer Rinse " & TimerString(Timer.TimeRemaining)
      End If
      If .IO_Tank1Manual_SW Then State = InManual
      If Timer.Finished Then
        If Not .WK.IsOn Then
          If (DispenseTank = 1) Then .AddReady = True
          If (DispenseTank = 2) Then .ReserveReady = True
        End If
        If .LA.KP1.IsOn Then .LAActive = True
        DispenseTank = 0
        .DispenseTank = 0
        .Tank1Ready = False
        State = KARinseToDrain
        Timer = .Parameters_Tank1RinseToDrainTime
      End If
      
    Case KARinseToDrain
      StateString = "Tank 1 Rinse To Drain " & TimerString(Timer.TimeRemaining)
      DrugroomDisplay = "Rinse To Drain " & TimerString(Timer.TimeRemaining)
      If .IO_Tank1Manual_SW Then State = InManual
      If Timer.Finished Or (.Tank1Level >= 500) Then
        State = KATransferToDrain
        Timer = .Parameters_Tank1DrainTime
        NumberOfRinses = NumberOfRinses - 1
      End If
    
    Case KATransferToDrain
      If (.Tank1Level > 10) Then
        Timer = .Parameters_Tank1DrainTime
        StateString = "Tank 1 To Drain " & Pad(.Tank1Level / 10, "0", 3) & "%"
        DrugroomDisplay = "Transfer To Drain " & Pad(.Tank1Level / 10, "0", 3) & "%"
      Else
        StateString = "Tank 1 To Drain " & TimerString(Timer.TimeRemaining)
        DrugroomDisplay = "Transfer To Drain " & TimerString(Timer.TimeRemaining)
      End If
      If .IO_Tank1Manual_SW Then State = InManual
      If Timer.Finished Then
        If NumberOfRinses > 0 Then
          State = KARinseToDrain
          Timer = .Parameters_Tank1RinseToDrainTime
        Else
          If Not .WK.IsOn Then
            If (DispenseTank = 1) Then .AddReady = True
            If (DispenseTank = 2) Then .ReserveReady = True
          End If
          If .LA.KP1.IsOn Then .LAActive = True
          DispenseTank = 0
          .DispenseTank = 0
          .Tank1Ready = False
          .KR.ACCommand_Cancel
          Cancel
        End If
      End If
         
    Case InManual
      StateString = "Drugroom Switch In Manual "
      DrugroomDisplay = "Switch In Manual"
      .Tank1Ready = False
      HeatPrepTimer = .Parameters_Tank1HeatPrepTimer
      If Not .IO_Tank1Manual_SW Then
        HeatPrepTimer = .Parameters_Tank1HeatPrepTimer
        State = Fill
      End If
    
  End Select
  
  Loop Until (StatePrev = State)    'Loop until state does not change
  
  End With
End Sub

Public Sub Cancel()
  FillOn = False
  State = Off
  HeatOn = False
  AddMixing = False
  AddDispenseError = False
  WaitReadyTimer.Pause
  AlarmRedyeIssue = False
  ManualAdd = False

  'This is to clear prepare properties and hopefully resolve issues in LA where wrong calloff is used
  AddCallOff = 0
  AddFillLevel = 0
  AddMixTime = 0
  AddRecipeStep = ""
  DesiredTemperature = 0
  DispenseCalloff = 0
  DispenseTank = 0
  Dispenseproducts = ""
  Dispensestate = 0
        
  DispenseDyes = False
  DispenseChems = False
  DispenseDyesOnly = False
  DispenseChemsOnly = False
  DispenseDyesChems = False
  
End Sub

Friend Sub ProgramStart()
On Error Resume Next
  Dispenseproducts = ""
  DispenseCalloff = 0
 ' DispenseTank = 0
  AddCallOff = 0
  AddRecipeStep = ""
  Dim i1 As Long, i2 As Long
  For i1 = 1 To 64
    For i2 = 1 To 8
      RecipeSteps(i1, i2) = ""
    Next i2
  Next i1
End Sub

Public Sub CopyTo(Target As acAddPrepare)
  With Target
    .State = State
    .AddFillLevel = AddFillLevel
    .AddCallOff = AddCallOff
    .OverrunTimer.TimeRemaining = OverrunTimer.TimeRemaining
    .AddMixing = AddMixing
    .DesiredTemperature = DesiredTemperature
    .AddMixTime = AddMixTime
    .OverrunTime = OverrunTime
    .MixTimer.TimeRemaining = MixTimer.TimeRemaining
    .WaitReadyTimer.TimeElapsed = WaitReadyTimer.TimeElapsed
        
    .ManualAdd = ManualAdd
    .AlarmRedyeIssue = AlarmRedyeIssue
    .AddDispenseError = AddDispenseError
    .AddRecipeStep = AddRecipeStep
    .DispenseCalloff = DispenseCalloff
    .DispenseTank = DispenseTank
    .Dispensestate = Dispensestate
    .Dispenseproducts = Dispenseproducts
    .Timer.TimeRemaining = Timer.TimeRemaining
    
    .DrugroomDisplay = DrugroomDisplay
    .DispenseDyesChems = DispenseDyesChems
    .DispenseDyesOnly = DispenseDyesOnly
    .DispenseChemsOnly = DispenseChemsOnly
    .DispenseTimer.TimeRemaining = DispenseTimer.TimeRemaining
    .MixTimer.TimeRemaining = MixTimer.TimeRemaining
        
  End With
End Sub

Public Property Get IsOn() As Boolean
  IsOn = (State <> Off)
End Property
Public Property Get IsWaitingToDispense() As Boolean
  IsWaitingToDispense = (State = DispenseWaitTurn)
End Property
Public Property Get IsDispensing() As Boolean
  IsDispensing = (State = DispenseWaitReady) Or (State = DispenseWaitProducts) Or (State = DispenseWaitResponse)
End Property
Friend Property Get IsDispenseWaitReady() As Boolean
  IsDispenseWaitReady = (State = DispenseWaitReady)
End Property
Friend Property Get IsDispenseReadyOverrun() As Boolean
  IsDispenseReadyOverrun = IsDispenseWaitReady And DispenseTimer.Finished
End Property
Friend Property Get IsDispenseWaitResponse() As Boolean
  IsDispenseWaitResponse = (State = DispenseWaitResponse)
End Property
Friend Property Get IsDispenseResponseOverrun() As Boolean
  IsDispenseResponseOverrun = IsDispenseWaitResponse And DispenseTimer.Finished
End Property

Public Property Get IsFill() As Boolean
  IsFill = ((State = Fill) Or FillOn) And (Not State = InManual)
End Property
Public Property Get IsHeating() As Boolean
  IsHeating = (HeatOn And ((State = Heat) Or (State = Slow) Or (State = Fast))) And _
                (Not State = InManual)
End Property
Friend Property Get IsHeatTankOverrun() As Boolean
  IsHeatTankOverrun = ((State = Fill) Or (State = Heat)) And HeatPrepTimer.Finished
End Property
Public Property Get IsSlow() As Boolean
  IsSlow = (State = Slow)
End Property
Public Property Get IsFast() As Boolean
  IsFast = (State = Fast)
End Property
Friend Property Get IsWaitReady() As Boolean
  IsWaitReady = (State = Slow) Or (State = Fast) Or (State = MixForTime)
End Property
Public Property Get IsMixing() As Boolean
  IsMixing = (State = MixForTime)
End Property
Public Property Get IsReady() As Boolean
  IsReady = (State = Ready)
End Property
Public Property Get IsMixerOn() As Boolean
  IsMixerOn = AddMixing And Not ((State = InManual) Or IsDispensing Or IsWaitingToDispense Or IsWaitIdle)
End Property
Public Property Get IsOverrun() As Boolean
  IsOverrun = IsWaitReady And OverrunTimer.Finished
End Property
Friend Property Get IsInterlocked() As Boolean
 IsInterlocked = (State = KAInterlock)
End Property
Friend Property Get IsTransfer() As Boolean
  IsTransfer = (State = KATransfer1) Or (State = KATransfer2) Or (State = KARinseToDrain) Or (State = KATransferToDrain)
End Property
Friend Property Get IsTransferToAddition() As Boolean
 IsTransferToAddition = ((State = KATransfer1) Or (State = KARinse) Or _
                       (State = KATransfer2)) And (DispenseTank = 1)
End Property
Friend Property Get IsTransferToReserve() As Boolean
 IsTransferToReserve = ((State = KATransfer1) Or (State = KARinse) Or _
                       (State = KATransfer2)) And (DispenseTank = 2)
End Property
Friend Property Get IsTransferToDrain() As Boolean
 IsTransferToDrain = IsRinseToDrain Or (State = KATransferToDrain)
End Property
Friend Property Get IsRinse() As Boolean
  If (State = KARinse) Then IsRinse = True
End Property
Friend Property Get IsRinseToDrain() As Boolean
  If (State = KARinseToDrain) Then IsRinseToDrain = True
End Property
Friend Property Get IsPaused() As Boolean
  IsPaused = (State = KAPause)
End Property
Friend Property Get IsWaitIdle() As Boolean
  IsWaitIdle = (State = WaitIdle)
End Property
Friend Property Get IsDelayed() As Boolean
  IsDelayed = (IsWaitReady And IsOverrun)
End Property
Friend Property Get IsInManual() As Boolean
  IsInManual = (State = InManual)
End Property

#End If

End Class
