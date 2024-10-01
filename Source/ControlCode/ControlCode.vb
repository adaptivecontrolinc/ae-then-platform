Imports Utilities.Translations

Public Class ControlCode : Inherits MarshalByRefObject : Implements ACControlCode

  Public Parent As ACParent
  Public Parameters As Parameters
  Public Alarms As Alarms
  Public IO As IO
  '      Public Analogs as Analogs    ' Check ML code for usage

  Public ExpertModeTimer As New Timer With {.Minutes = 60}

  Public BootTime As Date = Date.Now
  Public CurrentTime As String ' TODO Check and remove
  Public ComputerName As String
  'Public MachineSize As Integer
  'Public MachineType As String
  'Public Batch As BatchData

  Public TemperatureControl As Temperature
  ' ?? Public TemperatureControlContacts As New acTempControlContacts

  Public PumpControl As PumpControl
  Public LevelFlush As LevelFlush
  Public LidControl As LidControl

  Public Simulation As Simulation

  ' Drugroom Preview Variables
  Public StepNumberWas As Integer
  Public StepOverrunMins As Integer
  Public TimeInStepValueWas As Integer
  Public DrugroomPreview As DrugroomPreview
  Public DrugroomCalloffRefreshed As Boolean


  ' KitchenPreview  ' TODO Instead?

  Public MimicBatch As String           ' Declare these here so the remote Mimic can read the data
  Public MimicBatchNotes As String      '
  Public MimicRecipeSteps As String     '
  Public MimicRecipe As String          '

  ' Automatic Dispense Variables
  Public DrugroomDisplayJob1 As String
  Public DispenseCallOff As Integer
  Public DispenseState As Integer        'This filled in by AutoDispenser
  Public DispenseProducts As String
  Public DispenseTank As Integer
  Public DispenseTankWas As Integer

  Public Sub New(ByVal parent As ACParent)
    Me.Parent = parent

    Parameters = New Parameters(Me)
    IO = New IO(Me)
    Alarms = New Alarms(Me)

    ' Load settings first so we can override anything (settings.xml)
    Settings.Load()

    CultureName = parent.CultureName  ' make our translations in the same language as batchcontrol is doing it

    ' Initialize the local database
    InitializeLocalDatabase()

    'Reset Demo Parameters
    Parameters.Demo = 0

    ComputerName = My.Computer.Name

    ' Initialize Command Modules
    AddControl = New AddControl(parent, Me)
    ReserveControl = New ReserveControl(parent, Me)
    ManualSteamTest = New ManualSteamTest(parent, Me)

    TemperatureControl = New Temperature(Me)
    PumpControl = New PumpControl(Me)
    LevelFlush = New LevelFlush(Me)
    LidControl = New LidControl(Me)

    DrugroomPreview = New DrugroomPreview(Me)

    Simulation = New Simulation(Me)

  End Sub


#Region " DATABASE INITIALIZATION "

  Private Sub InitializeLocalDatabase()
    'Update the local database tables with the necessary recipe field mods if they do not already exist
    Dim sql As String = Nothing
    Try

      ' If the dyelots table does not yet exist (because we have just created a blank database),
      ' then create it ourself for the recipe information.

      If Parent.DbGetDataTable("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='Dyelots'").Rows.Count = 0 Then
        sql = "CREATE TABLE Dyelots(" &
                       "  Dyelot                nvarchar(50) NOT NULL," &
                       "  Redye                 int NOT NULL," &
                       "  Machine               nvarchar(50)," &
                       "  StartTime             datetime," &
                       "  EndTime               datetime," &
                       "  StandardTime          float," &
                       "  EarliestStartTime     datetime," &
                       "  State                 smallint," &
                       "  Blocked               smallint," &
                       "  Committed             smallint," &
                       "  Program               ntext," &
                       "  Parameters            ntext," &
                       "  Notes                 ntext," &
                       "  Color                 int," &
                       "  HistoryUploaded       smallint," &
                       "  ExternalOrder         int," &
                       "  RecipeID              int," &
                       "  RecipeCode            nvarchar(50)," &
                       "  RecipeName            nvarchar(250)," &
                       "  StyleID               int," &
                       "  StyleCode             nvarchar(50)," &
                       "  StyleName             nvarchar(250)," &
                       "  Batched               datetime," &
                       "  PRIMARY KEY (Dyelot, Redye));"
        Parent.DbExecute(sql)
      End If

      ' Customize ProgramGroups to include MaximumGradient field - used by BatchControl to limit the StandardTime associated with gradient of 9.9
      '   sql = "CREATE TABLE ProgramGroups (MaximumGradient int)"
      '   Parent.DbUpdateSchema(sql)

      ' Update MaximumGradient Value ((Was '30'))
      '    sql = "UPDATE ProgramGroups SET MaximumGradient=30 WHERE Name='ThenPlatform'"
      '    Parent.DbExecute(sql)

      ' Machines Table
      '    sql = "CREATE TABLE Machines (Locked SMALLINT)"
      '    Parent.DbUpdateSchema(sql)

    Catch ex As Exception
      Utilities.Log.LogError(ex, sql)
    End Try
  End Sub

#End Region

  Public Sub Run() Implements ACControlCode.Run
    Try


      '------------------------------------------------------------------------------------------
      ' Run simulation 
      '------------------------------------------------------------------------------------------
      If Settings.Demo >= 1 Then
        Parent.Mode = Mode.Debug
        Parameters.Demo = 1
      End If
      If (Parent.Mode = Mode.Debug) AndAlso Parameters.Demo > 0 Then Simulation.Run()


      '------------------------------------------------------------------------------------------
      ' Flashers for alarms and prepare lamps
      '------------------------------------------------------------------------------------------
      Static FastFlasher As New Flasher(400) : FlashFast = FastFlasher.On
      Static SlowFlasher As New Flasher(800) : FlashSlow = SlowFlasher.On



      'Set program state change timers and determine Time-In-Step variables
      CheckProgramStateChanges()



      '------------------------------------------------------------------------------------------
      ' Some useful status flags 
      '------------------------------------------------------------------------------------------
      EStop = IO.EmergencyStop OrElse (Not IO.PowerOn)
      Dim NStop As Boolean = Not EStop
      Dim NHalt As Boolean = (Not Parent.IsPaused) AndAlso NStop
      Dim NHSafe As Boolean = NHalt And MachineSafe

      ' Use Temperature probe with highest value (for safety) unless Cooling is active
      Temp = IO.VesselTemp

      ' SAFETY FLAGS
      TempValid = (Temp > 300) AndAlso (Temp < 3000)
      '  TempValid = (Temp > 300) And (Temp < 3000) And IO_MainPumpStart And IO_MainPumpRunning And (Not IO_ContactThermometer)

      '  TempSafe = (Temp < 600) AndAlso IO.TempInterlock AndAlso Not IO.ContactTemp AndAlso (Temp < SafetyControl.DepressurizeTemp)
      TempSafe = (Temp < 2050) And IO.TempInterlock And Not IO.ContactTemp AndAlso (Temp < SafetyControl.DepressurizeTemp)
      '  VesselHeatSafe = (Not IO_ContactThermometer) And IO_MainPumpRunning And (VesTemp < 2800)


      PressSafe = IO.PressureInterlock AndAlso ((MachinePressure <= 50) OrElse RF.IsRunback) ' system pressure in tenths bar (4psi = 0.27579bar)
      ' .NET PressSafe = IO.PressureInterlock AndAlso ((MachinePressure <= 2758) OrElse RF.IsRunback) ' system pressure in tenths bar (4psi = 0.27579bar)
      MachineSafe = SafetyControl.IsDePressurized AndAlso TempSafe ' andalso PressSafe


      Static machineLidClosedTimer As New Timer
      If Not (IO.KierLidClosed OrElse IO.LockingPinRaised OrElse FirstScanDone) Then machineLidClosedTimer.Seconds = 5
      MachineClosed = IO.KierLidClosed AndAlso IO.LockingPinRaised AndAlso machineLidClosedTimer.Finished


      'Temperature Control 
      HeatEnabled = NHalt AndAlso TempValid AndAlso (Temp < 2800) AndAlso MachineClosed AndAlso IO.MainPump AndAlso
                                  IO.PumpRunning AndAlso (Not Alarms.MainPumpAlarm) AndAlso Not IO.ContactTemp

      TemperatureControl.Run(Temp, HeatEnabled)
      TemperatureControl.EnableDelay = 10
      Setpoint = TemperatureControl.Pid.PidSetpoint
      TempFinalValue = (TemperatureControl.Pid.FinalTemp \ 10).ToString("#0") & "F"


      'TODO - Add temperature control for heat by contacts - TC.Gradient
      '   TemperatureControlContacts.EnableDelay = 10
      '   TemperatureControlContacts.Run VesTemp, TC.Gradient, Me
      '   TemperatureControlContacts.CheckErrorsAndMakeAlarms VesTemp, TC.Gradient, Me


      ' Run the safety control to handle pressurisation
      SafetyControl.Run(Temp, TempSafe, PressSafe, PR.IsActive OrElse AirpadOn)


      '------------------------------------------------------------------------------------------
      ' Analog input calibration
      '------------------------------------------------------------------------------------------
      ' Level transmitters
      MachineLevel = Utilities.Calibrate.CalibrateAnalogInput(IO.VesselLevelInput, Parameters.VesselLevelTransmitterMin, Parameters.VesselLevelTransmitterMax)
      AddLevel = Utilities.Calibrate.CalibrateAnalogInput(IO.AddLevelInput, Parameters.AddLevelTransmitterMin, Parameters.AddLevelTransmitterMax)
      ReserveLevel = Utilities.Calibrate.CalibrateAnalogInput(IO.ReserveLevelInput, Parameters.ReserveLevelTransmitterMin, Parameters.ReserveLevelTransmitterMax)


      Tank1Level = Utilities.Calibrate.CalibrateAnalogInput(IO.Tank1LevelInput, Parameters.Tank1LevelTransmitterMin, Parameters.Tank1LevelTransmitterMax)




      ' Issue with signal noise due to tank 1 mixer running causing interference with UV Level transmitter
      Static tank1LevelWas As Integer
      ' Set delay off timer to prevent false signal as soon as mixer stops (takes a couple second for UV transmitter to restore proper signal)
      If IO.Tank1Mixer Then ' OrElse IO.Tank1MixerRunning Then
        ' Use last scan's level
        Tank1Level = tank1LevelWas

        ' Check level against last level at interval scans
        If tank1LevelScanTimer.Finished Then
          Dim deltaTank1Level As Integer = Math.Abs(Tank1Level - tank1LevelWas)
          ' Consider the level change delta since last scan
          If ((deltaTank1Level > Parameters.Tank1LevelDisregardDelta) AndAlso (Parameters.Tank1LevelDisregardDelta > 0)) OrElse (Tank1Level <= deltaTank1Level) Then
            Tank1Level = tank1LevelWas
          Else
            tank1LevelWas = Tank1Level
          End If
        End If
        ' Reset Mixer Active Timer
        tank1MixerActiveTimer.Seconds = 2
      Else
        If tank1MixerActiveTimer.Finished Then
          tank1LevelWas = Tank1Level
        End If
        ' Reset Level Scan TImer
        tank1LevelScanTimer.Seconds = 2
      End If

      ' Differential Pressure across the packages
      If (IO.DiffPressInput >= Parameters.PackageDiffPressZero) Then
        PackagePressureRange = Parameters.PackageDiffPress100 - Parameters.PackageDiffPressZero
        PackagePressureUncorrected = IO.DiffPressInput - Parameters.PackageDiffPressZero
      Else
        PackagePressureRange = Parameters.PackageDiffPressZero
        PackagePressureUncorrected = Parameters.PackageDiffPressZero - IO.DiffPressInput
      End If
      If PackagePressureRange > 0 Then PackageDifferentialPressure = MulDiv(PackagePressureUncorrected, Parameters.PackageDiffPressRange, PackagePressureRange)
      If PackageDifferentialPressure < 0 Then PackageDifferentialPressure = 0
      If PackageDifferentialPressure > Parameters.PackageDiffPressRange Then PackageDifferentialPressure = 1000
      '  PackageDpPsi = Convert.ToInt16((PackageDifferentialPressure / 100) * Utilities.Conversions.BarToPsi)
      PackageDpStr = (PackageDifferentialPressure / 10).ToString("#0.0") & " psi"


      ' Machine Pressure transmitter
      MachinePressure = Utilities.Calibrate.CalibrateAnalogInput(IO.VesselPressInput, Parameters.VesselPressureMin, Parameters.VesselPressureMax, Parameters.VesselPressureRange)
      MachinePressureDisplay = (MachinePressure / 10).ToString("#0.0") & "psi"


      ' [2016-05-19] Added Machine Pressure Max flag to record the maximum pressure that a machine sees during a program (resets at program start)
      '              Used for troubleshooting against Airpad/Airpad sequence
      If MachinePressure > MachinePressureMax Then
        MachinePressureMax = MachinePressure
      End If


      ' Flow Rate
      FlowRate = Utilities.Calibrate.CalibrateAnalogInput(IO.FlowRateInput, Parameters.FlowRateMachineMin, Parameters.FlowRateMachineMax, Parameters.FlowRateMachineRange)
      If Parameters.FlowRateMachineRange > 0 Then
        FlowRatePercent = CInt(((FlowRate / Parameters.FlowRateMachineRange) * 1000))
      Else : FlowRatePercent = 0
      End If


      ' Aninp8 Test Value
      TestValue = Utilities.Calibrate.CalibrateAnalogInput(IO.TestInput, Parameters.TestTransmitterMin, Parameters.TestTransmitterMax, Parameters.TempTransmitterRange)


      ' Tank 1 POTs
      Tank1TimePot = Utilities.Calibrate.CalibrateAnalogInput(IO.TimePotInput, Parameters.Tank1TimePotMin, Parameters.Tank1TimePotMax, Parameters.Tank1TimePotRange)
      Tank1TempPot = Utilities.Calibrate.CalibrateAnalogInput(IO.TempPotInput, Parameters.Tank1TempPotMin, Parameters.Tank1TempPotMax, Parameters.Tank1TempPotRange)



      '------------------------------------------------------------------------------------------
      ' Level Flush Control
      '------------------------------------------------------------------------------------------
      LevelFlush.Run()


      '------------------------------------------------------------------------------------------
      ' Circulation Flowmeter input calibration
      '------------------------------------------------------------------------------------------
      MachineFlowRatePv = Utilities.Calibrate.CalibrateAnalogInput(IO.FlowRateInput, Parameters.FlowRateMachineMin, Parameters.FlowRateMachineMax, Parameters.FlowRateMachineRange)


      '------------------------------------------------------------------------------------------
      ' Lid Control
      '------------------------------------------------------------------------------------------
      LidLocked = IO.LockingPinRaised AndAlso IO.KierLidClosed
      LevelOkToOpenLid = (MachineLevel <= Parameters.LevelOkToOpenLid)
      LidControl.Run()


#If 0 Then
      
'lid locked
  LidLocked = IO_LockingPinRaised And IO_LidLoweredSwitch
  If (VesselLevel <= Parameters_LevelOkToOpenLid) Then
     LevelOkToOpenLid = True
  Else
     LevelOkToOpenLid = False
  End If
          
'lid control stuff
 LidControl.Run MachineSafe, LevelOkToOpenLid, IO_MainPumpStart, _
              IO_LidLoweredSwitch, IO_LockingPinRaised, _
              IO_CloseLid_PB, IO_OpenLid_PB, IO_EStop_PB, _
              Parameters_LockingBandOpeningTime, _
              Parameters_LockingBandClosingTime, _
              Parameters_LidRaisingTime
            
#End If



      '------------------------------------------------------------------------------------------
      ' Pump Control
      '------------------------------------------------------------------------------------------
      PumpControl.Run()


      ' TODO
#If 0 Then
      'turn off pump if turn on off to many times
  If PumpWasRunning And Not IO_MainPumpStart Then
    PumpOnCount = PumpOnCount + 1
    End If
  If IO_MainPumpStart And Not PumpWasRunning Then
    ResetPumpOnCountTimer = 300
  End If
  PumpWasRunning = IO_MainPumpStart
  If ResetPumpOnCountTimer.Finished Then PumpOnCount = 0
  If PumpOnCount >= Parameters_PumpOnOffCountTooHigh Then
    PumpOnCountTooHigh = True
  Else
    PumpOnCountTooHigh = False
  End If
  If IO_RemoteHalt Then
    PumpOnCountTooHigh = False
    PumpOnCount = 0
  End If
#End If


      '------------------------------------------------------------------------------------------
      ' Airpad Control
      '------------------------------------------------------------------------------------------
      ' 100000 = 10000.0mbar = 10.0bar  >>  100 = 10.0mbar
      If (MachinePressure >= (Parameters.AirpadCutoffPressure + 100)) OrElse ((MachinePressure >= (Parameters.ATBleedPressure - 500) AndAlso (AT.IsLowPressure OrElse AC.IsLowPressure))) OrElse HD.IoAirpadCutoff Then
        AirpadCutoff = True
      Else
        AirpadCutoff = False
      End If

      ' Use vent to bleed pressure from machine - use a delay timer to hold vent open for parameter time
      If (MachinePressure > Parameters.AirpadBleedPressure) OrElse ((MachinePressure > Parameters.ATBleedPressure) AndAlso (AT.IsLowPressure OrElse AC.IsLowPressure)) OrElse HD.IoAirpadBleed Then
        AirpadBleed = True
        AirpadBleedDelayOff.Milliseconds = MinMax(Parameters.AirpadBleedTime, 250, 10000)
      Else
        If (AirpadBleedDelayOff.Finished) Then AirpadBleed = False
      End If
      ' Count number of times airpad bleed actuates - cannot see I/O in history
      Static airpadBleedWas As Boolean
      If AirpadBleed AndAlso Not airpadBleedWas Then
        AirpadBleedCount += 1
      End If
      airpadBleedWas = AirpadBleed




      '------------------------------------------------------------------------------------------
      ' Level Control
      '------------------------------------------------------------------------------------------
      ' Bath Volume based on level
      VolumeBasedOnLevel = CInt(Parameters.SystemVolumeAt0Level + (((Parameters.SystemVolumeAt100Level - Parameters.SystemVolumeAt0Level) * MachineLevel) / 1000))





      'Reset Tank volumes
      ' TODO?  If FI.ResetFlowmeter OrElse RI.IsResetMeter OrElse RC.IsResetMeter Then MachineVolume = 0

      ' Calculate the system flow rate in tenths of units/seconds (LitersPerMin/Kilogram)
      If FlowRateTimer.Finished Then
        If (BatchWeight > 0) Then
          FlowratePerWt = CInt((MachineFlowRatePv / BatchWeight) * 10)
        Else : FlowratePerWt = 0
        End If

        ' Calculate number of contacts per min for tc command
        If IO.PumpRunning AndAlso (VolumeBasedOnLevel > 0) Then
          NumberofContactsPerMin = Math.Round((MachineFlowRatePv / VolumeBasedOnLevel), 3)
        Else : NumberofContactsPerMin = 0
        End If

        ' Calculate number of contacts
        SystemVolume = Math.Round(SystemVolume + (MachineFlowRatePv / 60), 3)
        If (NumberofContactsPerMin > 0) AndAlso (SystemVolume >= VolumeBasedOnLevel) Then
          SystemVolume = Math.Round(SystemVolume - VolumeBasedOnLevel, 3)
          NumberOfContacts = NumberOfContacts + 1
          TotalNumberOfContacts = TotalNumberOfContacts + 1
        End If
        FlowRateTimer.Seconds = 1
      End If

      ' figure out the recorded level stuff
      If GetRecordedLevel Then
        If GetRecordedLevelTimer.Finished And Not GotRecordedLevel Then
          GotRecordedLevel = True
          MachineLevelRecorded = MachineLevel
        End If
      Else
        MachineLevelRecorded = 0
        GotRecordedLevel = False
        GetRecordedLevelTimer.Seconds = Parameters.RecordLevelTimer
      End If

      'Idle Drain Control
      If (Parameters.IdleDrainTime > 0) AndAlso (Not Parent.IsProgramRunning) Then
        If ProgramStoppedTimer.Seconds > (Parameters.IdleDrainTime * 60) Then
          IdleDrainActive = True
        Else : IdleDrainActive = False
        End If
      Else : IdleDrainActive = False
      End If

      ' Kier Isolate Function
      If PowerOnTimer.Finished AndAlso FirstScanDone AndAlso Not Parent.IsProgramRunning Then
        ' Default value - open kier isolation valves
        '       If KierBSafeToSwitch AndAlso Not KierBLoaded Then KierBLoaded = True
      End If


      '------------------------------------------------------------------------------------------
      ' Add & Reserve Control
      '------------------------------------------------------------------------------------------
      AddControl.Run()
      ReserveControl.Run()



      ' TANK 1 CONTROL ********************************************************************************************
      ' Kitchen Ready Pushbutton
      Static previousTank1Ready As Boolean
      If IO.Tank1ReadyPb AndAlso Not previousTank1Ready Then
        Tank1Ready = Not Tank1Ready
      End If
      previousTank1Ready = IO.Tank1ReadyPb


      ' Drugroom Mixer Level
      If (Tank1Level > Parameters.Tank1MixerOnLevel) Then Tank1MixerEnable = True
      '  If (Tank1Level <= Parameters_Tank1MixerOffLevel) AndAlso (Not IO.Tank1MixerRunning) Then Tank1MixerEnable = False
      If (Tank1Level <= Parameters.Tank1MixerOffLevel) Then Tank1MixerEnable = False
      If IO.Tank1FillCold Then Tank1MixerEnable = False


      'Toggle Tank 1 Mixer is pushbutton is pressed and tank 1 manual switch is in "Manual"
      Static PreviousTank1Mixer_PB As Boolean
      If IO.Tank1MixerRequestPb And Not PreviousTank1Mixer_PB Then Tank1MixerRequest = Not Tank1MixerRequest
      PreviousTank1Mixer_PB = IO.Tank1MixerRequestPb
      If Not (Tank1MixerEnable AndAlso IO.Tank1ManualSw) Then Tank1MixerRequest = False


      DrugroomPreview.Run()

      ' Dispensing Stuff
      If LA.KP1.IsOn Then
        DispenseCallOff = LA.KP1.DispenseCalloff
      ElseIf KP.KP1.IsOn Then
        DispenseCallOff = KP.KP1.DispenseCalloff
      Else
        DispenseCallOff = 0
        Parameters.DispenseTestCalloff = 0
      End If

      ' For test
      If Parameters.DispenseTestCalloff > 0 AndAlso Parameters.DispenseTestCalloff < 99 Then DispenseCallOff = Parameters.DispenseTestCalloff
      If Parameters.DispenseTestCalloff = 101 Then
        DispenseCallOff = 0
      End If

      ' Testing - remove once sorted
      'DispenseState = Parameters.DispenseTestState
      'If Parameters.DispenseTestCalloff > 0 Then
      '  DispenseCalloff = Parameters.DispenseTestCalloff
      'End If
      'If Parameters.DispenseTestState > 0 Then
      '  DispenseState = Parameters.DispenseTestState
      'End If
      If DispenseTank <> DispenseTankWas Then
        IO.Siren = True
      End If
      DispenseTankWas = DispenseTank




      '------------------------------------------------------------------------------------------
      ' Digital Outputs (In Order)
      '------------------------------------------------------------------------------------------
      ' CARD #1
      Dim sirenStandard As Boolean = ((Parent.IsAlarmActive AndAlso Parent.IsAlarmUnacknowledged) OrElse Parent.IsSignalUnacknowledged)
      Dim sirenLidActive As Boolean = LidControl.IsActive AndAlso FlashSlow

      IO.Siren = (Parameters.EnableSiren = 1) AndAlso (sirenStandard OrElse sirenLidActive) AndAlso Not (Parent.IsSleeping)
      IO.LampAlarm = (Parent.IsAlarmActive AndAlso FlashSlow) OrElse (Parent.IsAlarmUnacknowledged AndAlso FlashFast) OrElse
                     (Parent.AckState = AckStateValue.QQ) OrElse Not ((IO.KierLidClosed And IO.LockingPinRaised) Or (IO.LidRaisedSw))
      IO.LampSignal = (Parent.IsSignalUnacknowledged AndAlso FlashFast) OrElse
                      (LD.IoLampSignal OrElse PH.IoLampSignal OrElse SA.IoLampSignal OrElse UL.IoLampSignal) ' RP.IsSlow Or RP.IsFast Or AP.IsSlow Or AP.IsFast) And SlowFlash


      IO.LampDelay = KP.KP1.IsOverrun ' A&E vb6 uses delay lamp for slow kitchen (Delay <> DelayValue.NormalRunning) AndAlso FlashSlow 
      IO.LampMachineSafe = MachineSafe AndAlso (IO.Vent OrElse RF.IsFillFromVessel)
      IO.HxDrain = NStop AndAlso TemperatureControl.IoHxDrain
      IO.MainPump = NStop AndAlso Parent.IsProgramRunning AndAlso MachineClosed AndAlso PumpControl.IoMainPump
      IO.MainPumpReset = Remote_MainPumpReset
      IO.AddPumpStart = NHalt And (AC.IoAddPump OrElse AT.IoAddPump OrElse AddControl.IoAddPump)
      IO.HoldDown = (IO.LockingPinRaised OrElse IO.KierLidClosed) AndAlso Not LidControl.IoReleasePin
      IO.SystemBlock = Not (FI.IsActive OrElse RT.IoSystemBlock)
      IO.LevelFlush = (NHSafe OrElse (NHalt AndAlso HD.HDCompleted)) AndAlso (DR.IoLevelFlush OrElse HD.IoLevelFlush OrElse LevelFlush.IsFlushing)

      ' CARD #2
      IO.Vent = (NStop AndAlso SafetyControl.IsDePressurizing OrElse SafetyControl.IsDePressurized) OrElse AirpadBleed
      IO.AirPad = NStop AndAlso MachineClosed AndAlso (SafetyControl.IsPressurized OrElse RF.IoAirpad) AndAlso Not (HD.IoAirpadBleed OrElse AirpadCutoff)
      IO.HeatSelect = NHalt AndAlso (TemperatureControl.IoHeatSelect OrElse ManualSteamTest.IoKierHeatOn)
      IO.CoolSelect = NHalt AndAlso TemperatureControl.IoCoolSelect
      IO.Condensate = NHalt AndAlso (TemperatureControl.IoHeatReturn OrElse ManualSteamTest.IoKierHeatOn)
      IO.WaterDump = NHalt AndAlso TemperatureControl.IoCoolReturn
      IO.Fill = (NHSafe OrElse (NHalt And HD.HDCompleted)) AndAlso MachineClosed AndAlso ((FI.IoFillSelect OrElse HD.IoFill OrElse RC.IoMachineFill OrElse RH.IoMachineFill OrElse RI.IoMachineFill) OrElse ManualSteamTest.IoKierFillOn)

      Dim safeTopWashRequest As Boolean = (IdleDrainActive OrElse DR.IoTopWash OrElse FI.IoTopWash OrElse HD.IoTopWash OrElse LD.IoTopWash OrElse RC.IoTopWash OrElse RH.IoTopWash OrElse RI.IoTopWash OrElse RT.IoTopWash)
      IO.TopWash = (NHSafe AndAlso MachineClosed AndAlso safeTopWashRequest OrElse (NHalt AndAlso (FI.IoTopWash OrElse RT.IoTopWash) AndAlso HD.HDCompleted))
      IO.MachineDrain = (NHSafe OrElse (NHalt AndAlso HD.HDCompleted)) AndAlso (IdleDrainActive OrElse DR.IoDrain OrElse HD.IoMachineDrain OrElse LD.IoDrain OrElse UL.IoDrain)
      IO.MachineDrainHot = NHalt AndAlso (Settings.HighTempDrainEnabled = 1) AndAlso MachineClosed AndAlso (HD.IoHighTempDrain)
      IO.FlowReverse = NStop AndAlso (PumpControl.IoFlowReverse)
      IO.OpenLockingBand = NHSafe AndAlso LidControl.IoLockingBandOpen

      'CARD #3
      IO.OutletInToOut = NStop AndAlso Not IO.FlowReverse
      IO.OutletOutToIn = NStop AndAlso (IO.FlowReverse OrElse DR.IsOn)
      IO.InletInToOut = NStop AndAlso Parent.IsProgramRunning AndAlso Not (IO.FlowReverse OrElse RT.IsTransfer)
      IO.InletOutToIn = NStop AndAlso (IO.FlowReverse OrElse DR.IsOn)


#If 0 Then
      ' GMX .Net
      'FlowReverse = (FR.IsOutToIn Or FC.IsOutToIn) And Not (HD.IsOn Or DR.IsOn Or RI.IsOn Or RH.IsOn Or RC.IsOn Or FI.IsOn Or RT.IsGravityTransfer Or RT.IsTransfer Or RW.IsTransfer)
      ' Pump Hold In-To-Out = NOT (HD.IsOn OrElse DR.IsOn OrElse RI.IsOn OrElse RH.IsOn OrElse RC.IsOn OrElse FI.IsOn OrElse RT.IsTransfer)
      ' IO.FlowOutletInOut = NStop AndAlso Not PumpControl.IoFlowReverse
      ' IO.FlowInletInOut = NStop AndAlso Not (PumpControl.IoFlowReverse OrElse RT.IsTransfer)
#End If


      IO.CloseLockingBand = LidControl.IoLockingBandClose
      IO.AddFillCold = NHalt AndAlso (AC.IoAddFill OrElse AT.IoAddFill OrElse AddControl.IoFillCold)
      IO.AddFillHot = NHalt AndAlso (AddControl.IoFillHot)
      IO.AddDrain = NHalt AndAlso (AC.IoAddDrain OrElse AT.IoAddDrain OrElse AddControl.IoDrain)
      IO.AddTransfer = NHalt AndAlso (AC.IoAddTransfer OrElse AT.IoAddTransfer OrElse AddControl.IoTransfer)
      IO.AddMixing = NHalt AndAlso (AC.IoAddMix OrElse AT.IoAddMix OrElse AddControl.IoAddMix)
      IO.AddRunback = NHSafe AndAlso AddControl.IoFillRunback
      IO.ReserveFillCold = NHalt AndAlso (RD.IoReserveFillCold OrElse RF.IoReserveFillCold OrElse RF.IoReserveFillHot OrElse RT.IoReserveFillCold OrElse
                                      ReserveControl.IoReserveFillCold OrElse ReserveControl.IoReserveFillHot OrElse ManualSteamTest.IoReserveFillCold)


      'CARD #4  
      IO.ReserveFillHot = NHalt And (RF.IoReserveFillHot OrElse ReserveControl.IoReserveFillHot)
      IO.LowerLockingPin = NHSafe AndAlso LidControl.IoReleasePin
      IO.ReserveDrain = NHalt AndAlso (RD.IoDrain OrElse RT.IoReserveDrain OrElse ReserveControl.IoReserveDrain)
      IO.ReserveTransfer = (NHSafe OrElse (NHalt AndAlso HD.HDCompleted)) AndAlso (RT.IoReserveTransfer OrElse ReserveControl.IoReserveTransfer)
      IO.ReserveRunback = NHalt AndAlso TempSafe AndAlso RF.IoReserveRunback

      ' Check Reserve Level and Level Enable parameter before heating Reserve tank - John Smith recommended a hardcoded minimum value of 40.0% to prevent any possible live steam issues resulting
      If (Parameters.ReserveHeatEnableLevel < 400) Then Parameters.ReserveHeatEnableLevel = 400
      If (Parameters.ReserveHeatEnableLevel > 750) Then Parameters.ReserveHeatEnableLevel = 750
      Dim ReserveHeatEnabled As Boolean = (ReserveLevel > Parameters.ReserveHeatEnableLevel) AndAlso Not (Alarms.ReserveTempProbeError OrElse Alarms.ReserveTempTooHigh)
      IO.ReserveHeat = NHalt AndAlso ReserveHeatEnabled AndAlso (RF.IoReserveHeat OrElse ReserveControl.IoReserveHeat) OrElse
                       NStop AndAlso ManualSteamTest.IoReserveHeat

      IO.ReserveMixerStart = NHalt AndAlso Parent.IsProgramRunning AndAlso ReserveMixerOn AndAlso (RF.IoReserveMixer OrElse RT.IoReserveMixer) AndAlso (Not ReserveControl.IoReserveTransfer)
      '                 TODO: Also   Not (RT.IsGravityTransfer Or RT.IsTransfer Or RD.IsRinseToDrain Or RD.IsTransferToDrain  Or ManualReserveTransfer.IsTransferring Or ManualReserveDrain.IsDraining)

      IO.Tank1TransferToDrain = NHalt AndAlso Not (IO.Tank1TransferToAdd OrElse IO.Tank1TransferToReserve)
      IO.Tank1TransferToAdd = NHalt AndAlso (KA.IoTank1ToAdd OrElse KP.KP1.IoToAddTank OrElse LA.KP1.IoToAddTank)
      IO.Tank1TransferToReserve = NHalt AndAlso (KA.IoTank1ToReserve OrElse KP.KP1.IoToReserveTank OrElse LA.KP1.IoToReserveTank)
      IO.OpenLid = NHSafe AndAlso (LidControl.IoRaiseLid OrElse (IO.LidRaisedSw AndAlso (Not IO.CloseLidPb) AndAlso (Not IO.KierLidClosed) AndAlso (Not IO.LockingPinRaised) AndAlso (Not IO.PumpRunning)))
      IO.CloseLid = NHalt AndAlso LidControl.IoLowerLid



      ' KITCHEN PLC - DL06 
      IO.Tank1Mixer = NHalt AndAlso Tank1MixerEnable AndAlso (Tank1MixerRequest OrElse KA.IoTank1Mixer OrElse KP.KP1.IoMixerOn OrElse LA.KP1.IoMixerOn)
      IO.Tank1FillCold = NHalt AndAlso (IO.Tank1FillColdPb OrElse CK.IoTank1Fill OrElse KP.KP1.IoFillCold OrElse KA.IoTank1Rinse OrElse LA.KP1.IoFillCold OrElse ManualSteamTest.IoTank1Heat)

      Dim tank1SteamEnable As Boolean = IO.Tank1Mixer AndAlso Not (Alarms.Tank1TempProbeError OrElse Alarms.Tank1TempTooHigh)
      IO.Tank1Steam = tank1SteamEnable AndAlso (KP.KP1.IoSteam OrElse LA.KP1.IoSteam OrElse IO.Tank1HeatPb)

      IO.Tank1PrepareLamp = NStop AndAlso (Tank1Ready OrElse ((KA.IsWaitReady OrElse KP.KP1.IoTank1Lamp OrElse LA.KP1.IoTank1Lamp) AndAlso FlashFast))
      IO.Tank1Transfer = NHalt AndAlso (Not IO.Tank1ManualSw) AndAlso (CK.IoTank1Transfer OrElse KA.IoTank1Transfer OrElse KP.KP1.IoTransfer OrElse LA.KP1.IoTransfer)
      IO.Tank1FillHot = IO.Tank1FillHotPb AndAlso (Not IO.CityWaterSw)
      IO.CityWaterFill = (NHalt AndAlso IO.CityWaterSw AndAlso (CK.IoTank1Fill OrElse KA.IoTank1Rinse OrElse KP.KP1.IoFillCold OrElse LA.KP1.IoFillCold)) OrElse
                            (IO.CityWaterSw AndAlso (IO.Tank1FillColdPb OrElse IO.Tank1FillHotPb))

      IO.DispenseStateLamp = KP.KP1.IsWaitingToDispense Or KP.KP1.IsDispensing Or (KP.KP1.AddDispenseError And FlashSlow)
      IO.KitchenY10 = False
      IO.KitchenY11 = False
      IO.KitchenY12 = False
      IO.KitchenY13 = False
      IO.KitchenY14 = False
      IO.KitchenY15 = False
      IO.KitchenY16 = False
      IO.KitchenY17 = False


#If 0 Then
      Then Platform Outputs:

      
  IO_BeaconAlarm = (Parent.IsAlarmActive And SlowFlash) Or Not ((IO_LidLoweredSwitch And IO_LockingPinRaised) Or (IO_LidRaised_SW))
  IO_BeaconSignal = ((Not (Parent.ButtonText = "")) Or SignalRequestedWas Or RP.IsSlow _
                      Or RP.IsFast Or AP.IsSlow Or AP.IsFast) And SlowFlash
  IO_BeaconDelay = KP.KP1.IsOverrun
  IO_Siren = Parent.IsAlarmUnacknowledged Or Parent.IsSignalUnacknowledged Or SignalRequested
  IO_SafeLamp = MachineSafe And (IO_Vent Or RF.IsFillWithVessel)
  IO_HxDrain = Not (IO_HeatSelect Or IO_CoolSelect)
  IO_MainPumpStart = NStop And PumpRequest And (PumpLevel Or DR.IsOn Or RT.IsTransfer) And LidLocked And Not PumpOnCountTooHigh
  IO_MainPumpReset = MainPumpResetPB

  IO_AddPumpStart = NHalt And ((AdditionMixOn And AF.AddMixing) Or _
                                AT.IsAddPumping Or AC.IsAddPumping Or AD.IsAddPumping Or _
                                ManualAddTransfer.IsTransferring)
                                      
  IO_HoldDown_Coupling = (IO_LockingPinRaised And IO_LidLoweredSwitch)
  IO_SystemBlock = Not (FI.IsFilling Or RT.IsGravityTransfer Or RT.IsTransfer Or RW.IsTransfer)
  IO_LevelGaugeFlush = FI.IsFillAndFlushing Or DR.IoFlushingLevel Or HD.IoFlushingLevel
  IO_Vent = ((SafetyControl.IsDepressurising Or SafetyControl.IsDepressurised) And Not AirpadOn) _
            Or AirpadBleed Or HD.IoVent Or HD.HDCompleted
  IO_Airpad = NStop And (SafetyControl.IsPressurised Or AirpadOn) And LidLocked And _
              Not (AirpadCutoff Or HD.IsWaitingForPressure Or HD.IsHoldingForTime Or HD.IoVent Or HD.HDCompleted)
  IO_HeatSelect = NHalt And (TemperatureControl.IsHeating Or TemperatureControlContacts.IsHeating Or ManualSteamTest.KierHeatOn)
  IO_CoolSelect = NHalt And (TemperatureControl.IsCooling Or TemperatureControlContacts.IsCooling)
  IO_Condensate = NHalt And (TemperatureControl.IsHeating Or TemperatureControlContacts.IsHeating)
  IO_CoolWaterReturn = Not (TemperatureControl.IsHeating Or TemperatureControlContacts.IsHeating Or ManualSteamTest.KierHeatOn)
  
  'Following delay timer added to allow slower opening topwash valve to fully open before opening fill valve to prevent safety
  '   pressure interlock when filling/rinsing causing a cyclic delay (added 2009/06/12)
  Static FillDelayTimer As New acTimer
  Static FillDesired As Boolean
  FillDesired = (((NHSafe Or (NHalt And HD.HDCompleted)) And LidLocked And (FI.IsFilling Or RI.IsRinsing Or RI.IsFilling Or _
            RH.IsRinsing Or RH.IsFilling Or RC.IsRinsing Or RC.IsFilling)))
  If Not FillDesired Then FillDelayTimer.TimeRemaining = MiniMaxi(Parameters_FillOpenDelayTime, 0, 10)
  IO_Fill = (FillDesired And FillDelayTimer.Finished) Or ManualSteamTest.KierFillOn
  
  IO_TopWash = (NHSafe And (RI.IsRinsing Or RH.IsRinsing Or DR.IsDraining Or RC.IsRinsing Or _
               FI.IsFilling Or RT.IsTransfer Or RT.IsGravityTransfer Or RW.IsTransfer Or UL.IsUnloading Or UL.IsWaitingToUnload)) Or _
               NHalt And (HD.IoDrain Or HD.IsFillingToCool Or HD.HasFilledWaitingToCool Or _
               ((FI.IsFilling Or RT.IsTransfer Or RT.IsGravityTransfer Or RW.IsTransfer) And HD.HDCompleted))
  IO_Drain = NHSafe And (DR.IoDrain Or HD.IoDrain)
  IO_HotDrain = NHalt And HD.IoHotDrain And LidLocked
  IO_FlowReverse = (FR.IsOutToIn Or FC.IsOutToIn) And Not (HD.IsOn Or DR.IsOn Or RI.IsOn Or RH.IsOn Or RC.IsOn Or FI.IsOn Or _
                   RT.IsGravityTransfer Or RT.IsTransfer Or RW.IsTransfer)
  IO_OpenLockingBand = (NHSafe And LidControl.IsOpenLockingBand)
  IO_OutletInToOut = Not IO_FlowReverse
  IO_OutletOutToIn = IO_FlowReverse Or DR.IsOn
  IO_InletInToOut = Parent.IsProgramRunning And Not (IO_FlowReverse Or RT.IsTransfer)
  IO_InletOutToIn = IO_FlowReverse Or DR.IsOn
  IO_CloseLockingBand = LidControl.IsCloseLockingBand
  IO_AddTankColdFill = NHalt And (AD.IsFillForMix Or AF.IsFillWithCold Or AF.IsFillWithMixed Or AT.IsRinse Or _
                                  AT.IsFillForMix Or AC.IsFillForMix Or AC.IsRinse Or _
                       ManualAddDrain.IsRinsing Or ManualAddTransfer.IsRinsing Or ManualAddFill.IsFillingCold)
  IO_AddTankHotFill = NHalt And (AF.IsFillWithHot Or AF.IsFillWithMixed Or ManualAddFill.IsFillingHot)
  IO_AddTankDrain = NHalt And (AT.IsTransferToDrain Or AD.IsTransferToDrain Or AC.IsTransferToDrain _
                    Or ManualAddDrain.IsDraining Or ManualAddTransfer.IsDraining)
  IO_AddTankTransfer = NHalt And (AT.IsTransfer Or AC.IsTransfer Or ManualAddTransfer.IsTransferring)
  
  IO_AddTankMixing = NHalt And ((AdditionMixOn And AF.AddMixing) Or AT.IsDosePaused Or AC.IsDosePaused Or _
                                 AT.IsMixForTime Or AC.IsMixForTime Or AD.IsMixingForTime) _
                           And Not (AT.IsTransfer Or AT.IsMixCloseMix Or AC.IsTransfer Or AC.IsMixCloseMix Or _
                                    AD.IsMixCloseMix Or ManualAddTransfer.IsTransferring)
  IO_AddTankRunback = NHalt And TempSafe And AF.IsFillWithVessel
  IO_ReserveColdFill = (NHalt And (RF.IsFillWithCold Or RF.IsFillWithCold Or ManualReserveFill.IsFillingCold Or _
                       RW.IsFilling Or RW.IsRinse Or RW.IsRinseToDrain Or RT.IsRinse Or RT.IsRinseToDrain Or _
                       RD.IsRinseToDrain Or ManualReserveDrain.IsRinsing Or ManualReserveTransfer.IsRinsing)) Or _
                       ManualSteamTest.RTFillOn
  IO_ReserveHotFill = NHalt And (RF.IsFillWithHot Or RF.IsFillWithHot Or ManualReserveFill.IsFillingHot)
  IO_LowerLockingPin = (NHSafe And LidControl.IsReleasePin)
  IO_ReserveDrain = NHalt And (RT.IsRinseToDrain Or RT.IsTransferToDrain Or RD.IsTransferToDrain _
                    Or ManualReserveDrain.IsDraining Or ManualReserveTransfer.IsDraining Or RW.IsRinseToDrain _
                    Or RW.IsTransferToDrain)
  IO_ReserveTransfer = (NHSafe Or (NHalt And HD.HDCompleted)) And (RT.IsGravityTransfer Or RT.IsTransfer Or ManualReserveTransfer.IsTransferring _
                      Or RW.IsTransfer)
  IO_ReserveBackflow = NHalt And TempSafe And RF.IsFillWithVessel
  IO_ReserveHeating = (NHalt And IO_ReserveMixer And (RF.IsHeating Or ManualReserveHeat.IsHeating Or RW.IsHeating) And _
                     (Not Alarms_ReserveTempProbeError) And (Not Alarms_ReserveTempTooHigh)) Or ManualSteamTest.RTHeatOn
  IO_ReserveMixer = NHalt And Parent.IsProgramRunning And ReserveMixerOn And (RF.ReserveMixing Or RW.IsHeating) And _
                              Not (RT.IsGravityTransfer Or RT.IsTransfer Or RD.IsRinseToDrain Or RD.IsTransferToDrain _
                              Or ManualReserveTransfer.IsTransferring Or ManualReserveDrain.IsDraining)
  IO_Tank1ToDrain = (NHalt And (Parameters_DDSEnabled <> 1) And Not (KA.IsTransferToAddition Or KA.IsTransferToReserve Or _
                                                                     KP.KP1.IsTransferToReserve Or KP.KP1.IsTransferToAddition Or _
                                                                     LA.KP1.IsTransferToReserve Or LA.KP1.IsTransferToAddition)) Or _
                    (NHalt And (Parameters_DDSEnabled = 1) And Not (IO_DyeDispenseReserve Or IO_DyeDispenseAdd))
  IO_Tank1ToAdd = (NHalt And (Parameters_DDSEnabled <> 1) And (KA.IsTransferToAddition Or KP.KP1.IsTransferToAddition Or LA.KP1.IsTransferToAddition)) Or _
                  (NHalt And (Parameters_DDSEnabled = 1) And IO_DyeDispenseAdd)
  IO_Tank1ToReserve = (NHalt And (Parameters_DDSEnabled <> 1) And (KA.IsTransferToReserve Or KP.KP1.IsTransferToReserve Or LA.KP1.IsTransferToReserve)) Or _
                  (NHalt And (Parameters_DDSEnabled = 1) And IO_DyeDispenseReserve)
  IO_OpenLid = NHSafe And (LidControl.IsRaiseLid Or (IO_LidRaised_SW And (Not IO_CloseLid_PB) And _
              (Not IO_LidLoweredSwitch) And (Not IO_LockingPinRaised) And (Not IO_MainPumpRunning)))
  IO_CloseLid = NHalt And LidControl.IsLowerLid
  IO_PlcWatchDog1 = True
  
  ' Drugroom Stuff
  IO_Tank1Mixer = (NHalt And (Not KA.IsTransfer) And DrugroomMixerOn And (KP.KP1.IsMixerOn Or LA.KP1.IsMixerOn)) Or _
                   (IO_Tank1Mixer_PB And DrugroomMixerOn) Or (Tank1MixerRequest)
  IO_Tank1ColdFill = ((NHalt And (Not IO_CityWater_SW) And _
                     (KP.KP1.IsFill Or LA.KP1.IsFill Or CK.IsFilling Or _
                     KA.IsRinse Or KA.IsRinseToDrain Or CK.IsRinseToDrain Or _
                     KP.KP1.IsRinse Or KP.KP1.IsRinseToDrain Or LA.KP1.IsRinse Or LA.KP1.IsRinseToDrain)) Or _
                     (IO_Tank1FillCold_PB And (Not IO_CityWater_SW))) Or ManualSteamTest.DRFillOn
  IO_Tank1Steam = ((NHalt And IO_Tank1Mixer And (KP.KP1.IsHeating Or LA.KP1.IsHeating) And (Not Alarms_Tank1TempProbeError) And (Not Alarms_Tank1TempTooHigh)) Or _
                  IO_Tank1Heating_PB And IO_Tank1Mixer And ((Not Alarms_Tank1TempProbeError) Or (Not Alarms_Tank1TempTooHigh))) Or _
                  ManualSteamTest.DRHeatOn
  IO_Tank1PrepareLamp = (KP.KP1.IsSlow) Or (KP.KP1.IsFast And FastFlash) Or _
                        (LA.KP1.IsSlow) Or (LA.KP1.IsFast And FastFlash) Or _
                        ((KP.KP1.IsHeatTankOverrun Or LA.KP1.IsHeatTankOverrun) And FastFlash)
  IO_Tank1Transfer = NHalt And (Not IO_Tank1Manual_SW) And (KA.IsTransfer Or CK.IsTransferToDrain Or KP.KP1.IsTransfer Or LA.KP1.IsTransfer)
  IO_Tank1HotFill = IO_Tank1FillHot_PB And (Not IO_CityWater_SW)
  IO_CityWaterFill = (NHalt And IO_CityWater_SW And (KA.IsRinse Or KA.IsRinseToDrain Or CK.IsRinseToDrain Or _
                      KP.KP1.IsFill Or LA.KP1.IsFill Or CK.IsFilling Or KP.KP1.IsRinse Or KP.KP1.IsRinseToDrain Or _
                      LA.KP1.IsRinse Or LA.KP1.IsRinseToDrain)) _
                      Or ((IO_Tank1FillHot_PB Or IO_Tank1FillCold_PB) And IO_CityWater_SW)
  IO_DispenseStateLamp = KP.KP1.IsWaitingToDispense Or KP.KP1.IsDispensing Or (KP.KP1.AddDispenseError And SlowFlash)
  IO_PlcWatchDog2 = True
  
'analog outputs===========================================================

'Analog Heat/Cool Control
    IO_HeatCoolOutput = 0
    If NHalt And IO_MainPumpRunning Then
        If TemperatureControl.IsHeating Or TemperatureControl.IsCooling Then IO_HeatCoolOutput = TemperatureControl.Output
        If TemperatureControlContacts.IsHeating Or TemperatureControlContacts.IsCooling Then IO_HeatCoolOutput = TemperatureControlContacts.Output
    ElseIf ManualSteamTest.KierHeatOn Then
        IO_HeatCoolOutput = 1000
    End If
    If IO_HeatCoolOutput < 0 Then IO_HeatCoolOutput = 0
    If IO_HeatCoolOutput > 1000 Then IO_HeatCoolOutput = 1000


'Analog Fill Control
' copied


'Pump speed
  If Not IO_MainPumpRunning Then PumpAccelerationTimer = Parameters_PumpAccelerationTime
  'Set timer if we've just changed direction
  Static FlowReverseWas As Boolean
  If (DP.IsOn Or FL.IsOn) And IO_MainPumpRunning And (Not FI.IsActive) Then
    'When we change direction
    If IO_FlowReverse Then
      If Not FlowReverseWas Then FlowReverseTimer = Parameters_FlowReverseTime
    Else
      If FlowReverseWas Then FlowReverseTimer = Parameters_FlowReverseTime
    End If
    'Now check whether we should adjust the pump speed
    If FlowReverseTimer.Finished And PumpAccelerationTimer.Finished And (Not (FR.IsSwitching Or FC.IsSwitching)) Then
        If PumpSpeedAdjustTimer.Finished Then
          PumpSpeedAdjustTimer = Parameters_PumpSpeedAdjustTime
          If IO_FlowReverse Then
            If ((FL.OutToInFlowRate > 0) And (FL.IsOn)) Then
              If (PackageDifferentialPressure > Parameters_MaxDifferentialPressure) And (Parameters_MaxDifferentialPressure > 0) Then
                DPError = Parameters_MaxDifferentialPressure - PackageDifferentialPressure
                DPAdjust = MulDiv(DPError, Parameters_DPAdjustFactor, 100)
                If DPAdjust < -Parameters_DPAdjustMaximum Then DPAdjust = -Parameters_DPAdjustMaximum
                If DPAdjust > Parameters_DPAdjustMaximum Then DPAdjust = Parameters_DPAdjustMaximum
                OutToInSpeed = OutToInSpeed + DPAdjust
              Else
                FLError = FL.OutToInFlowRate - FlowRate
                FLAdjust = MulDiv(FLError, Parameters_FLAdjustFactor, 100)
                If FLAdjust < -Parameters_FLAdjustMaximum Then FLAdjust = -Parameters_FLAdjustMaximum
                If FLAdjust > Parameters_FLAdjustMaximum Then FLAdjust = Parameters_FLAdjustMaximum
                OutToInSpeed = OutToInSpeed + FLAdjust
              End If
            ElseIf (DP.OutToInPressure > 0) Then
              DPError = DP.OutToInPressure - PackageDifferentialPressure
              DPAdjust = MulDiv(DPError, Parameters_DPAdjustFactor, 100)
              If DPAdjust < -Parameters_DPAdjustMaximum Then DPAdjust = -Parameters_DPAdjustMaximum
              If DPAdjust > Parameters_DPAdjustMaximum Then DPAdjust = Parameters_DPAdjustMaximum
              OutToInSpeed = OutToInSpeed + DPAdjust
            End If
          Else
            If ((FL.InToOutFlowRate > 0) And (FL.IsOn)) Then
              If (PackageDifferentialPressure > Parameters_MaxDifferentialPressure) And (Parameters_MaxDifferentialPressure > 0) Then
                DPError = Parameters_MaxDifferentialPressure - PackageDifferentialPressure
                DPAdjust = MulDiv(DPError, Parameters_DPAdjustFactor, 100)
                If DPAdjust < -Parameters_DPAdjustMaximum Then DPAdjust = -Parameters_DPAdjustMaximum
                If DPAdjust > Parameters_DPAdjustMaximum Then DPAdjust = Parameters_DPAdjustMaximum
                InToOutSpeed = InToOutSpeed + DPAdjust
              Else
                FLError = FL.InToOutFlowRate - FlowRate
                FLAdjust = MulDiv(FLError, Parameters_FLAdjustFactor, 100)
                If FLAdjust < -Parameters_FLAdjustMaximum Then FLAdjust = -Parameters_FLAdjustMaximum
                If FLAdjust > Parameters_FLAdjustMaximum Then FLAdjust = Parameters_FLAdjustMaximum
                InToOutSpeed = InToOutSpeed + FLAdjust
              End If
            ElseIf (DP.InToOutPressure > 0) Then
              DPError = DP.InToOutPressure - PackageDifferentialPressure
              DPAdjust = MulDiv(DPError, Parameters_DPAdjustFactor, 100)
              If DPAdjust < -Parameters_DPAdjustMaximum Then DPAdjust = -Parameters_DPAdjustMaximum
              If DPAdjust > Parameters_DPAdjustMaximum Then DPAdjust = Parameters_DPAdjustMaximum
              InToOutSpeed = InToOutSpeed + DPAdjust
           End If
          End If
        End If
    End If
  ElseIf FP.IsOn Then
    InToOutSpeed = FP.FPPercentInToOut
    OutToInSpeed = FP.FPPercentOutToIn
    PumpSpeedAdjustTimer = Parameters_PumpSpeedAdjustTime
   Else
    InToOutSpeed = Parameters_PumpSpeedStart
    OutToInSpeed = Parameters_PumpSpeedStart
    PumpSpeedAdjustTimer = Parameters_PumpSpeedAdjustTime
  End If
  FlowReverseWas = IO_FlowReverse 'remember which direction we were
  If InToOutSpeed < 0 Then InToOutSpeed = 0
  If InToOutSpeed > 1000 Then InToOutSpeed = 1000
  If OutToInSpeed < 0 Then OutToInSpeed = 0
  If OutToInSpeed > 1000 Then OutToInSpeed = 1000
  
  If IO_FlowReverse Then
    IO_PumpSpeedOutput = OutToInSpeed
  Else
    IO_PumpSpeedOutput = InToOutSpeed
  End If
  
  If FR.IsSwitching Or FR.IsDeceling Or FC.IsSwitching Or FC.IsDeceling Then
     IO_PumpSpeedOutput = Parameters_PumpReversalSpeed
     PumpAccelerationTimer = Parameters_PumpAccelerationTime
  End If
  
  If FI.IsActive Then
     IO_PumpSpeedOutput = Parameters_FillSettlePumpSpeed
     PumpAccelerationTimer = Parameters_PumpAccelerationTime
  End If
  
  If RI.IsActive Or RH.IsActive Or RC.IsActive Or RW.IsTransfer Or RT.IsTransfer Then
     IO_PumpSpeedOutput = Parameters_RinsePumpSpeed
     PumpAccelerationTimer = Parameters_PumpAccelerationTime
  End If
  
  If DR.IsActive Then
     IO_PumpSpeedOutput = Parameters_DrainPumpSpeed
     PumpAccelerationTimer = Parameters_PumpAccelerationTime
  End If
  

#End If


      '------------------------------------------------------------------------------------------
      ' ANALOG OUTPUTS IN ORDER
      '------------------------------------------------------------------------------------------

      ' ANOUT1 - HEAT/COOL OUTPUT
      IO.HeatCoolOutput = 0
      If NHalt AndAlso IO.PumpRunning Then
        If TemperatureControl.IsHeating AndAlso IO.PumpRunning Then IO.HeatCoolOutput = CType(MinMax(TemperatureControl.IoOutput, 0, 1000), Short)
        If TemperatureControl.IsCooling Then IO.HeatCoolOutput = CType(MinMax(TemperatureControl.IoOutput, 0, 1000), Short)
      ElseIf ManualSteamTest.IoKierHeatOn Then
        IO.HeatCoolOutput = 1000
      End If
      ' Limit Values
      If IO.HeatCoolOutput < 0 Then IO.HeatCoolOutput = 0
      If IO.HeatCoolOutput > 1000 Then IO.HeatCoolOutput = 1000


      ' BLEND FILL OUTPUT
      IO.BlendFillOutput = 0
      If NHalt Then
        If FI.IoFillSelect Then IO.BlendFillOutput = CShort(FI.IOBlendOutput)
        If RC.IoMachineFill Then IO.BlendFillOutput = CShort(RC.IOBlendOutput)
        If RH.IoMachineFill Then IO.BlendFillOutput = CShort(RH.IOBlendOutput)
        If RI.IoMachineFill Then IO.BlendFillOutput = CShort(RI.IOBlendOutput)
        If HD.IsFillingToCool Then IO.BlendFillOutput = CShort(Parameters.HDFillPosition)
        If IO.CityWaterSw Then IO.BlendFillOutput = 0
      End If
      ' Limits
      If IO.BlendFillOutput < 0 Then IO.BlendFillOutput = 0
      If IO.BlendFillOutput > 1000 Then IO.BlendFillOutput = 1000




      '\\\  SPARE OUTPUTS    \\\\
      IO.AnalogOutput3 = 0

      ' ANOUT4 - PUMP SPEED OUTPUT
      IO.PumpSpeedOutput = PumpControl.IoMainPumpSpeed



      ' Keep at end of run section
      If Parent.IsPaused Then WasPausedTimer.TimeRemaining = 15
      WasPaused = Not (WasPausedTimer.Finished)


      '------------------------------------------------------------------------------------------
      ' Run alarms, delays, utilities, parameters, sleep, hibernate
      '------------------------------------------------------------------------------------------
      Alarms.Run()
      RunStatusColor()
      RunDelays()
      RunSleep()
      RunHibernate()



      Parameters.MaybeSetDefaults()

      CalculateUtilities()

      '------------------------------------------------------------------------------------------
      ' Run remote push buttons - this can be in only one place
      '------------------------------------------------------------------------------------------
      Parent.PressButtons(AdvancePb OrElse PressRunButton, IO.HaltPb OrElse PressPauseButton, False, False, False)
      PressRunButton = False : PressPauseButton = False

      '------------------------------------------------------------------------------------------
      ' Clear expert mode if it has been set for too long 
      '------------------------------------------------------------------------------------------   
      'If Parent.Mode = Mode.Override OrElse Parent.Mode = Mode.Test OrElse Not Parent.ExpertMode Then ExpertModeTimer.Minutes = 60
      'If ExpertModeTimer.Finished AndAlso Parent.ExpertMode Then
      '  Parent.SetExpertMode(False)
      'End If


      ' Set First Scan Done Flag
      FirstScanDone = True

    Catch ex As Exception
      Parent.LogException(ex)
    End Try
  End Sub

  Private Sub RunSleep()
    ' Pause program if we go into sleep mode
    Static IsSleepingWas As Boolean
    If Parent.IsSleeping AndAlso Not IsSleepingWas Then
      If Parent.IsProgramRunning Then PressPauseButton = True
    End If
    IsSleepingWas = Parent.IsSleeping
  End Sub

  Private Sub RunHibernate()
    Dim hibernateDelay = MinMax(Parameters.HibernateDelay, 30, 120)

    ' If no optomux comms for 60 seconds call hibernate
    If Not Alarms.PlcComs1 Then SystemHibernateTimer.Seconds = hibernateDelay
    If Not IO.UpsBatteryMode Then SystemHibernateTimer.Seconds = hibernateDelay + 1
    '    If (Parameters.HibernateEnable = 0) Then SystemHibernateTimer.Seconds = hibernateDelay + 1

    ' Check for system hibernate
    If SystemHibernateTimer.Finished AndAlso (Not SystemHibernate) AndAlso (Parent.Mode <> Mode.Debug) Then
      If Parameters.HibernateEnable = 1 Then
        SystemHibernate = True
        Application.SetSuspendState(PowerState.Hibernate, False, False)
      End If
    End If

    ' Reset system hibernate flag
    If (SystemHibernate AndAlso (Not SystemHibernateTimer.Finished)) OrElse (Parameters.HibernateEnable = 0) Then
      SystemHibernate = False
    End If
  End Sub



  Sub RunStatusColor()
    ' Set default colors
    Dim DefaultColor As Integer = Utilities.Color.ConvertColorToStatusColor(Drawing.Color.FromArgb(255, 255, 255))           ' White
    Dim AlarmColor As Integer = Utilities.Color.ConvertColorToStatusColor(Drawing.Color.FromArgb(255, 128, 128))             ' Red
    Dim AlarmSilencedColor As Integer = Utilities.Color.ConvertColorToStatusColor(Drawing.Color.FromArgb(255, 255, 128))     ' Yellow
    Dim SignalColor As Integer = Utilities.Color.ConvertColorToStatusColor(Drawing.Color.FromArgb(128, 192, 255))            ' Blue
    Dim OperatorColor As Integer = Utilities.Color.ConvertColorToStatusColor(Drawing.Color.FromArgb(192, 255, 255))          ' Turquoise
    Dim DelayColor As Integer = Utilities.Color.ConvertColorToStatusColor(Drawing.Color.FromArgb(255, 192, 128))             ' Orange
    'Dim DelayColor As Integer = Utilities.Color.ConvertColorToStatusColor(Drawing.Color.FromArgb(255, 192, 192))             ' Salmon

    ' Use parameter values if set
    'If Parameters.StatusColorAlarm > 0 Then AlarmColor = Parameters.StatusColorAlarm
    'If Parameters.StatusColorAlarmSilenced > 0 Then AlarmSilencedColor = Parameters.StatusColorAlarmSilenced
    'If Parameters.StatusColorSignal > 0 Then SignalColor = Parameters.StatusColorSignal
    'If Parameters.StatusColorFloor > 0 Then OperatorColor = Parameters.StatusColorFloor

    ' Start with the default
    StatusColor = DefaultColor


    ' Use default color if no program running
    If Not Parent.IsProgramRunning Then Exit Sub

    'If Parent.IsSignalUnacknowledged Then StatusColor = SignalColor
    If Parent.IsAlarmActive Then StatusColor = AlarmSilencedColor
    If Parent.IsAlarmUnacknowledged Then StatusColor = AlarmColor

    ' Get current command
    Dim Command As String = CurrentCommand
    If String.IsNullOrEmpty(Command) Then Exit Sub

    'Change status color if one of these commands is active
    Select Case Command
      Case "LD" : StatusColor = OperatorColor
      Case "PH" : StatusColor = OperatorColor
      Case "SA" : StatusColor = OperatorColor
      Case "UL" : StatusColor = OperatorColor
    End Select

  End Sub

  Private Sub SetStatusColor(color As Color)
    StatusColor = color.ToArgb And 16777215
  End Sub

  Public ReadOnly Property Temperature() As String
    Get
      Temperature = (Temp \ 10) & "F"
    End Get
  End Property




#Region " CONTROL METHODS "

  Public Function ReadInputs(ByVal dinp() As Boolean, ByVal aninp() As Short, ByVal temp() As Short) As Boolean Implements ACControlCode.ReadInputs
    Return IO.ReadInputs(Parent, dinp, aninp, temp, Me)
  End Function

  Public Sub WriteOutputs(ByVal dout() As Boolean, ByVal anout() As Short) Implements ACControlCode.WriteOutputs
    IO.WriteOutputs(dout, anout, Me)
  End Sub

  Public Sub StartUp() Implements ACControlCode.StartUp
    ' Start the programmed stopped timer so we can display idle for xx mins...
    ProgramStoppedTimer.Start()

    PowerOnTimer.Seconds = 10
  End Sub

  Public Sub ShutDown() Implements ACControlCode.ShutDown
    SystemShuttingDown = True

    ' Update/Save settings
    Settings.Save()
  End Sub

  Public Sub ProgramStart() Implements ACControlCode.ProgramStart
    'Reset program timers
    ProgramStoppedTime = ProgramStoppedTimer.Seconds
    ProgramStoppedTimer.Stop()
    ProgramRunTimer.Start()
    StartTime = Date.UtcNow.ToLocalTime.ToShortTimeString

    ' Kitchen Preview Stuff
    LoadRecipeDetailsActiveJob()

    DrugroomPreview.LoadDyelot()
    DrugroomPreview.Dyelot = Dyelot
    DrugroomPreview.ReDye = ReDye


    ' Calculate StandardTime based on every step that's about to start running
    Dim totalTime As TimeSpan
    For i As Integer = 0 To Parent.ProgramStepCount - 1
      Dim programStep = Parent.ProgramStep(i)
      Dim stepNumber As Integer = programStep.StepNumber
      Dim stepCode As String = programStep.Command

      totalTime += programStep.TotalTime
    Next i
    ' Update StandardTime, rounded to parts of a day (6hr = 1/4day)
    StandardTimeCalculated = Math.Round(totalTime.TotalDays, 6)

    ' Update Standard Time if value set
    If StandardTimeCalculated <> 0 Then UpdateStandardTime()

    ' Clear mimic variabes
    Remote_MainPumpReset = False

    ' Clear variables
    CycleTime = 0
    SteamUsed = 0
    PowerKWS = 0
    PowerKWH = 0
    AirpadBleedCount = 0
    GetRecordedLevel = False
    MachinePressureMax = 0

    ' Clear Delay
    Delay = DelayValue.NormalRunning

    ' Cancel control objects
    TemperatureControl.Cancel()
    FR.Clear()
    FC.Clear()
    LevelFlush.Reset()
    PumpControl.CountInToOut = 0
    PumpControl.CountOutToIn = 0

    ' Way to simulate operator pressing main tab at operator interface
    Parent.PressButton(ButtonPosition.Operator, 1)

  End Sub

  Public Sub UpdateStandardTime()
    Dim sql As String = "UPDATE Dyelots SET StandardTime=" & StandardTimeCalculated & " " & "WHERE Dyelot='" & Dyelot & "' AND ReDye=" & ReDye.ToString("#0")
    Try
      Parent.DbExecute(sql)

    Catch ex As Exception
      Utilities.Log.LogError(ex, sql)
    End Try
  End Sub

  Public Sub ProgramStop() Implements ACControlCode.ProgramStop
    ' Reset program timers
    ProgramRunTime = ProgramRunTimer.Seconds
    ProgramRunTimer.Stop()
    ProgramStoppedTimer.Start()

    ' Clear Utility Usage
    LastProgramCycleTime = CycleTime
    '   WaterUsedLastCycle = WaterUsed
    CycleTime = 0
    CycleTimeDisplay = ""
    SteamUsed = 0
    PowerKWS = 0
    PowerKWH = 0
    SystemVolume = 0
    GetRecordedLevel = False
    AirpadBleedCount = 0

    ' Clear batch data
    BatchWeight = 0
    NumberOfContacts = 0
    TotalNumberOfContacts = 0
    PackageHeight = 0
    PackageType = 0

    ' TODO?
    'AddReadyFromPB = f 
    'ReserveReadyFromPB = f 

    Alarms.LidLockedTimer.Seconds = 5
    Alarms.LidNotLocked = False
    Alarms.LosingLevelInTheVessel = False

    ' Signals
    Parent.Signal = ""
    If Not LA.KP1.IsOn Then Tank1Ready = False

    ' Cancel control objects
    TemperatureControl.Cancel()
    ' TemperatureControlContacts.Cancel() ' TODO
    PumpControl.StopMainPump()
    AirpadOn = False

    ' Drugroom Preview Stuff
    Dyelot = ""
    ReDye = 0
    RecipeCode = ""
    RecipeName = ""
    RecipeProducts = ""
    StyleCode = ""

    ' Clear variables
    Delay = DelayValue.NormalRunning

    'Explicitly call cancel on all commands just to be sure
    For Each command As ACCommand In Commands : command.Cancel() : Next
  End Sub



#End Region

#Region " PROGRAM STATE CHANGES "

  Private Sub CheckProgramStateChanges()
    'Program running state changes
    Static ProgramWasRunning As Boolean

    If Parent.IsProgramRunning Then
      ' Update Tank1 Status 
      CheckTank1Status()

      ' Update CycleTime Variables
      CycleTime = ProgramRunTimer.Seconds
      CycleTimeDisplay = TimerString(CycleTime)

      ' Calculate end time stuff.
      If GetEndTimeTimer.Finished OrElse (StepStandardTime <> StepStandardTimeWas) Then
        StepStandardTimeWas = StepStandardTime
        EndTimeMins = GetEndTime(Me)
        'see if current step is overruning or not
        If TimeInStepValue <= StepStandardTime Then
          EndTimeMins += StepStandardTime - TimeInStepValue
        Else
          EndTimeMins += 1
        End If
        EndTime = Date.Now.AddMinutes(EndTimeMins).ToString

        ' Reset GetEndTime timer
        GetEndTimeTimer.Seconds = 60
      End If


      ' Dispenser & Drugroom Preview Values

    Else
      ' Program Not Running
      CycleTime = 0
      CycleTimeDisplay = ""
      'Reset End-Time Stuff
      EndTime = ""
      StartTime = ""
      EndTimeMins = 0
      'Reset Time-In-Step stuff
      TimeInStepValue = 0
      StepStandardTime = 0
    End If
    ProgramWasRunning = Parent.IsProgramRunning

    'time-in-step routine
    Static DisplayTIS As Boolean
    If TimeInStepValue > StepStandardTime Then
      If TwoSecondTimer.Finished Then
        DisplayTIS = Not DisplayTIS
        TwoSecondTimer.Seconds = 2
      End If
    Else
      DisplayTIS = True
    End If
    If DisplayTIS Then
      TimeInStep = CStr(TimeInStepValue)
    Else
      TimeInStep = Translate("Overrun")
    End If

    Dim tis As String = Parent.TimeInStep
    Dim f As Integer = tis.IndexOf("/")
    If f <> -1 Then
      TimeInStepValue = CType(tis.Substring(0, f), Integer)
      StepStandardTime = CType(tis.Substring(f + 1), Integer)
    Else
      TimeInStepValue = 0 : StepStandardTime = 0
    End If

#If 0 Then ' TODO vb6 and remove
Private Function TISValue(Index As Long) As Long
  On Error GoTo Err_TISValue
    Dim Values() As String
    Values = Split(Parent.TimeInStep, "/")
    TISValue = Values(Index)
  Exit Function
Err_TISValue:
  TISValue = 0
End Function
#End If


    ' vb6 TimeToGo 
    TimeToGo = 0
    If AC.IsOn Then TimeToGo = AC.Timer.Seconds
    If AT.IsOn Then TimeToGo = AT.Timer.Seconds
    If DR.IsOn Then TimeToGo = DR.Timer.Seconds
    If FI.IsOn Then TimeToGo = FI.Timer.Seconds ' TODO Check (FITimer)
    If RI.IsOn Then TimeToGo = RI.Timer.Seconds
    If RH.IsOn Then TimeToGo = RH.Timer.Seconds
    If RT.IsOn Then TimeToGo = RT.Timer.Seconds
    If TM.IsOn Then TimeToGo = TM.Timer.Seconds
    If HD.IsOn Then TimeToGo = HD.Timer.Seconds
    If RW.IsActive Then TimeToGo = RW.Timer.Seconds

  End Sub

  Private Function GetEndTime(ByVal controlCode As ControlCode) As Integer
    Dim returnValue As Integer = 0
    Try
      With controlCode

        ' These separators are used by programs and prefixed steps
        Dim LineFeed As String = Convert.ToChar(13) & Convert.ToChar(10)
        Dim separator1 As String = Convert.ToChar(255)
        Dim separator2 As String = ","

        'Get current program and step number
        Dim programNumber As String = .Parent.ProgramNumberAsString
        Dim stepNumber As Integer : stepNumber = .Parent.StepNumber

        ' Get all program steps for this program
        Dim PrefixedSteps As String : PrefixedSteps = .Parent.PrefixedSteps
        Dim ProgramSteps() As String : ProgramSteps = PrefixedSteps.Split(LineFeed.ToCharArray, StringSplitOptions.RemoveEmptyEntries)

        ' Variables for Command Step parameters & Notes
        Dim stepString As String
        Dim stepParameters() As String
        Dim stepNotes() As String

        ' Loop through each step starting from the beginning of the program
        Dim startChecking As Boolean
        For i = ProgramSteps.GetLowerBound(0) To ProgramSteps.GetUpperBound(0)
          ' Ignore step 0
          If i > 0 Then
            ' Current Step
            Dim [Step] = ProgramSteps(i).Split(separator1.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            If [Step].GetUpperBound(0) >= 1 Then
              If startChecking Then returnValue += CInt([Step](3))
              If ([Step](0)) = programNumber AndAlso (CDbl([Step](1)) = stepNumber - 1) Then startChecking = True

              ' Grab program step string with notes & parameters, if any
              stepString = [Step](5)
              stepNotes = stepString.Split("'".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
              stepParameters = stepNotes(0).Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            End If
          End If
          ' Reached the upper bound of the program steps array
          If i = ProgramSteps.GetUpperBound(0) Then Exit For
        Next i

        ' Made it here ok
        Return returnValue

      End With
    Catch ex As Exception
      Utilities.Log.LogError(ex)
    End Try

    'Return Default Value 
    Return returnValue
  End Function


#End Region


  ' TODO - Organize following

  ReadOnly Property CurrentStep As ProgramStep
    Get
      Try
        If Not Parent.IsProgramRunning Then Return Nothing
        Dim stepNumber = Parent.CurrentStep
        Dim maxStep = Parent.ProgramStepCount
        If stepNumber > 0 AndAlso stepNumber < maxStep Then Return Parent.ProgramStep(stepNumber)
      Catch ex As Exception
      End Try
      Return Nothing
    End Get
  End Property

  ReadOnly Property CurrentCommand As String
    Get
      If Not Parent.IsProgramRunning Then Return Nothing
      If CurrentStep Is Nothing Then Return Nothing
      Return CurrentStep.Command
    End Get
  End Property

  ReadOnly Property IsSignalling As Boolean
    Get
      Return Parent.IsSignalUnacknowledged OrElse Parent.AckState = AckStateValue.UnackMessage OrElse Parent.AckState = AckStateValue.AckMessage
    End Get
  End Property

End Class
