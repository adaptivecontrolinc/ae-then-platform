Imports Utilities.Translations

Partial Class ControlCode

  '==============================================================================
  'Program and step time displays
  '==============================================================================
  Private CycleTime As Integer
  Public CycleTimeDisplay As String
  Private LastProgramCycleTime As Integer

  'Time control system has been idle
  Public ProgramStoppedTimer As New TimerUp
  Public ProgramStoppedTime As Integer

  'Time program has been running
  Public ProgramRunTimer As New TimerUp
  Public ProgramRunTime As Integer

  'Current time in step and step overrun
  Private TwoSecondTimer As New Timer
  Public TimeInStepValue As Integer, StepStandardTime As Integer, TimeInStep As String
  Public StartTime As String
  Public EndTime As String
  Public EndTimeMins As Integer
  Public GetEndTimeTimer As New Timer
  Public StepStandardTimeWas As Integer
  Public StandardTimeCalculated As Double
  Public TimeToGo As Integer

  ' Flashers
  Private Property _fastFlasher As New Flasher(400)
  Private Property _slowFlasher As New Flasher(800)
  Friend FlashFast As Boolean, FlashSlow As Boolean

  'Variables for troubleshooting histories
  <GraphTrace(0, 1, 0, 1, "Red")> Public WasPaused As Boolean
  Private WasPausedTimer As New Timer
  Public EStop As Boolean

  Public OkToStart As Boolean = True 'MUST BE SET TO TRUE TO RELEASE (BC 3.2.176+)

  '==============================================================================
  ' Graph variables
  ' First graph label line defines the level scale
  ' Second graph label line defines the temperature scale
  ' %d = not divide value by ten for display
  ' %f = divide value by ten for display
  '==============================================================================


  <GraphTrace(0, 3000, 0, 10000, "DarkBlue", "%tF"), GraphLabel("50", 500), GraphLabel("100", 1000),
    GraphLabel("150", 1500), GraphLabel("200", 2000), GraphLabel("250", 2500), GraphLabel("300", 3000)>
  Public Temp As Integer

  <GraphTrace(1, 3000, 0, 10000, "Red", "%tF")> Public Setpoint As Integer
  Public TempFinalValue As String


  'Attribute PackageDifferentialPressure.VB_VarDescription = "MaximumValue=1000\r\nMaximumY=3333\r\nFormat=%t psi\r\nColor=9\r\nOrder=9"
  <GraphTrace(0, 1000, 0, 5000, "Blue", "%t%")> Public MachineLevel As Integer

  'Attribute VesselPressure.VB_VarDescription = "MaximumValue=1000\r\nMaximumY=3333\r\nFormat=%t psi\r\nColor=8\r\nOrder=8"
  <GraphTrace(1, 1000, 0, 3333, "DarkRed", "%tpsi")> Public MachinePressure As Integer  ' VesselPressure
  Public MachinePressureDisplay As String
  Public MachinePressurePsi As Integer
  Public MachinePressureMax As Integer


  <GraphTrace(1, 1000, 0, 3333, "Purple", "%tpsi")> Public PackageDifferentialPressure As Integer

  <GraphTrace(1, 1000, 0, 3333, "DarkBlue", "%t%")> Public AddLevel As Integer
  <GraphTrace(1, 1000, 0, 3333, "DarkGreen", "%t%")> Public ReserveLevel As Integer



  '==============================================================================
  'Various Variables
  '==============================================================================
  ' Property to handle both Advance Pushbuttons
  Public ReadOnly Property AdvancePb As Boolean
    Get
      Return (IO.AdvancePb)
    End Get
  End Property

  'Safety Variables
  Public MachineSafe As Boolean
  Public TempValid As Boolean, TempSafe As Boolean
  Public PressSafe As Boolean
  Public MachineClosed As Boolean
  Public MachineSafeToOpen As Boolean
  Public VesselHeatSafe As Boolean

  'BatchWeight LiquorRation WorkingVolume WorkingLevel
  Public BatchWeight As Integer
  ' Public MachineVolume As Integer  ' Use VolumeBasedOnLevel later
  Public WaterUsed As Integer


  Public PackageHeight As Integer
  Public PackageType As Integer   'A-H = 1-9 for calculation
  Public WorkingLevel As Integer


  Public LidLocked As Boolean
  Public LevelOkToOpenLid As Boolean
#If 0 Then
    Public LidLockTimer As New acTimer
  Public OpenLid As Boolean
  Public CloseLid As Boolean
  Public LidControl As New acLidControl
#End If

  ' Level Control
  Public SetLevel As Integer 'todo
  Public IdleDrainActive As Boolean
  Public FirstFillComplete As Boolean

  ' Losing level variables
  Public GetRecordedLevel As Boolean
  Public GotRecordedLevel As Boolean
  Public GetRecordedLevelTimer As New Timer
  Public MachineLevelRecorded As Integer

  ' Add Tank Control
  Public AddControl As AddControl
  Public AddReady As Boolean

  ' Reserve Tank Control
  Public ReserveControl As ReserveControl
  Public ReserveReady As Boolean
  Public ReserveMixerOn As Boolean

  ' Kitchen Tank Ready Flags
  Public Tank1Ready As Boolean
  Public Tank1MixerRequest As Boolean
  Public Tank1MixerEnable As Boolean
  '  Public DrugroomMixerOn As Boolean


  <GraphTrace(0, 1000, 0, 5000, "Green", "%t%")> Public Tank1Level As Integer
  Private tank1MixerActiveTimer As New Timer
  Private tank1LevelScanTimer As New Timer

  ' TODO - use?
  Public Tank1TimePot As Integer
  Public Tank1TempPot As Integer

  ' Test Input (Aninp8)
  Public TestValue As Integer

  ' Look Ahead variables
  Public LAActive As Boolean
  Public LARequest As Boolean
  Public LASqlString As String
  Public LASetSqlString As Boolean

  ' Temperature Control
  Public ManualSteamTest As ManualSteamTest
  Public HeatEnabled As Boolean
  Public SteamRequest As Boolean
  Public SteamRequestTimer As New Timer

  'class module press control
  Public SafetyControl As New SafetyControl

  ' Airpad control
  Public AirpadBleed As Boolean
  Public AirpadBleedCount As Integer
  Public AirpadBleedDelayOff As New Timer
  Public AirpadCutoff As Boolean
  Public AirpadOn As Boolean


  ' Package Differential Pressure 
  ' Public PackageDpPsi As Integer
  Public PackageDpStr As String
  Public PackagePressureRange As Integer, PackagePressureUncorrected As Integer

  Public VolumeBasedOnLevel As Integer 'in tenths
  Public SystemVolume As Double


  Public FlowRate As Integer
  <GraphTrace(1, 1000, 0, 3333, "Purple", "%t %")> Public FlowRatePercent As Integer
  Public FlowRateTimer As New Timer
  Public FlowratePerWt As Integer           ' Metric Units: LitersPerMinute/Kilogram
  Public NumberOfContacts As Integer
  Public TotalNumberOfContacts As Integer
  Public NumberofContactsPerMin As Double


  ' Desired Flowrate
  Public MachineFlowRatePv As Integer
  Public MachineFlowRateSv As Integer
  Public MachineFlowRateTimer As New Timer



  '  Public PumpRequest As Boolean ' TODO - Check and use with PumpControl.Start


  ' TODO Use These?
  Public PressRunButton As Boolean
  Public PressPauseButton As Boolean


#If 0 Then

  'Class module for pump and reel control
  Private FlowReverseTimer As New acTimer
  Public FlowReverseInToOutTimer As New acTimer
  Public FlowReverseOutToInTimer As New acTimer
  Public PumpRequest As Boolean, PumpLevel As Boolean
  Private PumpLevelTimer As New acTimer
  Public FirstScanDoneTimer As New acTimer

  
'stuff for checking the pump on off count
  Public PumpOnCount As Long
  Public PumpWasRunning As Boolean
  Public ResetPumpOnCountTimer As New acTimer
  Public PumpOnCountTooHigh As Boolean
  Public PumpLevelLostTimer As New acTimer
  

  'flowrate stuff
  Public FlowRate As Long
  Public FlowRatePercent As Long
Attribute FlowRatePercent.VB_VarDescription = "MaximumValue=1000\r\nMaximumY=3333\r\nFormat=%t %\r\nColor=3\r\nOrder=3"
  Public FlowRateRange As Long, FlowRateUncorrected As Long
  
  
  'pump speed control
  Public PumpAccelerationTimer As New acTimer
  Public PumpSpeedAdjustTimer As New acTimer
  Public DPAdjust As Long, DPError As Long, InToOutSpeed As Long, OutToInSpeed As Long
  Public FLAdjust As Long, FLError As Long

  
'mimic stuff
'main screen
  Public MainPumpResetPB As Boolean
Attribute MainPumpResetPB.VB_VarDescription = "IO=Mimic"
  
'Addition tank
  Public ManualAddFillHot As Boolean
  Public ManualAddFillCold As Boolean
  Public ManualAddFillLevel As Long
  Public ManualAddDrainPB As Boolean
  Public ManualAddDrainFinished As Boolean
  Public ManualAddTransferPB As Boolean
  Public ManualAddTransferFinished As Boolean
  Public AddPreparePB As Boolean
  Public ManualAddFill As New ACManualAddFill
  Public ManualAddDrain As New ACManualAddDrain
  Public ManualAddTransfer As New ACManualAddTransfer
  
'Reserve Tank
  Public ManualReserveFillHot As Boolean
  Public ManualReserveFillCold As Boolean
  Public ManualReserveFillLevel As Long
  Public ManualReserveDrainPB As Boolean
  Public ManualReserveDrainFinished As Boolean
  Public ManualReserveTransferPB As Boolean
  Public ManualReserveTransferFinished As Boolean
  Public ManualReserveHeatPB As Boolean
  Public ManualReserveDesiredTemp As Long
  Public ReservePreparePB As Boolean
  Public ManualReserveFill As New ACManualReserveFill
  Public ManualReserveDrain As New ACManualReserveDrain
  Public ManualReserveTransfer As New ACManualReserveTransfer
  Public ManualReserveHeat As New ACManualReserveHeat
 
'Maintenance Steam Test
  Public SteamTestPB As Boolean
  Public ManualSteamOverride As Boolean
  Public ManualSteamTest As New ACManualSteamTest
  


  
' Drugroom Preview Variables
  Public StepNumberWas As Long
  Public StepOverrunMins As Long
  Public TimeInStepValueWas As Long
  Public DrugroomPreview As New acDrugroomPreview
  Public DrugroomCalloffRefreshed As Boolean
   
' Automatic Dispense Variables
  Public DrugroomDisplayJob1 As String
  Public DispenseCalloff As Long
  Public Dispensestate As Long        'This filled in by AutoDispenser
  Public Dispenseproducts As String
  Public DispenseTank As Long

  'losing level variables
  Public GetRecordedLevel As Boolean
  Public GotRecordedLevel As Boolean
  Public GetRecordedLevelTimer As New acTimer
  Public RecordedLevel As Long
  

#End If




  'For Production Reports
  Friend PowerKWS As Integer
  Public PowerKWH As Integer
  Public SteamUsed As Integer

  Public SteamNeeded As Integer
  Public TempRise As Integer
  Public FinalTempWas As Integer
  Public FinalTemp As Integer
  Public StartTemp As Integer

  '  Public Utilities_Timer As New acTimer
  '  Public OperatorDelayTimer As New acTimer
  ' private MainPumpHP as inte 



  'System shutdown flag
  Public SystemShuttingDown As Boolean
  Public FirstScanDone As Boolean
  Public FirstScanDoneTimer As New Timer
  Public PowerOnTimer As New Timer

  ' Hibernate
  Public SystemHibernate As Boolean
  Public SystemHibernateTimer As New Timer

  'Plant Explorer host status color
  Public StatusColor As Integer

End Class
