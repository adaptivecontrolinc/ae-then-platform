Public Class ControlCodevb6
#If 0 Then
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"AC"
Attribute VB_Ext_KEY = "Member1" ,"AF"
Attribute VB_Ext_KEY = "Member2" ,"AP"
Attribute VB_Ext_KEY = "Member3" ,"AV"
Attribute VB_Ext_KEY = "Member4" ,"BA"
Attribute VB_Ext_KEY = "Member5" ,"BO"
Attribute VB_Ext_KEY = "Member6" ,"BR"
Attribute VB_Ext_KEY = "Member7" ,"BS"
Attribute VB_Ext_KEY = "Member8" ,"BW"
Attribute VB_Ext_KEY = "Member9" ,"CC"
Attribute VB_Ext_KEY = "Member10" ,"CO"
Attribute VB_Ext_KEY = "Member11" ,"DR"
Attribute VB_Ext_KEY = "Member12" ,"FI"
Attribute VB_Ext_KEY = "Member13" ,"FP"
Attribute VB_Ext_KEY = "Member14" ,"HE"
Attribute VB_Ext_KEY = "Member15" ,"IJ"
Attribute VB_Ext_KEY = "Member16" ,"LD"
Attribute VB_Ext_KEY = "Member17" ,"LR"
Attribute VB_Ext_KEY = "Member18" ,"PD"
Attribute VB_Ext_KEY = "Member19" ,"PR"
Attribute VB_Ext_KEY = "Member20" ,"PS"
Attribute VB_Ext_KEY = "Member21" ,"RC"
Attribute VB_Ext_KEY = "Member22" ,"RD"
Attribute VB_Ext_KEY = "Member23" ,"RG"
Attribute VB_Ext_KEY = "Member24" ,"RI"
Attribute VB_Ext_KEY = "Member25" ,"RS"
Attribute VB_Ext_KEY = "Member26" ,"SA"
Attribute VB_Ext_KEY = "Member27" ,"TF"
Attribute VB_Ext_KEY = "Member28" ,"TM"
Attribute VB_Ext_KEY = "Member29" ,"TP"
Attribute VB_Ext_KEY = "Member30" ,"TT"
Attribute VB_Ext_KEY = "Member31" ,"UL"
Attribute VB_Ext_KEY = "Member32" ,"VL"
Attribute VB_Ext_KEY = "Member33" ,"VV"
Attribute VB_Ext_KEY = "Member34" ,"WT"
'===============================================================================================
'A&E Then Package Machines                      *Look to Version File for information & updates*
'===============================================================================================
  Implements ACControlCode
  Option Explicit
  
  Public Parent As Parent
  Public ProgramStoppedTimer As New acTimerUp 'Program Stopped Timer
  Public ProgramStoppedTime As Long
  Public TimeInStepValue As Long, StepStandardTime As Long, TimeInStep As String
  Public StepStandardTimeWas As Long
  Private TwoSecondTimer As New acTimer

'Define PLC "driver"
  Public IO_Modbus1 As ACSerialDevices.ACModbus
  Public IO_Modbus1Open As Boolean
  Public PLC1WatchdogTimer As New acTimer
  Private Const PLC1WatchdogTime As Long = 5

'Define PLC "driver"
  Public IO_Modbus2 As ACSerialDevices.ACModbus
  Public IO_Modbus2Open As Boolean
  Public PLC2WatchdogTimer As New acTimer
  Private Const PLC2WatchdogTime As Long = 5

'The Temp raw values (before scaling or smoothing) - for calibration work
  Public IO_TEMP1FromPLC As Integer, IO_TEMP2FromPLC As Integer
  Public IO_TEMP3FromPLC As Integer, IO_TEMP4FromPLC As Integer, IO_TEMP5FromPLC As Integer
  Public IO_TEMP1Smooth As Integer, IO_TEMP2Smooth As Integer
  Public IO_TEMP3Smooth As Integer, IO_TEMP4Smooth As Integer, IO_TEMP5Smooth As Integer

'The Aninp raw values (before scaling or smoothing) - for calibration work
  Public IO_ANINP1Raw As Integer, IO_ANINP2Raw As Integer
  Public IO_ANINP3Raw As Integer, IO_ANINP4Raw As Integer
  Public IO_ANINP5Raw As Integer, IO_ANINP6Raw As Integer
  Public IO_ANINP7Raw As Integer, IO_ANINP8Raw As Integer
  Public IO_ANINP9Raw As Integer, IO_ANINP10Raw As Integer
  Public IO_ANINP11Raw As Integer
  Public IO_ANINP1Smooth As Integer, IO_ANINP2Smooth As Integer
  Public IO_ANINP3Smooth As Integer, IO_ANINP4Smooth As Integer
  Public IO_ANINP5Smooth As Integer, IO_ANINP6Smooth As Integer
  Public IO_ANINP7Smooth As Integer, IO_ANINP8Smooth As Integer
  Public IO_ANINP9Smooth As Integer, IO_ANINP10Smooth As Integer
  Public IO_ANINP11Smooth As Integer

'temp analog outputs
  Public TempANOUT1 As Long
  Public TempANOUT2 As Long
  Public TempANOUT3 As Long
  Public TempANOUT4 As Long
 
'Get current time
  Public CurrentTime As String
  
'flasher stuff
  Public FastFlasher As New acFlasher: Public FastFlash As Boolean
  Public SlowFlasher As New acFlasher: Public SlowFlash As Boolean

'System shutdown flag
  Private SystemShuttingDown As Boolean

'===============================================================================================
'IO DEFINITIONS
'===============================================================================================
'
'MAIN PANEL
'Digital Inputs 1-8
  Public IO_RemoteRun As Boolean               'DINP  1
Attribute IO_RemoteRun.VB_VarDescription = "IO=DINP 1\r\nAllowOverride=True\r\nIcon=Pushbutton\r\nHelp=Remote run button to silence signal or alarm and advance the program."
  Public IO_RemoteHalt As Boolean              'DINP  2
Attribute IO_RemoteHalt.VB_VarDescription = "IO=DINP 2\r\nAllowOverride=True\r\nIcon=Pushbutton\r\nHelp=Remote halt. "
  Public IO_PowerOn As Boolean                 'DINP  3
Attribute IO_PowerOn.VB_VarDescription = "IO=DINP 3\r\nAllowOverride=True"
  'spare                                       'DINP  4
  Public IO_EStop_PB As Boolean                'DINP  5
Attribute IO_EStop_PB.VB_VarDescription = "IO=DINP 5\r\nAllowOverride=True\r\nIcon=Pushbutton\r\nHelp=E-stop input"
  Public IO_TempInterlock As Boolean           'DINP  6
Attribute IO_TempInterlock.VB_VarDescription = "IO=DINP 6\r\nAllowOverride=False"
  Public IO_PressureInterlock As Boolean       'DINP  7
Attribute IO_PressureInterlock.VB_VarDescription = "IO=DINP 7\r\nAllowOverride=False"
  Public IO_ContactThermometer As Boolean      'DINP  8
Attribute IO_ContactThermometer.VB_VarDescription = "IO=DINP 8\r\nAllowOverride=False\r\nIcon=TemperatureSwitch\r\nHelp=High temperature safety switch. Will be made when the temperature is less than 280F."
  Public IO_PressBelow1PSI1 As Boolean         'DINP  9
Attribute IO_PressBelow1PSI1.VB_VarDescription = "IO=DINP 9\r\nAllowOverride=False"
  Public IO_PressBelow1PSI2 As Boolean         'DINP 10
Attribute IO_PressBelow1PSI2.VB_VarDescription = "IO=DINP 10\r\nAllowOverride=False\r\nIcon=PressureSwitch\r\n"
  'spare                                       'DINP 11
  Public IO_MainPumpRunning As Boolean         'DINP 12
Attribute IO_MainPumpRunning.VB_VarDescription = "IO=DINP 12\r\nAllowOverride=False"
  Public IO_MainPumpFaulted As Boolean         'DINP 13
Attribute IO_MainPumpFaulted.VB_VarDescription = "IO=DINP 13\r\nAllowOverride=True"
  Public IO_AddPumpRunning As Boolean          'DINP 14
Attribute IO_AddPumpRunning.VB_VarDescription = "IO=DINP 14\r\nAllowOverride=False"
  Public IO_LockingPinRaised As Boolean        'DINP 15
Attribute IO_LockingPinRaised.VB_VarDescription = "IO=DINP 15\r\nAllowOverride=False"
  Public IO_LidLoweredSwitch As Boolean        'DINP 16
Attribute IO_LidLoweredSwitch.VB_VarDescription = "IO=DINP 16\r\nAllowOverride=False"
  
'Digital Inputs 17-24
  Public IO_OpenLid_PB As Boolean              'DINP 17
Attribute IO_OpenLid_PB.VB_VarDescription = "IO=DINP 17\r\nAllowOverride=False"
  Public IO_CloseLid_PB As Boolean             'DINP 18
Attribute IO_CloseLid_PB.VB_VarDescription = "IO=DINP 18\r\nAllowOverride=False"
  Public IO_DyeDispenseReserve As Boolean      'DINP 19
Attribute IO_DyeDispenseReserve.VB_VarDescription = "IO=DINP 19\r\nAllowOverride=False"
  Public IO_DyeDispenseAdd As Boolean          'DINP 20
Attribute IO_DyeDispenseAdd.VB_VarDescription = "IO=DINP 20\r\nAllowOverride=False"
  Public IO_LidRaised_SW As Boolean            'DINP 21
Attribute IO_LidRaised_SW.VB_VarDescription = "IO=DINP 21\r\nAllowOverride=False"
  'spare                                       'DINP 22
  'spare                                       'DINP 23
  'spare                                       'DINP 24
  
'digital inputs in drugroom
'digital Inputs 65-72
  Public IO_Tank1Mixer_PB As Boolean           'DINP 65
Attribute IO_Tank1Mixer_PB.VB_VarDescription = "IO=DINP 65\r\nAllowOverride=False"
  Public IO_Tank1FillCold_PB As Boolean        'DINP 66
Attribute IO_Tank1FillCold_PB.VB_VarDescription = "IO=DINP 66\r\nAllowOverride=False"
  Public IO_Tank1FillHot_PB As Boolean         'DINP 67
Attribute IO_Tank1FillHot_PB.VB_VarDescription = "IO=DINP 67\r\nAllowOverride=False"
  Public IO_Tank1Heating_PB As Boolean         'DINP 68
Attribute IO_Tank1Heating_PB.VB_VarDescription = "IO=DINP 68\r\nAllowOverride=False"
  Public IO_Tank1Ready_PB As Boolean           'DINP 69
Attribute IO_Tank1Ready_PB.VB_VarDescription = "IO=DINP 69\r\nAllowOverride=False"
  Public IO_Tank1Manual_SW As Boolean          'DINP 70
Attribute IO_Tank1Manual_SW.VB_VarDescription = "IO=DINP 70\r\nAllowOverride=False"
  'spare                                       'DINP 71
  'spare                                       'DINP 72
  Public IO_CityWater_SW As Boolean            'DINP 73
Attribute IO_CityWater_SW.VB_VarDescription = "IO=DINP 73\r\nAllowOverride=False"
  
'Main Panel digital outputs
'Digital Output 1-12
  Public IO_Siren As Boolean                   'DOUT  1
Attribute IO_Siren.VB_VarDescription = "IO=DOUT 1\r\nAllowOverride=True\r\nIcon=RedBeacon\r\nHelp=Alarm beacon. On while an alarm has not been acknowledged."
  Public IO_BeaconAlarm As Boolean             'DOUT  2
Attribute IO_BeaconAlarm.VB_VarDescription = "IO=DOUT 2\r\nAllowOverride=True\r\nIcon=AmberBeacon\r\nHelp=Signal beacon. On while an operator signal has not been acknowledged."
  Public IO_BeaconSignal As Boolean            'DOUT  3
Attribute IO_BeaconSignal.VB_VarDescription = "IO=DOUT 3\r\nAllowOverride=True\r\nIcon=Lamp\r\nHelp="
  Public IO_BeaconDelay As Boolean             'DOUT  4
Attribute IO_BeaconDelay.VB_VarDescription = "IO=DOUT 4\r\nAllowOverride=True\r\nIcon=Siren\r\nHelp=Alarm/signal siren. On while an alarm or operator signal has not been acknowledged."
  Public IO_SafeLamp As Boolean                'DOUT  5
Attribute IO_SafeLamp.VB_VarDescription = "IO=DOUT 5\r\nAllowOverride=True\r\nIcon=GreenLamp\r\nHelp=Machine safe lamp. On when machine has de-pressurized."
  Public IO_HxDrain As Boolean                 'DOUT  6
Attribute IO_HxDrain.VB_VarDescription = "IO=DOUT 6\r\nAllowOverride=True\r\nHelp=Heat Exchanger Drain valve, On if not heating or cooling.  Not installed on all machines."
  Public IO_MainPumpStart As Boolean           'DOUT  7
Attribute IO_MainPumpStart.VB_VarDescription = "IO=DOUT 7\r\nAllowOverride=True"
  Public IO_MainPumpReset As Boolean           'DOUT  8
Attribute IO_MainPumpReset.VB_VarDescription = "IO=DOUT 8\r\nAllowOverride=True"
  Public IO_AddPumpStart As Boolean            'DOUT  9
Attribute IO_AddPumpStart.VB_VarDescription = "IO=DOUT 9\r\nAllowOverride=True"
  Public IO_HoldDown_Coupling As Boolean       'DOUT 10
Attribute IO_HoldDown_Coupling.VB_VarDescription = "IO=DOUT 10\r\nAllowOverride=True"
  Public IO_SystemBlock As Boolean             'DOUT 11
Attribute IO_SystemBlock.VB_VarDescription = "IO=DOUT 11\r\nAllowOverride=True"
  Public IO_LevelGaugeFlush As Boolean         'DOUT 12
Attribute IO_LevelGaugeFlush.VB_VarDescription = "IO=DOUT 12\r\nAllowOverride=True"
  
'Digital Output 13-24
  Public IO_Vent As Boolean                    'DOUT 13
Attribute IO_Vent.VB_VarDescription = "IO=DOUT 13\r\nAllowOverride=True"
  Public IO_Airpad As Boolean                  'DOUT 14
Attribute IO_Airpad.VB_VarDescription = "IO=DOUT 14\r\nAllowOverride=True"
  Public IO_HeatSelect As Boolean              'DOUT 15
Attribute IO_HeatSelect.VB_VarDescription = "IO=DOUT 15\r\nAllowOverride=True"
  Public IO_CoolSelect As Boolean              'DOUT 16
Attribute IO_CoolSelect.VB_VarDescription = "IO=DOUT 16\r\nAllowOverride=True"
  Public IO_Condensate As Boolean              'DOUT 17
Attribute IO_Condensate.VB_VarDescription = "IO=DOUT 17\r\nAllowOverride=True"
  Public IO_CoolWaterReturn As Boolean         'DOUT 18
Attribute IO_CoolWaterReturn.VB_VarDescription = "IO=DOUT 18\r\nAllowOverride=True"
  Public IO_Fill As Boolean                    'DOUT 19
Attribute IO_Fill.VB_VarDescription = "IO=DOUT 19\r\nAllowOverride=True"
  Public IO_TopWash As Boolean                 'DOUT 20
Attribute IO_TopWash.VB_VarDescription = "IO=DOUT 20\r\nAllowOverride=True"
  Public IO_Drain As Boolean                   'DOUT 21
Attribute IO_Drain.VB_VarDescription = "IO=DOUT 21\r\nAllowOverride=True"
  Public IO_HotDrain As Boolean                'DOUT 22
Attribute IO_HotDrain.VB_VarDescription = "IO=DOUT 22\r\nAllowOverride=True"
  Public IO_FlowReverse As Boolean             'DOUT 23
Attribute IO_FlowReverse.VB_VarDescription = "IO=DOUT 23\r\nAllowOverride=True"
  Public IO_OpenLockingBand As Boolean         'DOUT 24
Attribute IO_OpenLockingBand.VB_VarDescription = "IO=DOUT 24\r\nAllowOverride=True"
  
'Digital Output 25-36
  Public IO_OutletInToOut As Boolean           'DOUT 25
Attribute IO_OutletInToOut.VB_VarDescription = "IO=DOUT 25\r\nAllowOverride=True"
  Public IO_OutletOutToIn As Boolean           'DOUT 26
Attribute IO_OutletOutToIn.VB_VarDescription = "IO=DOUT 26\r\nAllowOverride=True"
  Public IO_InletInToOut As Boolean            'DOUT 27
Attribute IO_InletInToOut.VB_VarDescription = "IO=DOUT 27\r\nAllowOverride=True"
  Public IO_InletOutToIn As Boolean            'DOUT 28
Attribute IO_InletOutToIn.VB_VarDescription = "IO=DOUT 28\r\nAllowOverride=True"
  Public IO_CloseLockingBand As Boolean        'DOUT 29
Attribute IO_CloseLockingBand.VB_VarDescription = "IO=DOUT 29\r\nAllowOverride=True"
  Public IO_AddTankColdFill As Boolean         'DOUT 30
Attribute IO_AddTankColdFill.VB_VarDescription = "IO=DOUT 30\r\nAllowOverride=True"
  Public IO_AddTankHotFill As Boolean          'DOUT 31
Attribute IO_AddTankHotFill.VB_VarDescription = "IO=DOUT 31\r\nAllowOverride=True"
  Public IO_AddTankDrain As Boolean            'DOUT 32
Attribute IO_AddTankDrain.VB_VarDescription = "IO=DOUT 32\r\nAllowOverride=True"
  Public IO_AddTankTransfer As Boolean         'DOUT 33
Attribute IO_AddTankTransfer.VB_VarDescription = "IO=DOUT 33\r\nAllowOverride=True"
  Public IO_AddTankMixing As Boolean           'DOUT 34
Attribute IO_AddTankMixing.VB_VarDescription = "IO=DOUT 34\r\nAllowOverride=True"
  Public IO_AddTankRunback As Boolean          'DOUT 35
Attribute IO_AddTankRunback.VB_VarDescription = "IO=DOUT 35\r\nAllowOverride=True"
  Public IO_ReserveColdFill As Boolean         'DOUT 36
Attribute IO_ReserveColdFill.VB_VarDescription = "IO=DOUT 36\r\nAllowOverride=True"
  
'Digital Output 37-48
  Public IO_ReserveHotFill As Boolean          'DOUT 37
Attribute IO_ReserveHotFill.VB_VarDescription = "IO=DOUT 37\r\nAllowOverride=True"
  Public IO_LowerLockingPin As Boolean         'DOUT 38
Attribute IO_LowerLockingPin.VB_VarDescription = "IO=DOUT 38\r\nAllowOverride=True"
  Public IO_ReserveDrain As Boolean            'DOUT 39
Attribute IO_ReserveDrain.VB_VarDescription = "IO=DOUT 39\r\nAllowOverride=True"
  Public IO_ReserveTransfer As Boolean         'DOUT 40
Attribute IO_ReserveTransfer.VB_VarDescription = "IO=DOUT 40\r\nAllowOverride=True"
  Public IO_ReserveBackflow As Boolean         'DOUT 41
Attribute IO_ReserveBackflow.VB_VarDescription = "IO=DOUT 41\r\nAllowOverride=True"
  Public IO_ReserveHeating As Boolean          'Dout 42
Attribute IO_ReserveHeating.VB_VarDescription = "IO=DOUT 42\r\nAllowOverride=True"
  Public IO_ReserveMixer As Boolean            'DOUT 43
Attribute IO_ReserveMixer.VB_VarDescription = "IO=DOUT 43\r\nAllowOverride=True"
  Public IO_Tank1ToDrain As Boolean            'DOUT 44
Attribute IO_Tank1ToDrain.VB_VarDescription = "IO=DOUT 44\r\nAllowOverride=True"
  Public IO_Tank1ToAdd As Boolean              'DOUT 45
Attribute IO_Tank1ToAdd.VB_VarDescription = "IO=DOUT 45\r\nAllowOverride=True"
  Public IO_Tank1ToReserve As Boolean          'DOUT 46
Attribute IO_Tank1ToReserve.VB_VarDescription = "IO=DOUT 46\r\nAllowOverride=True"
  Public IO_OpenLid As Boolean                 'DOUT 47
Attribute IO_OpenLid.VB_VarDescription = "IO=DOUT 47\r\nAllowOverride=True"
  Public IO_CloseLid As Boolean                'DOUT 48
Attribute IO_CloseLid.VB_VarDescription = "IO=DOUT 48\r\nAllowOverride=True"
  Public IO_PlcWatchDog1 As Boolean            'DOUT 57 not really an output
Attribute IO_PlcWatchDog1.VB_VarDescription = "IO=DOUT 57\r\nAllowOverride=True"
  
'Drugroom digital outputs
  Public IO_Tank1Mixer As Boolean              'DOUT 65
Attribute IO_Tank1Mixer.VB_VarDescription = "IO=DOUT 65\r\nAllowOverride=True\r\n"
  Public IO_Tank1ColdFill As Boolean           'DOUT 66
Attribute IO_Tank1ColdFill.VB_VarDescription = "IO=DOUT 66\r\nAllowOverride=True\r\nIcon=Valve"
  Public IO_Tank1Steam As Boolean              'DOUT 67
Attribute IO_Tank1Steam.VB_VarDescription = "IO=DOUT 67\r\nAllowOverride=True\r\nIcon=Valve"
  Public IO_Tank1PrepareLamp As Boolean        'DOUT 68
Attribute IO_Tank1PrepareLamp.VB_VarDescription = "IO=DOUT 68\r\nAllowOverride=True\r\n"
  Public IO_Tank1Transfer As Boolean           'DOUT 69
Attribute IO_Tank1Transfer.VB_VarDescription = "IO=DOUT 69\r\nAllowOverride=True"
  Public IO_Tank1HotFill As Boolean            'DOUT 70
Attribute IO_Tank1HotFill.VB_VarDescription = "IO=DOUT 70\r\nAllowOverride=True"
  'dout 71 doesnt exist the first card only has 6 outputs
  'dout 72 doesnt exist the first card only has 6 outputs
  
  Public IO_CityWaterFill As Boolean           'dout 73
Attribute IO_CityWaterFill.VB_VarDescription = "IO=DOUT 73\r\nAllowOverride=True"
  Public IO_DispenseStateLamp As Boolean       'Dout 74
Attribute IO_DispenseStateLamp.VB_VarDescription = "IO=DOUT 74\r\nAllowOverride=True"

  Public IO_PlcWatchDog2 As Boolean            'DOUT 81 not really an output
Attribute IO_PlcWatchDog2.VB_VarDescription = "IO=DOUT 81\r\nAllowOverride=True\r\n"

'Temperature inputs
  Public IO_VesselTemp As Integer              'TEMP  1
Attribute IO_VesselTemp.VB_VarDescription = "IO=TEMP 1\r\nAllowOverride=False\r\nIcon=TemperatureProbe\r\nHelp=Machine temperature probe, located after the heat exchanger."
  Public IO_ReserveTemp As Integer             'TEMP  2
Attribute IO_ReserveTemp.VB_VarDescription = "IO=TEMP 2\r\nAllowOverride=False"
  Public IO_BlendFillTemp As Integer           'TEMP  3
Attribute IO_BlendFillTemp.VB_VarDescription = "IO=TEMP 3\r\nAllowOverride=False\r\nIcon=TemperatureProbe\r\nHelp=blend fill temps."
  Public IO_CondensateTemp As Integer          'TEMP  4
Attribute IO_CondensateTemp.VB_VarDescription = "IO=TEMP 4\r\nAllowOverride=False\r\nIcon=TemperatureProbe\r\nHelp=Temperature for Condensate Trap return temp."
    
'Drugroom temperature inputs
  Public IO_Tank1Temp As Integer               'TEMP 5
Attribute IO_Tank1Temp.VB_VarDescription = "IO=TEMP 5\r\nAllowOverride=False\r\nIcon=TemperatureProbe\r\n"
  
'Analog inputs
  Public IO_VesselLevelInput As Integer        'ANINP 1
Attribute IO_VesselLevelInput.VB_VarDescription = "IO=ANINP 1\r\nAllowOverride=true "
  Public IO_ReserveLevelInput As Integer       'ANINP 2
Attribute IO_ReserveLevelInput.VB_VarDescription = "IO=ANINP 2\r\nAllowOverride=true "
  Public IO_AdditionLevelInput As Integer      'ANINP 3
Attribute IO_AdditionLevelInput.VB_VarDescription = "IO=ANINP 3\r\nAllowOverride=true "
  Public IO_PackageDiffPressInput As Integer   'ANINP 4
Attribute IO_PackageDiffPressInput.VB_VarDescription = "IO=ANINP 4\r\nAllowOverride=true "
  Public IO_VesselPressInput As Integer        'ANINP 5
Attribute IO_VesselPressInput.VB_VarDescription = "IO=ANINP 5\r\nAllowOverride=true "
  Public IO_FlowRateInput As Integer           'ANINP 6
Attribute IO_FlowRateInput.VB_VarDescription = "IO=ANINP 6\r\nAllowOverride=true "
  Public IO_VesTempInput As Integer            'ANINP 7
Attribute IO_VesTempInput.VB_VarDescription = "IO=ANINP 7\r\nAllowOverride=true "
  Public IO_TestInput As Integer               'ANINP 8
Attribute IO_TestInput.VB_VarDescription = "IO=ANINP 8\r\nAllowOverride=true "
  
'Drugroom analog inputs
  Public IO_Tank1LevelInput As Integer         'ANINP 9
Attribute IO_Tank1LevelInput.VB_VarDescription = "IO=ANINP 9\r\nAllowOverride=true "
  Public IO_TimePotInput As Integer            'ANINP 10
Attribute IO_TimePotInput.VB_VarDescription = "IO=ANINP 10\r\nAllowOverride=true "
  Public IO_TempPotInput As Integer            'ANINP 11
Attribute IO_TempPotInput.VB_VarDescription = "IO=ANINP 11\r\nAllowOverride=true "
    
'Analog outputs
  Public IO_HeatCoolOutput As Integer          'ANOUT  1
Attribute IO_HeatCoolOutput.VB_VarDescription = "IO=ANOUT 1\r\nAllowOverride=true\r\nIcon=GlobeValve\r\nHelp=Heat/cool Output (4-20mA)."
  Public IO_BlendFillOutput As Integer         'ANOUT  2
Attribute IO_BlendFillOutput.VB_VarDescription = "IO=ANOUT 2\r\nAllowOverride=true"
  'spare                                       'ANOUT  3
  Public IO_PumpSpeedOutput As Integer         'ANOUT  4
Attribute IO_PumpSpeedOutput.VB_VarDescription = "IO=ANOUT 4\r\nAllowOverride=true "
  
'
'===============================================================================================
'END IO DEFINITIONS
'===============================================================================================

'Stuff to log
  Public SetpointF As Long
Attribute SetpointF.VB_VarDescription = "MaximumValue=3000\r\nMaximumY=10000\r\nMinimumValue=1\r\nFormat=%t F\r\nColor=1\r\nOrder=1"
  Public TempFinalValue As String
  Public BatchWeight As Long
  
'Number of packages
  Public PackageHeight As Long
  Public PackageType As Long 'A-H = 1-9 for calculation
  Public WorkingLevel As Long
  
'Safety variables
  Public MachineSafe As Boolean
  Public PressSafe As Boolean
  Public TempSafe As Boolean
  Public TempValid As Boolean
  Public VesselHeatSafe As Boolean
  
'Variables for temp, level
  Public VesTemp As Long
Attribute VesTemp.VB_VarDescription = "MaximumValue=3000\r\nMaximumY=10000\r\nLabel=50,500\r\nLabel=100,1000\r\nLabel=150,1500\r\nLabel=200,2000\r\nLabel=250,2500\r\nFormat=%t F\r\nColor=2\r\nOrder=2\r\n"
  Private CycleTime As Long
  Private LastProgramCycleTime As Long
  Public CycleTimeDisplay As String

'time on status stuff
  Public TimeToGo As Long

'importing stuff from database
  Public StandardTime As Double
  Public StartTime As String
  Public EndTime As String
  Public EndTimeMins As Long
  Public GetEndTimeTimer As New acTimer
  
'lid locked stuff
  Public LidLocked As Boolean
  Public LevelOkToOpenLid As Boolean
  Public LidLockTimer As New acTimer
  Public OpenLid As Boolean
  Public CloseLid As Boolean
  Public LidControl As New acLidControl

'Class module for safety control
  Public SafetyControl As New acSafetyControl

'Class module for temperature control
  Public TemperatureControl As New acTemperatureControl

'Class module for temperature control contacts
  Public TemperatureControlContacts As New acTempControlContacts

'Class module for pump and reel control
  Private FlowReverseTimer As New acTimer
  Public FlowReverseInToOutTimer As New acTimer
  Public FlowReverseOutToInTimer As New acTimer
  Public PumpRequest As Boolean, PumpLevel As Boolean
  Private PumpLevelTimer As New acTimer
  Public FirstScanDoneTimer As New acTimer

'pause the program if sleeping
  Public IsSleepingWas As Boolean
  
'for bath turnover control
 Public VolumeBasedOnLevel As Long 'in tenths
 Public SystemVolume As Double
 Public FlowRateTimer As New acTimer
 Public GallonsPerMinPerPound As Long
 Public NumberOfContacts As Long
 Public TotalNumberOfContacts As Long
 Public NumberofContactsPerMin As Double

'desired flowrate
  Public DesiredFlowRate As Long
  Public LowFlowRateTimer As New acTimer
  
'desired pressure
  Public DesiredPressure As Long

'Airpad control
  Public AirpadOn As Boolean
  Public AirpadCutoff As Boolean
  Public AirpadBleed As Boolean
 
'drugroom mixer
  Public DrugroomMixerOn As Boolean
  Public Tank1MixerRequest As Boolean
 
'Reserve Tank mixer
  Public ReserveMixerOn As Boolean
 
'stuff for checking the pump on off count
  Public PumpOnCount As Long
  Public PumpWasRunning As Boolean
  Public ResetPumpOnCountTimer As New acTimer
  Public PumpOnCountTooHigh As Boolean
  Public PumpLevelLostTimer As New acTimer

'Addition tank mix
  Public AdditionMixOn As Boolean
 
'Vessel level
  Public VesselLevel As Long
Attribute VesselLevel.VB_VarDescription = "MaximumValue=1000\r\nMaximumY=3333\r\nFormat=%t %\r\nColor=4\r\nOrder=4"
  Public VesselLevelRange As Long, VesselLevelUncorrected As Long

'Reserve tank level
  Public ReserveLevel As Long
Attribute ReserveLevel.VB_VarDescription = "MaximumValue=1000\r\nMaximumY=3333\r\nFormat=%t %\r\nColor=5\r\nOrder=5"
  Public ReserveLevelRange As Long, ReserveLevelUncorrected As Long

'Addition tank level
  Public AdditionLevel As Long
Attribute AdditionLevel.VB_VarDescription = "MaximumValue=1000\r\nMaximumY=3333\r\nFormat=%t %\r\nColor=6\r\nOrder=6"
  Public AdditionLevelRange As Long, AdditionLevelUncorrected As Long

'package differential pressure
  Public PackageDifferentialPressure As Long
Attribute PackageDifferentialPressure.VB_VarDescription = "MaximumValue=1000\r\nMaximumY=3333\r\nFormat=%t psi\r\nColor=9\r\nOrder=9"
  Public PackagePressureRange As Long, PackagePressureUncorrected As Long

'Vessel Pressure
  Public VesselPressure As Long
Attribute VesselPressure.VB_VarDescription = "MaximumValue=1000\r\nMaximumY=3333\r\nFormat=%t psi\r\nColor=8\r\nOrder=8"
  Public VesselPressureRange As Long, VesselPressureUncorrected As Long
  
'flowrate stuff
  Public FlowRate As Long
  Public FlowRatePercent As Long
Attribute FlowRatePercent.VB_VarDescription = "MaximumValue=1000\r\nMaximumY=3333\r\nFormat=%t %\r\nColor=3\r\nOrder=3"
  Public FlowRateRange As Long, FlowRateUncorrected As Long
  
'TestInput stuff (Aninp8)
  Public TestValue As Long
  Public TestValuePercent As Long
  Public TestValueRange As Long, TestValueUncorrected As Long
    
'Tank 1 level
  Public Tank1Level As Long
Attribute Tank1Level.VB_VarDescription = "MaximumValue=1000\r\nMaximumY=3333\r\nFormat=%t %\r\nColor=7\r\nOrder=7"
  Public Tank1LevelRange As Long, Tank1LevelUncorrected As Long
  
'Tank 1 Time pot
  Public Tank1TimePot As Long
  Public Tank1TimePotRange As Long, Tank1TimePotUncorrected As Long

'Tank 1 Temperature Pot
  Public Tank1TempPot As Long
  Public Tank1TempPotRange As Long, Tank1TempPotUncorrected As Long
 
' Drugroom Tank Control
  Public Tank1Ready As Boolean

'addition tank
  Public AddReady As Boolean
  Public AddReadyFromPB As Boolean
  
'reserve tank
  Public ReserveReady As Boolean
  Public ReserveReadyFromPB As Boolean
   
'LA is active flag
  Public LAActive As Boolean
  Public LARequest As Boolean

'pump speed control
  Public PumpAccelerationTimer As New acTimer
  Public PumpSpeedAdjustTimer As New acTimer
  Public DPAdjust As Long, DPError As Long, InToOutSpeed As Long, OutToInSpeed As Long
  Public FLAdjust As Long, FLError As Long

'===============================================================================================
'PARAMETERS
'===============================================================================================
'
'Setup
  Public Parameters_InitializeParameters As Long
Attribute Parameters_InitializeParameters.VB_VarDescription = "Category=Setup\r\nHelp=If set to the correct code, all of the parameters will be set to default values."
  Public Parameters_SmoothRate As Long
Attribute Parameters_SmoothRate.VB_VarDescription = "Category=Setup\r\nHelp=Smothing rate to filter noise on the analog inputs."
  Public Parameters_MainPumpHP As Long
Attribute Parameters_MainPumpHP.VB_VarDescription = "Category=Setup\r\nHelp=The main pump hp"
  
  Public Parameters_TempTransmitterMin As Long
Attribute Parameters_TempTransmitterMin.VB_VarDescription = "Category=Setup\r\nHelp=Minimum analog input from 4-20mA temperature transmitter.  Used to scale and calibrate displayed value."
  Public Parameters_TempTransmitterMax As Long
Attribute Parameters_TempTransmitterMax.VB_VarDescription = "Category=Setup\r\nHelp=Maximum analog input, in tenths percent, from 4-20mA temperature transmitter.  Used to scale and calibrate displayed value. (1000 = 100.0%)"
  Public Parameters_TempTransmitterRange As Long
Attribute Parameters_TempTransmitterRange.VB_VarDescription = "Category=Setup\r\nHelp=Total Range value for 4-20ma temperature transmitter, in tenths degree F.  (3000 = 300.0F)"
  
  Public Parameters_TestValueMin As Long
Attribute Parameters_TestValueMin.VB_VarDescription = "Category=Setup\r\nHelp=Minimum analog input from 4-20mA temperature transmitter.  Used to scale and calibrate displayed value."
  Public Parameters_TestValueMax As Long
Attribute Parameters_TestValueMax.VB_VarDescription = "Category=Setup\r\nHelp=Maximum analog input from 4-20mA temperature transmitter.  Used to scale and calibrate displayed value."
  Public Parameters_TestValueRange As Long
Attribute Parameters_TestValueRange.VB_VarDescription = "Category=Setup\r\nHelp=Total Range value for 4-20ma transmitter, in tenths.  Units not defined."
  Public Parameters_TestValueLimit As Long
Attribute Parameters_TestValueLimit.VB_VarDescription = "Category=Setup\r\nHelp=Total Range limit value, in tenths, above which alarm will activate."
  Public Parameters_TestValueLimitTime As Long
Attribute Parameters_TestValueLimitTime.VB_VarDescription = "Category=Setup\r\nHelp=Time, in seconds, to delay Test value over limit."
  
  Public Parameters_TempCondensateLimit As Long
  Public Parameters_TempCondensateLimitTime As Long
  
  Public Parameters_SmoothRateRTD As Long
Attribute Parameters_SmoothRateRTD.VB_VarDescription = "Category=Setup\r\nHelp=Smooth rate for the RTD Inputs, to filter noise on the RTD input signals."
  
'Dispense
  Public Parameters_DispenseEnabled As Long
Attribute Parameters_DispenseEnabled.VB_VarDescription = "Category=Dispense\r\nHelp=Enables the AutoDispenser on the AdaptiveServer to schedule dispenses."
  Public Parameters_DDSEnabled As Long
Attribute Parameters_DDSEnabled.VB_VarDescription = "Category=Dispense\r\nHelp=Set to '1' for machines using the DDS; determines when the drugroom transfer valves open and looks for the DDS destination inputs."
  Public Parameters_DispenseResponseDelayTime As Long
Attribute Parameters_DispenseResponseDelayTime.VB_VarDescription = "Category=Dispense\r\nHelp=Time, in minutes, to wait before alarming DispenseResponseDelay.  Set to '0' to disable the alarm."
  Public Parameters_DispenseReadyDelayTime As Long
Attribute Parameters_DispenseReadyDelayTime.VB_VarDescription = "Category=Dispense\r\nHelp=Time, in minutes, to wait before alarming DispenseReadyDelay.  Set to '0' to disable the alarm."

'addition tank control
  Public Parameters_AddMixOffLevel As Long
Attribute Parameters_AddMixOffLevel.VB_VarDescription = "Category=Addition Tank Control\r\nHelp=The level in tenths of a percent to turn off the addition tank mixing."
  Public Parameters_AddMixOnLevel As Long
Attribute Parameters_AddMixOnLevel.VB_VarDescription = "Category=Addition Tank Control\r\nHelp=The level in tenths of a percent to turn on the addition tank mixing."
  Public Parameters_AddTransferRinseTime As Long
Attribute Parameters_AddTransferRinseTime.VB_VarDescription = "Category=Addition Tank Control\r\nHelp=The time in seconds to rinse the addition tank to the vessel."
  Public Parameters_AddTransferRinseToDrainTime As Long
Attribute Parameters_AddTransferRinseToDrainTime.VB_VarDescription = "Category=Addition Tank Control\r\nHelp=The time in seconds to rinse the addition tank to drain."
  Public Parameters_AddTransferTimeAfterRinse As Long
Attribute Parameters_AddTransferTimeAfterRinse.VB_VarDescription = "Category=Addition Tank Control\r\nHelp=The time in seconds to transfer the addition tank to the vessel after rinsing."
  Public Parameters_AddTransferTimeBeforeRinse As Long
Attribute Parameters_AddTransferTimeBeforeRinse.VB_VarDescription = "Category=Addition Tank Control\r\nHelp=The time in seconds to transfer the addition tank to the vessel before rinsing."
  Public Parameters_AddTransferToDrainTime As Long
Attribute Parameters_AddTransferToDrainTime.VB_VarDescription = "Category=Addition Tank Control\r\nHelp=The time in seconds to transfer the addition tank to drain."
  Public Parameters_AdditionFillDeadband As Long
Attribute Parameters_AdditionFillDeadband.VB_VarDescription = "Category=Addition Tank Control\r\nHelp=The fill level deadband, in tenths of a percent. due to water pressure."
  Public Parameters_AdditionMaxTransferLevel As Long
Attribute Parameters_AdditionMaxTransferLevel.VB_VarDescription = "Category=Addition Tank Control\r\nHelp=If the addition tank level is above this level. The kitchen tank will not transfer to the addition tank."
  Public Parameters_AddRinseFillLevel As Long
Attribute Parameters_AddRinseFillLevel.VB_VarDescription = "Category=Addition Tank Control\r\nHelp=The level, in tenths, to fill the add tank too with rinse water, to flush the mix lines before sending to drain."
  Public Parameters_AddRinseMixTime As Long
Attribute Parameters_AddRinseMixTime.VB_VarDescription = "Category=Addition Tank Control\r\nHelp=The time to mix the rinse water through mix lines, before transfering to drain."
  Public Parameters_AddRinseMixPulseTime As Long
Attribute Parameters_AddRinseMixPulseTime.VB_VarDescription = "Category=Addition Tank Control\r\nHelp=The time delay before pulsing the addition mix valve off for 1second.  Set this value to greater than AddRinseMixTime to disable the addition mix pulsing procedure, in the event that add pump tripping occurs."
  Public Parameters_AddTransferSettleTime As Long
Attribute Parameters_AddTransferSettleTime.VB_VarDescription = "Category=Addition Tank Control\r\nHelp=The time, in seconds, to wait for the add tank to settle after the first rinse, before pumping the rinse water to the machine.  Should prevent long delays during add transfer commands."
  Public Parameters_AddRunbackPulseTime As Long
Attribute Parameters_AddRunbackPulseTime.VB_VarDescription = "Category=Addition Tank Control\r\nHelp=The time, in seconds, to pulse the Add Runback Valve both open and close during the AF Fill From Vessel Command to prevent overfilling.  Set to '0' to disable this timer, where the Runback will remain open until the D"
    
'airpad control
  Public Parameters_AirpadCutoffPressure As Long
Attribute Parameters_AirpadCutoffPressure.VB_VarDescription = "Category=Airpad Control\r\nHelp=The pressure, in tenths of a psi, to turn off the airpad valve at."
  Public Parameters_AirpadBleedPressure As Long
Attribute Parameters_AirpadBleedPressure.VB_VarDescription = "Category=Airpad Control\r\nHelp=The pressure, in tenths of a psi, to turn on the vent valve to bleed off pressure. If the pressure is to high."

'Blend Control
  Public Parameters_BlendDeadBand As Long
Attribute Parameters_BlendDeadBand.VB_VarDescription = "Category=Blend Control\r\nHelp=If the blend fill temperature error from the setpoint is less than this, in tenths F, the valve doesn't adjust"
  Public Parameters_BlendFactor As Long
Attribute Parameters_BlendFactor.VB_VarDescription = "Category=Blend Control\r\nHelp=Determines how fast we adjust the blend valve per degree F of error from the setpoint during a fill or rinse\r\n"
  Public Parameters_ColdWaterTemperature As Long
Attribute Parameters_ColdWaterTemperature.VB_VarDescription = "Category=Blend Control\r\nHelp=Set to the temperature of the cold water supplied to the machine, in tenths fahrenheit."
  Public Parameters_HotWaterTemperature As Long
Attribute Parameters_HotWaterTemperature.VB_VarDescription = "Category=Blend Control\r\nMinimum=320\r\nMaximum=1800\r\nHelp=Set to the temperature of the hot water supplied to the machine, in tenths fahrenheit.\r\n"

'differential pressure control
  Public Parameters_DPAdjustFactor As Long
Attribute Parameters_DPAdjustFactor.VB_VarDescription = "Category=Diff. Pressure Control\r\nHelp=During a DP command, the pump speed adjustment is proportional to this factor multiplied by the DP error."
  Public Parameters_DPAdjustMaximum As Long
Attribute Parameters_DPAdjustMaximum.VB_VarDescription = "Category=Diff. Pressure Control\r\nHelp=During a DP command, this is the maximum change allowed in pump speed, in tenths percent."
  Public Parameters_MaxDifferentialPressure As Long
Attribute Parameters_MaxDifferentialPressure.VB_VarDescription = "Category=Diff. Pressure Control\r\nHelp=The Maximum Differential pressure allowed in tenths of a psi."

'Drugroom control
  Public Parameters_CKFillLevel As Long
Attribute Parameters_CKFillLevel.VB_VarDescription = "Category=Drugroom\r\nHelp=The level to fill the drugroom tank to during a CK command. in tenths of a percent."
  Public Parameters_CKTransferToDrainTime1 As Long
Attribute Parameters_CKTransferToDrainTime1.VB_VarDescription = "Category=Drugroom\r\nHelp=The time in seconds to transfer tank 1 to drain before the rinse."
  Public Parameters_CKRinseTime As Long
Attribute Parameters_CKRinseTime.VB_VarDescription = "Category=Drugroom\r\nHelp=The time in seconds to rinse tank 1."
  Public Parameters_CKTransferToDrainTime2 As Long
Attribute Parameters_CKTransferToDrainTime2.VB_VarDescription = "Category=Drugroom\r\nHelp=The time in seconds to transfer tank 1 to drain after the rinse."

  Public Parameters_Tank1HighTempLimit As Long
Attribute Parameters_Tank1HighTempLimit.VB_VarDescription = "Category=Drugroom\r\nHelp=If the tank 1 temp. exceeds this number (in tenths).Tank 1 stops heating and alarm occurs."
  Public Parameters_Tank1DrainTime As Long
Attribute Parameters_Tank1DrainTime.VB_VarDescription = "Category=Drugroom\r\nHelp=Time, in seconds, tank 1 will transfer to the drain when tank 1 level reaches 0% after the transfer to drain is finished."
  Public Parameters_Tank1RinseTime As Long
Attribute Parameters_Tank1RinseTime.VB_VarDescription = "Category=Drugroom\r\nHelp=Time, in seconds, to rinse tank 1 to the machine after the transfer is finished."
  Public Parameters_Tank1RinseToDrainTime As Long
Attribute Parameters_Tank1RinseToDrainTime.VB_VarDescription = "Category=Drugroom\r\nHelp=Time, in seconds, to rinse tank 1 to drain after the transfer is finished."
  Public Parameters_Tank1TimeAfterRinse As Long
Attribute Parameters_Tank1TimeAfterRinse.VB_VarDescription = "Category=Drugroom\r\nHelp=Time, in seconds, to continue transferring to machine after the tank 1 rinse to machine is finished."
  Public Parameters_Tank1TimeBeforeRinse As Long
Attribute Parameters_Tank1TimeBeforeRinse.VB_VarDescription = "Category=Drugroom\r\nHelp=Time, in seconds, to continue transferring tank 1 to machine, once empty, before rinsing to machine."
  Public Parameters_DrugroomRinses As Long
Attribute Parameters_DrugroomRinses.VB_VarDescription = "Category=Drugroom\r\nHelp=Number of tank rinses during a drugroom transfer."
  Public Parameters_Tank1HeatDeadband As Long
Attribute Parameters_Tank1HeatDeadband.VB_VarDescription = "Category=Drugroom\r\nHelp=The temperature deadband, in tenths of a degree F, for the drugroom tank."
  Public Parameters_Tank1MixerOnLevel As Long
Attribute Parameters_Tank1MixerOnLevel.VB_VarDescription = "Category=Drugroom\r\nHelp=The level, in tenths of a percent, to turn on the mixer."
  Public Parameters_Tank1MixerOffLevel As Long
Attribute Parameters_Tank1MixerOffLevel.VB_VarDescription = "Category=Drugroom\r\nHelp=The level, in tenths of a percent, to turn off the mixer."
  Public Parameters_Tank1RinseLevel As Long
Attribute Parameters_Tank1RinseLevel.VB_VarDescription = "Category=Drugroom\r\nHelp=Level, in tenths of a percent, to fill tank 1 after the transfer is finished for the tank 1 rinsing to the machine."
  Public Parameters_Tank1FillLevelDeadBand As Long
Attribute Parameters_Tank1FillLevelDeadBand.VB_VarDescription = "Category=Drugroom\r\nHelp=The deadband level on a fill to the drugroom tank."
  Public Parameters_Tank1HeatPrepTimer As Long
Attribute Parameters_Tank1HeatPrepTimer.VB_VarDescription = "Category=Drugroom\r\nHelp=Time to fill and heat the drugroom tank before signalling that a failure may have occured resulting in taking too long to heat the drugroom tank.  Errors may be in mixer level too low and also fill valves defective."
  Public Parameters_Tank1PreFillLevel As Long
Attribute Parameters_Tank1PreFillLevel.VB_VarDescription = "Category=Drugroom\r\nHelp=Level, in tenths, to prefill the drugroom tank before dispensing."

'drugroom pots
  Public Parameters_Tank1TempPotMin As Long
Attribute Parameters_Tank1TempPotMin.VB_VarDescription = "Category=Drugroom Calibration\r\nHelp=The value that the controller reads from the temperature pot at the fully counter clockwise position. in tenths of a percent"
  Public Parameters_Tank1TempPotMax As Long
Attribute Parameters_Tank1TempPotMax.VB_VarDescription = "Category=Drugroom Calibration\r\nHelp=The value that the controller reads from the temperature pot at the fully clockwise position. in tenths of a percent"
  Public Parameters_Tank1TempPotRange As Long
Attribute Parameters_Tank1TempPotRange.VB_VarDescription = "Category=Drugroom Calibration\r\nHelp=The value for the fully clockwise position of the temperature pot.  in tenths of a degree F."
  Public Parameters_Tank1TimePotMin As Long
Attribute Parameters_Tank1TimePotMin.VB_VarDescription = "Category=Drugroom Calibration\r\nHelp=The value that the controller reads from the time pot at the fully counter clockwise position. in tenths of a percent"
  Public Parameters_Tank1TimePotMax As Long
Attribute Parameters_Tank1TimePotMax.VB_VarDescription = "Category=Drugroom Calibration\r\nHelp=The value that the controller reads from the time pot at the fully clockwise position. in tenths of a percent"
  Public Parameters_Tank1TimePotRange As Long
Attribute Parameters_Tank1TimePotRange.VB_VarDescription = "Category=Drugroom Calibration\r\nHelp=The value for the fully clockwise position of the time pot.  in minutes."
 
'Drain control
  Public Parameters_DRDrainTime As Long
Attribute Parameters_DRDrainTime.VB_VarDescription = "Category=Machine Drain\r\nHelp=The time in seconds to continue draining after expansion tank level is zero."
  Public Parameters_LevelGaugeFlushTime As Long
Attribute Parameters_LevelGaugeFlushTime.VB_VarDescription = "Category=Machine Drain\r\nHelp=The time in seconds to flush the  expansion tank differential pressure cell."
  Public Parameters_LevelGaugeSettleTime As Long
Attribute Parameters_LevelGaugeSettleTime.VB_VarDescription = "Category=Machine Drain\r\nHelp=The time in seconds to let the level settle after flushing the differential pressure cell."
  Public Parameters_HDDrainTime As Long
Attribute Parameters_HDDrainTime.VB_VarDescription = "Category=Machine Drain\r\nHelp=The time in seconds to continue draining after the drain level is made"
  Public Parameters_HDBlendFillPosition As Long
Attribute Parameters_HDBlendFillPosition.VB_VarDescription = "Category=Machine Drain\r\nHelp=The position of the blend fill valve, in tenths of a percent, during a hot drain when it fills to cool."
  Public Parameters_HDVentAlarmTime As Long
Attribute Parameters_HDVentAlarmTime.VB_VarDescription = "Category=Machine Drain\r\nHelp=If the vent valve is not open in this many seconds from the start of a hot drain. You will get an alarm."
  Public Parameters_HDHiFillLevel As Long
Attribute Parameters_HDHiFillLevel.VB_VarDescription = "Category=Machine Drain\r\nHelp=Stop filling to cool if this level is reached."
  Public Parameters_LevelFlushOverrideTime As Long
Attribute Parameters_LevelFlushOverrideTime.VB_VarDescription = "Category=Machine Drain\r\nHelp=Delay, in seconds, to wait when draining before automatically flushing the level gauge transmitter to confirm level gauge high side fully flooded.  Set to '0' to disable this feature.  During HD command, must have pressure sa"

'Fill control
  Public Parameters_FillSettleTime As Long
Attribute Parameters_FillSettleTime.VB_VarDescription = "Category=Fill Control\r\nHelp=The time to turn on the pump and let the level in the expansion tank settle before adding more water to the machine during a fill."
  Public Parameters_FillPrimePumpTime As Long
Attribute Parameters_FillPrimePumpTime.VB_VarDescription = "Category=Fill Control\r\nHelp=The time to turn the pump on to prime it."
  Public Parameters_RecordLevelTimer As Long
Attribute Parameters_RecordLevelTimer.VB_VarDescription = "Category=Fill Control\r\nHelp=The time in seconds after a fill to record the level. This recorded level will be used to check if the level is dropping in the cycle."
  Public Parameters_MaxLevelLossAmount As Long
Attribute Parameters_MaxLevelLossAmount.VB_VarDescription = "Category=Fill Control\r\nHelp=The amount of level in tenths of a percent that you can lose before getting an alarm."
  Public Parameters_FillOpenDelayTime As Long
Attribute Parameters_FillOpenDelayTime.VB_VarDescription = "Category=Fill Control\r\nHelp=Time, in seconds (max 10), to delay opening the fill valve.  Purpose to allow slower top wash valve to fully open to prevent pressure interlock."

'flow control
  Public Parameters_FLAdjustFactor As Long
Attribute Parameters_FLAdjustFactor.VB_VarDescription = "Category=Flow Rate Control\r\nHelp=During a FL command, the pump speed adjustment is proportional to this factor multiplied by the FL error."
  Public Parameters_FLAdjustMaximum As Long
Attribute Parameters_FLAdjustMaximum.VB_VarDescription = "Category=Flow Rate Control\r\nHelp=During a FL command, this is the maximum change allowed in pump speed, in tenths percent."
  Public Parameters_SystemVolumeAt0Level As Long
Attribute Parameters_SystemVolumeAt0Level.VB_VarDescription = "Category=Flow Rate Control\r\nHelp=Number of gallons of water in the vessel at zero level."
  Public Parameters_SystemVolumeAt100Level As Long
Attribute Parameters_SystemVolumeAt100Level.VB_VarDescription = "Category=Flow Rate Control\r\nHelp=Number of gallons of water in the vessel at 100% level."
  Public Parameters_FlowRateMin As Long
Attribute Parameters_FlowRateMin.VB_VarDescription = "Category=Flow Rate Control\r\nHelp=The value that the controller reads from the flow meter transmitter at 4ma. in tenths of a percent"
  Public Parameters_FlowRateMax As Long
Attribute Parameters_FlowRateMax.VB_VarDescription = "Category=Flow Rate Control\r\nHelp=The value that the controller reads from the flow meter transmitter at 20ma. in tenths of a percent"
  Public Parameters_FlowRateRange As Long
Attribute Parameters_FlowRateRange.VB_VarDescription = "Category=Flow Rate Control\r\nHelp=The flow rate equal to 20ma from the flow meter. in gallons per minute."
  Public Parameters_LowFlowRatePercent As Long
Attribute Parameters_LowFlowRatePercent.VB_VarDescription = "Category=Flow Rate Control\r\nHelp=If flow drops below this percent  for the low flow rate time give an alarm."
  Public Parameters_LowFlowRateTime As Long
Attribute Parameters_LowFlowRateTime.VB_VarDescription = "Category=Flow Rate Control\r\nHelp=The time for the flow rate."

'Level calibration
  Public Parameters_VesselLevelMin As Long
Attribute Parameters_VesselLevelMin.VB_VarDescription = "Category=Level Calibration\r\nHelp=The value that the controller reads from the vessel level transmitter at 4ma. in tenths of a percent"
  Public Parameters_VesselLevelMax As Long
Attribute Parameters_VesselLevelMax.VB_VarDescription = "Category=Level Calibration\r\nHelp=The value that the controller reads from the vessel level transmitter at 20ma. in tenths of a percent"
  Public Parameters_ReserveLevelMin As Long
Attribute Parameters_ReserveLevelMin.VB_VarDescription = "Category=Level Calibration\r\nHelp=The value that the controller reads from the reserve level transmitter at 4ma. in tenths of a percent"
  Public Parameters_ReserveLevelMax As Long
Attribute Parameters_ReserveLevelMax.VB_VarDescription = "Category=Level Calibration\r\nHelp=The value that the controller reads from the reserve level transmitter at 20ma. in tenths of a percent"
  Public Parameters_AdditionLevelMin As Long
Attribute Parameters_AdditionLevelMin.VB_VarDescription = "Category=Level Calibration\r\nHelp=The value that the controller reads from the addition level transmitter at 4ma. in tenths of a percent"
  Public Parameters_AdditionLevelMax As Long
Attribute Parameters_AdditionLevelMax.VB_VarDescription = "Category=Level Calibration\r\nHelp=The value that the controller reads from the addition level transmitter at 20ma. in tenths of a percent"
  Public Parameters_Tank1LevelMin As Long
Attribute Parameters_Tank1LevelMin.VB_VarDescription = "Category=Level Calibration\r\nHelp=The value that the controller reads from the tank 1 level transmitter at 4ma. in tenths of a percent"
  Public Parameters_Tank1LevelMax As Long
Attribute Parameters_Tank1LevelMax.VB_VarDescription = "Category=Level Calibration\r\nHelp=The value that the controller reads from the tank 1 level transmitter at 20ma. in tenths of a percent"

'lid control
  Public Parameters_LevelOkToOpenLid As Long
Attribute Parameters_LevelOkToOpenLid.VB_VarDescription = "Category=Lid Control\r\nHelp=The level in tenths of a percent that its ok to open the lid at."
  Public Parameters_LidRaisingTime As Long
Attribute Parameters_LidRaisingTime.VB_VarDescription = "Category=Lid Control\r\nHelp=The time in seconds to wait to ensure the lid is raised."
  Public Parameters_LockingBandOpeningTime As Long
Attribute Parameters_LockingBandOpeningTime.VB_VarDescription = "Category=Lid Control\r\nHelp=The time in seconds to wait to ensure the locking band is open."
  Public Parameters_LockingBandClosingTime As Long
Attribute Parameters_LockingBandClosingTime.VB_VarDescription = "Category=Lid Control\r\nHelp=The time in seconds to wait to ensure the locking band is closed."

'Pump control
  Public Parameters_FillSettlePumpSpeed As Long
Attribute Parameters_FillSettlePumpSpeed.VB_VarDescription = "Category=Pump Control\r\nHelp=The pump speed during a fill step. in tenths of a percent."
  Public Parameters_FlowReverseTime As Long
Attribute Parameters_FlowReverseTime.VB_VarDescription = "Category=Pump Control\r\nHelp=Do not adjust the pump speed for this many seconds after the flow reverse valve has switched with a DP or FL command active."
  Public Parameters_PumpAccelerationTime As Long
Attribute Parameters_PumpAccelerationTime.VB_VarDescription = "Category=Pump Control\r\nHelp=The acceleration time of the pump inverter."
  Public Parameters_PumpMinimumLevelTime As Long
Attribute Parameters_PumpMinimumLevelTime.VB_VarDescription = "Category=Pump Control\r\nHelp=The machine level must be above  ""Pump Minimum Level"" for this many seconds before the pump is allowed to start"
  Public Parameters_PumpMinimumLevel As Long
Attribute Parameters_PumpMinimumLevel.VB_VarDescription = "Category=Pump Control\r\nHelp=The machine level must be above this level for ""Pump Minimum Level Time"" seconds before the pump is allowed to start. in tenths of a percent."
  Public Parameters_PumpReversalSpeed As Long
Attribute Parameters_PumpReversalSpeed.VB_VarDescription = "Category=Pump Control\r\nHelp=The desired pump speed to reverse the valve at."
  Public Parameters_PumpSpeedAdjustTime As Long
Attribute Parameters_PumpSpeedAdjustTime.VB_VarDescription = "Category=Pump Control\r\nHelp=During a DP or FL command, the pump speed is adjusted after this many seconds."
  Public Parameters_PumpSpeedDecelTime As Long
Attribute Parameters_PumpSpeedDecelTime.VB_VarDescription = "Category=Pump Control\r\nHelp=The deceleration time of pump inverter. in seconds."
  Public Parameters_PumpSpeedDefault As Long
Attribute Parameters_PumpSpeedDefault.VB_VarDescription = "Category=Pump Control\r\nHelp=Default pump speed if no speed is programmed. In tenths of a percent."
  Public Parameters_PumpSpeedStart As Long
Attribute Parameters_PumpSpeedStart.VB_VarDescription = "Category=Pump Control\r\nHelp=The speed to start the pump at when a DP or FL command is active. in tenths of a percent."
  Public Parameters_RinsePumpSpeed As Long
Attribute Parameters_RinsePumpSpeed.VB_VarDescription = "Category=Pump Control\r\nHelp=The speed to run the pump at during a rinse. in tenths of a percent."
  Public Parameters_DrainPumpSpeed As Long
Attribute Parameters_DrainPumpSpeed.VB_VarDescription = "Category=Pump Control\r\nHelp=The speed to run the pump at during a drain. in tenths of a percent."
  Public Parameters_PumpOnOffCountTooHigh As Long
Attribute Parameters_PumpOnOffCountTooHigh.VB_VarDescription = "Category=Pump Control\r\nHelp=If the pump turns on and off more than this mainy times. Then the pump is disable from running and you get an alarm."

'Pressure Calibration
  Public Parameters_PackageDiffPresMin As Long
Attribute Parameters_PackageDiffPresMin.VB_VarDescription = "Category=Pressure Calibration\r\nHelp=The value that the controller reads from the differential pressure transmitter at 4ma. in tenths of a percent"
  Public Parameters_PackageDiffPresMax As Long
Attribute Parameters_PackageDiffPresMax.VB_VarDescription = "Category=Pressure Calibration\r\nHelp=The value that the controller reads from the differential pressure transmitter at 20ma. in tenths of a percent"
  Public Parameters_PackageDiffPresRange As Long
Attribute Parameters_PackageDiffPresRange.VB_VarDescription = "Category=Pressure Calibration\r\nHelp=The pressure equal to 20ma from the differential pressure transducer. in tenths of a psi."
  Public Parameters_VesselPressureMin As Long
Attribute Parameters_VesselPressureMin.VB_VarDescription = "Category=Pressure Calibration\r\nHelp=The value that the controller reads from the vessel pressure transmitter at 4ma. in tenths of a percent"
  Public Parameters_VesselPressureMax As Long
Attribute Parameters_VesselPressureMax.VB_VarDescription = "Category=Pressure Calibration\r\nHelp=The value that the controller reads from the vessel pressure transmitter at max input 20ma. in tenths of a percent"
  Public Parameters_VesselPressureRange As Long
Attribute Parameters_VesselPressureRange.VB_VarDescription = "Category=Pressure Calibration\r\nHelp=The pressure equal to 20ma from the vessel pressure transducer. in tenths of a psi."
  
'production reports
  Public Parameters_StandardOperatorTime As Long
Attribute Parameters_StandardOperatorTime.VB_VarDescription = "Category=Production Reports\r\nHelp=If operator takes longer than this number of seconds for a call. Then the delay is logged as a operator delay."
  Public Parameters_StandardLocalPrepTime As Long
Attribute Parameters_StandardLocalPrepTime.VB_VarDescription = "Category=Production Reports\r\nHelp=If operator takes longer than this number of seconds for a local tank prepare command. Then the delay is logged as a operator delay."
  Public Parameters_StandardReserveTransferTime As Long
Attribute Parameters_StandardReserveTransferTime.VB_VarDescription = "Category=Production Reports\r\nHelp=Time, in minutes, to transfer reserve tank contents once tank ready state is set before signalling a reserve transfer delay overrun."
  Public Parameters_StandardLoadTime As Long
Attribute Parameters_StandardLoadTime.VB_VarDescription = "Category=Production Reports\r\nHelp=Time, in minutes, to load a machine and signal ready."
  Public Parameters_StandardUnloadTime As Long
Attribute Parameters_StandardUnloadTime.VB_VarDescription = "Category=Production Reports\r\nHelp=Time, in minutes, to unload a machine and signal ready."
  Public Parameters_SevenDaySchedule As Long              'requested by allen michael 2009/06/15
Attribute Parameters_SevenDaySchedule.VB_VarDescription = "Category=Production Reports\r\nHelp=Set to '1', to enable seven day scheduling (full work week), else Load delays will be disregarded on sunday startup."
  
'Look Ahead
  Public Parameters_LookAheadEnabled As Long
Attribute Parameters_LookAheadEnabled.VB_VarDescription = "Category=Look Ahead\r\nHelp=When set to one enables the look ahead function."
  Public Parameters_LookAheadIgnoreBlocked As Long
Attribute Parameters_LookAheadIgnoreBlocked.VB_VarDescription = "Category=Look Ahead\r\nHelp=When set to one the look ahead function will prepare the next scheduled lot even if it is blocked."
  Public Parameters_EnableRedyeIssueAlarm As Long
Attribute Parameters_EnableRedyeIssueAlarm.VB_VarDescription = "Category=Look Ahead\r\nHelp=Set to '1' to enable KP testing LA's dyelot and redye values to determine if current batch has a redye issue to signal tank check alarm."

'reserve tank
  Public Parameters_ReserveTankHighTempLimit As Long
Attribute Parameters_ReserveTankHighTempLimit.VB_VarDescription = "Category=Reserve Tank Control\r\nHelp=High Temp Limit, in tenths degree, for reserve tank"
  Public Parameters_ReserveDrainTime As Long
Attribute Parameters_ReserveDrainTime.VB_VarDescription = "Category=Reserve Tank Control\r\nHelp=The time in seconds to drain the reserve tank at the end of a transfer."
  Public Parameters_ReserveHeatDeadband As Long
Attribute Parameters_ReserveHeatDeadband.VB_VarDescription = "Category=Reserve Tank Control\r\nHelp=The temperature deadband, in tenths of a degree F, for the reserve tank."
  Public Parameters_ReserveMixerOnLevel As Long
Attribute Parameters_ReserveMixerOnLevel.VB_VarDescription = "Category=Reserve Tank Control\r\nHelp=The level, in tenths of a percent, to turn on the mixer."
  Public Parameters_ReserveMixerOffLevel As Long
Attribute Parameters_ReserveMixerOffLevel.VB_VarDescription = "Category=Reserve Tank Control\r\nHelp=The level, in tenths of a percent, to turn off the mixer."
  Public Parameters_ReserveRinseTime As Long
Attribute Parameters_ReserveRinseTime.VB_VarDescription = "Category=Reserve Tank Control\r\nHelp=The time, in seconds, to rinse the reserve tank to the vessel."
  Public Parameters_ReserveRinseToDrainTime As Long
Attribute Parameters_ReserveRinseToDrainTime.VB_VarDescription = "Category=Reserve Tank Control\r\nHelp=The time, in seconds, to rinse the reserve tank to dran."
  Public Parameters_ReserveTimeAfterRinse As Long
Attribute Parameters_ReserveTimeAfterRinse.VB_VarDescription = "Category=Reserve Tank Control\r\nHelp=The time, in seconds, to transfer the reserve tank to the vessel after rinsing."
  Public Parameters_ReserveTimeBeforeRinse As Long
Attribute Parameters_ReserveTimeBeforeRinse.VB_VarDescription = "Category=Reserve Tank Control\r\nHelp=The time, in seconds, to transfer the reserve tank to the vessel before rinsing."
  Public Parameters_ReserveBackfillLevel As Long
Attribute Parameters_ReserveBackfillLevel.VB_VarDescription = "Category=Reserve Tank Control\r\nHelp=Level, in tenths, to backfill the reserve to on RF command."
  Public Parameters_ReserveBackFillStartTime As Long
Attribute Parameters_ReserveBackFillStartTime.VB_VarDescription = "Category=Reserve Tank Control\r\nHelp=The time, in seconds, used during Reserve Fill to open the reserve backfill for this many seconds before activating the airpad."
  
'rinse temperature band
  Public Parameters_RinseTemperatureAlarmBand As Long
Attribute Parameters_RinseTemperatureAlarmBand.VB_VarDescription = "Category=Rinse Control\r\nHelp=The temperature alarm band during a rinse, in tenths."
  Public Parameters_RinseTemperatureAlarmDelay As Long
Attribute Parameters_RinseTemperatureAlarmDelay.VB_VarDescription = "Category=Rinse Control\r\nHelp=The temperature alarm delay time, in seconds."
  Public Parameters_TopWashBlockageTimer As Long
Attribute Parameters_TopWashBlockageTimer.VB_VarDescription = "Category=Rinse Control\r\nHelp=Time, in seconds, to lose PressureInterlock input while rinsing before signalling TopWashBlocked Alarm.  Disable by setting this parameter value to '0'"

'Working Level Calculations
  Public Parameters_NumberOfSpindales As Long
Attribute Parameters_NumberOfSpindales.VB_VarDescription = "Category=Working Level\r\nHelp=Set to number of spindales available on machine.  Used to determine working level from package height = Number Packages / Number Of Spindales."
  Public Parameters_EnableFillByWorkingLevel As Long
Attribute Parameters_EnableFillByWorkingLevel.VB_VarDescription = "Category=Working Level\r\nHelp=Set to '1' to Enable the FI command to use the WorkingLevel based on package type and package number instead of the FI command specified fill level."
  Public Parameters_EnableReserveFillByWorkingLevel As Long
Attribute Parameters_EnableReserveFillByWorkingLevel.VB_VarDescription = "Category=Working Level\r\nHelp=Set to '1' to Enable the RF command to use the WorkingLevel based on package type and package number instead of the FI command specified fill level."
  
  Public Parameters_PackageType0Level1 As Long
Attribute Parameters_PackageType0Level1.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type 0 - Working Level, in tenths percent, for package height of 1."
  Public Parameters_PackageType0Level2 As Long
Attribute Parameters_PackageType0Level2.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type 0 - Working Level, in tenths percent, for package height of 2."
  Public Parameters_PackageType0Level3 As Long
Attribute Parameters_PackageType0Level3.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type 0 - Working Level, in tenths percent, for package height of 3."
  Public Parameters_PackageType0Level4 As Long
Attribute Parameters_PackageType0Level4.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type 0 - Working Level, in tenths percent, for package height of 4."
  Public Parameters_PackageType0Level5 As Long
Attribute Parameters_PackageType0Level5.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type 0 - Working Level, in tenths percent, for package height of 5."
  Public Parameters_PackageType0Level6 As Long
Attribute Parameters_PackageType0Level6.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type 0 - Working Level, in tenths percent, for package height of 6."
  Public Parameters_PackageType0Level7 As Long
Attribute Parameters_PackageType0Level7.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type 0 - Working Level, in tenths percent, for package height of 7."
  Public Parameters_PackageType0Level8 As Long
Attribute Parameters_PackageType0Level8.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type 0 - Working Level, in tenths percent, for package height of 8."
  Public Parameters_PackageType0Level9 As Long
Attribute Parameters_PackageType0Level9.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type 0 - Working Level, in tenths percent, for package height of 9."
  
  Public Parameters_PackageType1Level1 As Long
Attribute Parameters_PackageType1Level1.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type 1 (A) - Working Level, in tenths percent, for package height of 1."
  Public Parameters_PackageType1Level2 As Long
Attribute Parameters_PackageType1Level2.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type 1 (A) - Working Level, in tenths percent, for package height of 2."
  Public Parameters_PackageType1Level3 As Long
Attribute Parameters_PackageType1Level3.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type 1 (A) - Working Level, in tenths percent, for package height of 3."
  Public Parameters_PackageType1Level4 As Long
Attribute Parameters_PackageType1Level4.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type 1 (A) - Working Level, in tenths percent, for package height of 4."
  Public Parameters_PackageType1Level5 As Long
Attribute Parameters_PackageType1Level5.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type 1 (A) - Working Level, in tenths percent, for package height of 5."
  Public Parameters_PackageType1Level6 As Long
Attribute Parameters_PackageType1Level6.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type 1 (A) - Working Level, in tenths percent, for package height of 6."
  Public Parameters_PackageType1Level7 As Long
Attribute Parameters_PackageType1Level7.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type 1 (A) - Working Level, in tenths percent, for package height of 7."
  Public Parameters_PackageType1Level8 As Long
Attribute Parameters_PackageType1Level8.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type 1 (A) - Working Level, in tenths percent, for package height of 8."
  Public Parameters_PackageType1Level9 As Long
Attribute Parameters_PackageType1Level9.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type 1 (A) - Working Level, in tenths percent, for package height of 9."
  
  Public Parameters_PackageType2Level1 As Long
Attribute Parameters_PackageType2Level1.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type B - Working Level, in tenths percent, for package height of 1."
  Public Parameters_PackageType2Level2 As Long
Attribute Parameters_PackageType2Level2.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type B - Working Level, in tenths percent, for package height of 2."
  Public Parameters_PackageType2Level3 As Long
Attribute Parameters_PackageType2Level3.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type B - Working Level, in tenths percent, for package height of 3."
  Public Parameters_PackageType2Level4 As Long
Attribute Parameters_PackageType2Level4.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type B - Working Level, in tenths percent, for package height of 4."
  Public Parameters_PackageType2Level5 As Long
Attribute Parameters_PackageType2Level5.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type B - Working Level, in tenths percent, for package height of 5."
  Public Parameters_PackageType2Level6 As Long
Attribute Parameters_PackageType2Level6.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type B - Working Level, in tenths percent, for package height of 6."
  Public Parameters_PackageType2Level7 As Long
Attribute Parameters_PackageType2Level7.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type B - Working Level, in tenths percent, for package height of 7."
  Public Parameters_PackageType2Level8 As Long
Attribute Parameters_PackageType2Level8.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type B - Working Level, in tenths percent, for package height of 8."
  Public Parameters_PackageType2Level9 As Long
Attribute Parameters_PackageType2Level9.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type B - Working Level, in tenths percent, for package height of 9."
  
  Public Parameters_PackageType3Level1 As Long
Attribute Parameters_PackageType3Level1.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type C - Working Level, in tenths percent, for package height of 1."
  Public Parameters_PackageType3Level2 As Long
Attribute Parameters_PackageType3Level2.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type C - Working Level, in tenths percent, for package height of 2."
  Public Parameters_PackageType3Level3 As Long
Attribute Parameters_PackageType3Level3.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type C - Working Level, in tenths percent, for package height of 3."
  Public Parameters_PackageType3Level4 As Long
Attribute Parameters_PackageType3Level4.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type C - Working Level, in tenths percent, for package height of 4."
  Public Parameters_PackageType3Level5 As Long
Attribute Parameters_PackageType3Level5.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type C - Working Level, in tenths percent, for package height of 5."
  Public Parameters_PackageType3Level6 As Long
Attribute Parameters_PackageType3Level6.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type C - Working Level, in tenths percent, for package height of 6."
  Public Parameters_PackageType3Level7 As Long
Attribute Parameters_PackageType3Level7.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type C - Working Level, in tenths percent, for package height of 7."
  Public Parameters_PackageType3Level8 As Long
Attribute Parameters_PackageType3Level8.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type C - Working Level, in tenths percent, for package height of 8."
  Public Parameters_PackageType3Level9 As Long
Attribute Parameters_PackageType3Level9.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type C - Working Level, in tenths percent, for package height of 9."
  
  Public Parameters_PackageType4Level1 As Long
Attribute Parameters_PackageType4Level1.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type D - Working Level, in tenths percent, for package height of 1."
  Public Parameters_PackageType4Level2 As Long
Attribute Parameters_PackageType4Level2.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type D - Working Level, in tenths percent, for package height of 2."
  Public Parameters_PackageType4Level3 As Long
Attribute Parameters_PackageType4Level3.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type D - Working Level, in tenths percent, for package height of 3."
  Public Parameters_PackageType4Level4 As Long
Attribute Parameters_PackageType4Level4.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type D - Working Level, in tenths percent, for package height of 4."
  Public Parameters_PackageType4Level5 As Long
Attribute Parameters_PackageType4Level5.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type D - Working Level, in tenths percent, for package height of 5."
  Public Parameters_PackageType4Level6 As Long
Attribute Parameters_PackageType4Level6.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type D - Working Level, in tenths percent, for package height of 6."
  Public Parameters_PackageType4Level7 As Long
Attribute Parameters_PackageType4Level7.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type D - Working Level, in tenths percent, for package height of 7."
  Public Parameters_PackageType4Level8 As Long
Attribute Parameters_PackageType4Level8.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type D - Working Level, in tenths percent, for package height of 8."
  Public Parameters_PackageType4Level9 As Long
Attribute Parameters_PackageType4Level9.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type D - Working Level, in tenths percent, for package height of 9."

  Public Parameters_PackageType5Level1 As Long
Attribute Parameters_PackageType5Level1.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type E - Working Level, in tenths percent, for package height of 1."
  Public Parameters_PackageType5Level2 As Long
Attribute Parameters_PackageType5Level2.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type E - Working Level, in tenths percent, for package height of 2."
  Public Parameters_PackageType5Level3 As Long
Attribute Parameters_PackageType5Level3.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type E - Working Level, in tenths percent, for package height of 3."
  Public Parameters_PackageType5Level4 As Long
Attribute Parameters_PackageType5Level4.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type E - Working Level, in tenths percent, for package height of 4."
  Public Parameters_PackageType5Level5 As Long
Attribute Parameters_PackageType5Level5.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type E - Working Level, in tenths percent, for package height of 5."
  Public Parameters_PackageType5Level6 As Long
Attribute Parameters_PackageType5Level6.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type E - Working Level, in tenths percent, for package height of 6."
  Public Parameters_PackageType5Level7 As Long
Attribute Parameters_PackageType5Level7.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type E - Working Level, in tenths percent, for package height of 7."
  Public Parameters_PackageType5Level8 As Long
Attribute Parameters_PackageType5Level8.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type E - Working Level, in tenths percent, for package height of 8."
  Public Parameters_PackageType5Level9 As Long
Attribute Parameters_PackageType5Level9.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type E - Working Level, in tenths percent, for package height of 9."
  
  Public Parameters_PackageType6Level1 As Long
Attribute Parameters_PackageType6Level1.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type F - Working Level, in tenths percent, for package height of 1."
  Public Parameters_PackageType6Level2 As Long
Attribute Parameters_PackageType6Level2.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type F - Working Level, in tenths percent, for package height of 2."
  Public Parameters_PackageType6Level3 As Long
Attribute Parameters_PackageType6Level3.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type F - Working Level, in tenths percent, for package height of 3."
  Public Parameters_PackageType6Level4 As Long
Attribute Parameters_PackageType6Level4.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type F - Working Level, in tenths percent, for package height of 4."
  Public Parameters_PackageType6Level5 As Long
Attribute Parameters_PackageType6Level5.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type F - Working Level, in tenths percent, for package height of 5."
  Public Parameters_PackageType6Level6 As Long
Attribute Parameters_PackageType6Level6.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type F - Working Level, in tenths percent, for package height of 6."
  Public Parameters_PackageType6Level7 As Long
Attribute Parameters_PackageType6Level7.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type F - Working Level, in tenths percent, for package height of 7."
  Public Parameters_PackageType6Level8 As Long
Attribute Parameters_PackageType6Level8.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type F - Working Level, in tenths percent, for package height of 8."
  Public Parameters_PackageType6Level9 As Long
Attribute Parameters_PackageType6Level9.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type F - Working Level, in tenths percent, for package height of 9."
  
  Public Parameters_PackageType7Level1 As Long
Attribute Parameters_PackageType7Level1.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type G - Working Level, in tenths percent, for package height of 1."
  Public Parameters_PackageType7Level2 As Long
Attribute Parameters_PackageType7Level2.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type G - Working Level, in tenths percent, for package height of 2."
  Public Parameters_PackageType7Level3 As Long
Attribute Parameters_PackageType7Level3.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type G - Working Level, in tenths percent, for package height of 3."
  Public Parameters_PackageType7Level4 As Long
Attribute Parameters_PackageType7Level4.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type G - Working Level, in tenths percent, for package height of 4."
  Public Parameters_PackageType7Level5 As Long
Attribute Parameters_PackageType7Level5.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type G - Working Level, in tenths percent, for package height of 5."
  Public Parameters_PackageType7Level6 As Long
Attribute Parameters_PackageType7Level6.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type G - Working Level, in tenths percent, for package height of 6."
  Public Parameters_PackageType7Level7 As Long
Attribute Parameters_PackageType7Level7.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type G - Working Level, in tenths percent, for package height of 7."
  Public Parameters_PackageType7Level8 As Long
Attribute Parameters_PackageType7Level8.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type G - Working Level, in tenths percent, for package height of 8."
  Public Parameters_PackageType7Level9 As Long
Attribute Parameters_PackageType7Level9.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type G - Working Level, in tenths percent, for package height of 9."

  Public Parameters_PackageType8Level1 As Long
Attribute Parameters_PackageType8Level1.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type H - Working Level, in tenths percent, for package height of 1."
  Public Parameters_PackageType8Level2 As Long
Attribute Parameters_PackageType8Level2.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type H - Working Level, in tenths percent, for package height of 2."
  Public Parameters_PackageType8Level3 As Long
Attribute Parameters_PackageType8Level3.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type H - Working Level, in tenths percent, for package height of 3."
  Public Parameters_PackageType8Level4 As Long
Attribute Parameters_PackageType8Level4.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type H - Working Level, in tenths percent, for package height of 4."
  Public Parameters_PackageType8Level5 As Long
Attribute Parameters_PackageType8Level5.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type H - Working Level, in tenths percent, for package height of 5."
  Public Parameters_PackageType8Level6 As Long
Attribute Parameters_PackageType8Level6.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type H - Working Level, in tenths percent, for package height of 6."
  Public Parameters_PackageType8Level7 As Long
Attribute Parameters_PackageType8Level7.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type H - Working Level, in tenths percent, for package height of 7."
  Public Parameters_PackageType8Level8 As Long
Attribute Parameters_PackageType8Level8.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type H - Working Level, in tenths percent, for package height of 8."
  Public Parameters_PackageType8Level9 As Long
Attribute Parameters_PackageType8Level9.VB_VarDescription = "Category=Working Level\r\nHelp=Package Type H - Working Level, in tenths percent, for package height of 9."
  
'===============================================================================================
'COMMANDS - Make a public instance of each command
'===============================================================================================
  Public AC As New AC: Public AD As New AD: Public AF As New AF
  Public AP As New AP: Public AT As New AT: Public BO As New BO
  Public BR As New BR: Public BW As New BW: Public CC As New CC
  Public CK As New CK: Public CO As New CO: Public DP As New DP
  Public DR As New DR: Public ET As New ET: Public FC As New FC
  Public FI As New FI: Public FL As New FL: Public FP As New FP
  Public FR As New FR: Public HC As New HC: Public HD As New HD
  Public HE As New HE: Public KA As New KA: Public KP As New KP
  Public KR As New KR: Public LA As New LA: Public LD As New LD
  Public LS As New LS: Public NP As New NP: Public PH As New PH
  Public PR As New PR: Public PT As New PT: Public RC As New RC
  Public RD As New RD: Public RF As New RF: Public RH As New RH
  Public RI As New RI: Public RP As New RP: Public RT As New RT
  Public RW As New RW: Public SA As New SA: Public TC As New TC
  Public TM As New TM: Public TP As New TP: Public UL As New UL
  Public WK As New WK: Public WT As New WT

'===============================================================================================
'ALARMS - Put all alarms here, unless we can put them in a specific class module instead
'===============================================================================================
  Public Alarms_ReserveTempProbeError As Boolean
  Public Alarms_BlendFillTempProbeError As Boolean
  'Public Alarms_AddTankTempProbeError As Boolean
  Public Alarms_ControlPowerFailure As Boolean
  Public Alarms_DebugModeSelected As Boolean
  Public Alarms_EmergencyStopPB As Boolean
  Public Alarms_LidNotLocked As Boolean
  Public Alarms_OverrideModeSelected As Boolean
  Public Alarms_MainPLCNotResponding As Boolean
  Public Alarms_DrugroomPLCNotResponding As Boolean
  Public Alarms_MainPumpTripped As Boolean
  Public Alarms_AddPumpTripped As Boolean
  Public Alarms_VesselTempProbeError As Boolean
  Public Alarms_TemperatureTooHigh As Boolean
  Public Alarms_TestModeSelected As Boolean
  Public Alarms_LevelToLowForPump As Boolean
  Public Alarms_VentNotOpenOnHotDrain As Boolean
  Public Alarms_LidNotClosed As Boolean
  Public Alarms_LevelTooHighInAddTank As Boolean
  Public Alarms_LosingLevelInTheVessel As Boolean
  Public Alarms_CheckLevelTransmitter As Boolean
  Public Alarms_MainPumpTurnedOnOffTooManyTimes As Boolean
  Public Alarms_ReserveTempTooHigh As Boolean
  Public Alarms_TemperatureHigh As Boolean
  Public Alarms_TemperatureLow As Boolean
  Public Alarms_LowFlowRate As Boolean
  Public Alarms_Tank1TempProbeError As Boolean
  Public Alarms_Tank1TempTooHigh As Boolean
  Public Alarms_Tank1DispenseError As Boolean
  Public Alarms_Tank1TooLongToHeat As Boolean
  Public Alarms_DispenserCompleteDelay As Boolean
  Public Alarms_AutoDispenserDelay As Boolean
  Public Alarms_TopWashBlocked As Boolean
  Public Alarms_TooLongToHeat As Boolean
  Public Alarms_RedyeIssueCheckMachineTank As Boolean
  
  Public Alarms_CondensateTempTooHigh As Boolean
  Public Alarms_TestValueTooHigh As Boolean
'
'===============================================================================================
' CONTROL VARIABLES
'===============================================================================================
  
  Public FirstScanDone As Boolean

'sql string for setting committed field in database useing the mimic code.
  Public LASqlString As String
  Public LASetSqlString As Boolean

'booleans for signal stuff
  Public SignalRequested As Boolean
  Public SignalRequestedWas As Boolean
  Public SignalOnRequest As Boolean

'For Production Reports
  Public Utilities_Timer As New acTimer
  Public OperatorDelayTimer As New acTimer
  Public PowerKWS As Long
  Public PowerKWH As Long
  Public SteamUsed As Long
  Private MainPumpHP As Long
  Public SteamNeeded As Long
  Public TempRise As Long
  Private FinalTempWas As Long
  Private FinalTemp As Long
  Private StartTemp As Long
  
'Delay variable values are
  Public Delay As Long
  Public Delay_IsDelayed As Boolean
    
  Public Delay_NormalRunning As Boolean                 '0
  Public Delay_Boilout As Boolean                       '1
  Public Delay_Correction As Boolean                    '2
  'Public Delay_StepOverrun as boolean                  '3 not set in this code
  Public Delay_Sample As Boolean                        '4
  Public Delay_Load As Boolean                          '5
  Public Delay_Unload As Boolean                        '6
  Public Delay_CheckPH As Boolean                       '7
  Public Delay_Operator As Boolean                      '8
  Public Delay_Fill As Boolean                          '9
  Public Delay_Drain As Boolean                         '10
  Public Delay_Heating As Boolean                       '11
  Public Delay_Cooling As Boolean                       '12
  Public Delay_Paused As Boolean                        '13
  Public Delay_MachineError As Boolean                  '14
  Public Delay_ExpansionAdd As Boolean                  '15
  Public Delay_Tank1Prepare As Boolean                  '16
  Public Delay_ReserveTransfer As Boolean               '17
  Public Delay_DispenseWaitReady As Boolean             '18
  Public Delay_DispenseWaitResultDyes As Boolean        '19
  Public Delay_DispenseWaitResultChems As Boolean       '20
  Public Delay_DispenseWaitResultCombined As Boolean    '21
  Public Delay_LocalPrepare As Boolean                  '22
  Public Delay_Tank1Operator As Boolean                 '23
  Public Delay_Tank1DispenseError As Boolean            '24
  Public Delay_Sleeping As Boolean                      '101
     
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
  
  'sleep for modbus coms
  Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)
  
Public Sub ACControlCode_Run()

' Generate fast and slow flash - these will be used in a few places
  FastFlasher.Flash FastFlash, 400
  SlowFlasher.Flash SlowFlash, 800

' time to go stuff
  TimeToGo = 0
  If AC.IsOn Then TimeToGo = AC.Timer
  If AT.IsOn Then TimeToGo = AT.Timer
  If DR.IsOn Then TimeToGo = DR.Timer
  If FI.IsOn Then TimeToGo = FI.FITimer
  If RI.IsOn Then TimeToGo = RI.Timer
  If RH.IsOn Then TimeToGo = RH.Timer
  If RT.IsOn Then TimeToGo = RT.Timer
  If TM.IsOn Then TimeToGo = TM.Timer
  If HD.IsOn Then TimeToGo = HD.Timer
  If RW.IsActive Then TimeToGo = RW.Timer
  
  If Not FirstScanDone Then ProgramStoppedTimer.Start

'Parameter Initialization (if a magic number is found in Parameters_InitializeParameters)
  Parameters_MaybeSetDefaults
  
'Extract time in step
  If Parent.IsProgramRunning Then
    TimeInStepValue = TISValue(0)
    StepStandardTime = TISValue(1)
  Else
    TimeInStepValue = 0
    StepStandardTime = 0
  End If
  
  Static DisplayTIS As Boolean
  
  If TimeInStepValue > StepStandardTime Then
    If TwoSecondTimer.Finished Then
      DisplayTIS = Not DisplayTIS
      TwoSecondTimer = 2
    End If
  Else
    DisplayTIS = True
  End If
  If DisplayTIS Then
    TimeInStep = CStr(TimeInStepValue)
  Else
    TimeInStep = "Overrun"
  End If

'===============================================================================================
'PROGRAM STOP/START
'===============================================================================================
'
'Program running state changes
  Static ProgramRunTimer As New acTimerUp  ' program run timer
  Static ProgramWasRunning As Boolean
  
  If Parent.IsProgramRunning Then            'A Program is running
    CycleTime = ProgramRunTimer
    CycleTimeDisplay = TimerString(CycleTime)
    If Not ProgramWasRunning Then           'A Program has just started
      ProgramStoppedTime = ProgramStoppedTimer.TimeElapsed
      ProgramStoppedTimer.Pause
      ProgramRunTimer.Start
    End If
  Else
    If ProgramWasRunning Then                'A program has just finished
      ProgramStoppedTimer.Start
      LastProgramCycleTime = CycleTime
      PumpRequest = False
      TemperatureControl.Cancel
      TemperatureControlContacts.Cancel
    End If
    If Not LA.KP1.IsOn Then Tank1Ready = False
    GetRecordedLevel = False
    Alarms_LosingLevelInTheVessel = False
    CycleTime = 0
    AddReadyFromPB = False
    ReserveReadyFromPB = False
    CycleTimeDisplay = "0"
    PowerKWS = 0
    PowerKWH = 0
    SteamUsed = 0
    SystemVolume = 0
    NumberOfContacts = 0
    TotalNumberOfContacts = 0
    PumpOnCount = 0
    WorkingLevel = 0
    PackageType = 0
    PackageHeight = 0
  End If
  ProgramWasRunning = Parent.IsProgramRunning
'
'===============================================================================================
'END PROGRAM STOP/START
'===============================================================================================

'Dispenser & Drugroom Preview Determination
'if a program is running look to see if the next drop has a manual.
  If Parent.IsProgramRunning Then
     'for dispenser
     If Not (KP.KP1.IsOn Or LA.KP1.IsOn Or KA.IsOn) Then
       DrugroomDisplayJob1 = Parent.Job
     End If
     If KP.KP1.IsOn Then
        If KP.KP1.AddDispenseError And Not Tank1Ready Then
           DrugroomPreview.Tank1Calloff = KP.KP1.AddCallOff
           DrugroomPreview.DestinationTank = KP.KP1.DispenseTank
            If DrugroomCalloffRefreshed = False Then DrugroomPreview.GetTimeBeforeTransfer Me, DrugroomPreview.DestinationTank
            DrugroomCalloffRefreshed = True
           DrugroomPreview.PreviewStatus = "Dispense error for calloff " & DrugroomPreview.Tank1Calloff
        ElseIf KP.KP1.ManualAdd And Not Tank1Ready Then
           DrugroomPreview.Tank1Calloff = KP.KP1.AddCallOff
           DrugroomPreview.DestinationTank = KP.KP1.DispenseTank
            If DrugroomCalloffRefreshed = False Then DrugroomPreview.GetTimeBeforeTransfer Me, DrugroomPreview.DestinationTank
            DrugroomCalloffRefreshed = True
           DrugroomPreview.PreviewStatus = "Manual add for calloff " & DrugroomPreview.Tank1Calloff
        ElseIf DrugroomPreview.IsManualProduct(KP.KP1.AddCallOff, Parameters_DDSEnabled, Parameters_DispenseEnabled) And Not Tank1Ready Then
           DrugroomPreview.DestinationTank = KP.KP1.DispenseTank
           If DrugroomCalloffRefreshed = False Then DrugroomPreview.GetTimeBeforeTransfer Me, DrugroomPreview.DestinationTank
           DrugroomCalloffRefreshed = True
           If DrugroomPreview.CurrentRecipe Then
              DrugroomPreview.PreviewStatus = "Manual add for calloff " & DrugroomPreview.Tank1Calloff
           Else
              DrugroomPreview.PreviewStatus = "No Recipe manual add for calloff " & DrugroomPreview.Tank1Calloff
           End If
        Else
           DrugroomPreview.DestinationTank = 0
           DrugroomPreview.PreviewStatus = ""
           DrugroomPreview.Tank1Calloff = 0
           DrugroomPreview.TimeToTransferTimer.TimeRemaining = 0
        End If
     ElseIf LA.KP1.IsOn Then
        If LA.KP1.AddDispenseError And Not Tank1Ready Then
           DrugroomPreview.Tank1Calloff = LA.KP1.AddCallOff
           DrugroomPreview.DestinationTank = LA.KP1.DispenseTank
            If DrugroomCalloffRefreshed = False Then DrugroomPreview.GetTimeBeforeTransfer Me, DrugroomPreview.DestinationTank
            DrugroomCalloffRefreshed = True
           DrugroomPreview.PreviewStatus = "Dispense error for calloff " & DrugroomPreview.Tank1Calloff
        ElseIf LA.KP1.ManualAdd And Not Tank1Ready Then
           DrugroomPreview.Tank1Calloff = LA.KP1.AddCallOff
           DrugroomPreview.DestinationTank = LA.KP1.DispenseTank
            If DrugroomCalloffRefreshed = False Then DrugroomPreview.GetTimeBeforeTransfer Me, DrugroomPreview.DestinationTank
            DrugroomCalloffRefreshed = True
           DrugroomPreview.PreviewStatus = "Manual add for calloff " & DrugroomPreview.Tank1Calloff
        ElseIf DrugroomPreview.IsManualProduct(LA.KP1.AddCallOff, Parameters_DDSEnabled, Parameters_DispenseEnabled) And Not Tank1Ready Then
           DrugroomPreview.DestinationTank = LA.KP1.DispenseTank
           If DrugroomCalloffRefreshed = False Then DrugroomPreview.GetTimeBeforeTransfer Me, DrugroomPreview.DestinationTank
           DrugroomCalloffRefreshed = True
           If DrugroomPreview.CurrentRecipe Or DrugroomPreview.NextRecipe Then
              DrugroomPreview.PreviewStatus = "Manual add for calloff " & DrugroomPreview.Tank1Calloff
           Else
              DrugroomPreview.PreviewStatus = "No Recipe manual add for calloff " & DrugroomPreview.Tank1Calloff
           End If
        Else
           DrugroomPreview.DestinationTank = 0
           DrugroomPreview.PreviewStatus = ""
           DrugroomPreview.Tank1Calloff = 0
           DrugroomPreview.TimeToTransferTimer.TimeRemaining = 0
        End If
     ElseIf LAActive Then
           DrugroomPreview.DestinationTank = 0
           DrugroomPreview.PreviewStatus = ""
           DrugroomPreview.Tank1Calloff = 0
           DrugroomPreview.TimeToTransferTimer.TimeRemaining = 0
     ElseIf Not DrugroomCalloffRefreshed Then
        If DrugroomPreview.IsManualProduct(GetKPCalloff(Me), Parameters_DDSEnabled, Parameters_DispenseEnabled) And (Not LAActive) Then
           DrugroomPreview.GetNextDestinationTank Me
           DrugroomPreview.GetTimeBeforeTransfer Me, DrugroomPreview.DestinationTank
           If DrugroomPreview.CurrentRecipe Then
              DrugroomPreview.PreviewStatus = "Manual add for calloff " & DrugroomPreview.Tank1Calloff
           Else
              DrugroomPreview.PreviewStatus = "No Recipe manual add for calloff " & DrugroomPreview.Tank1Calloff
           End If
           DrugroomCalloffRefreshed = True
        Else
           DrugroomCalloffRefreshed = True
           DrugroomPreview.DestinationTank = 0
           DrugroomPreview.PreviewStatus = ""
           DrugroomPreview.Tank1Calloff = 0
           DrugroomPreview.TimeToTransferTimer.TimeRemaining = 0
        End If
     End If
    
    'rescan if step number has changed
    If StepNumberWas <> Parent.StepNumber Then
        StepNumberWas = Parent.StepNumber
        DrugroomCalloffRefreshed = False
    End If
    
  Else
    DrugroomCalloffRefreshed = False
  End If

' Dispensing Stuff
  If LA.KP1.IsOn Then
    DispenseCalloff = LA.KP1.DispenseCalloff
  ElseIf KP.KP1.IsOn Then
    DispenseCalloff = KP.KP1.DispenseCalloff
  Else
    DispenseCalloff = 0
  End If
 
'Turn off LAActive is parameter_LAenabled = 0
 If Parameters_LookAheadEnabled = 0 Then LAActive = False

'get end time stuff.
  If Parent.IsProgramRunning Then
    If GetEndTimeTimer.Finished Or (StepStandardTime <> StepStandardTimeWas) Then
      StepStandardTimeWas = StepStandardTime
      EndTimeMins = GetEndTime(Me)
      'see if current step is overruning or not
      If TimeInStepValue <= StepStandardTime Then
        EndTimeMins = EndTimeMins + (StepStandardTime - TimeInStepValue)
      Else
        EndTimeMins = EndTimeMins + (1)
      End If
      EndTime = DateAdd("n", EndTimeMins, Now)
      GetEndTimeTimer = 60
    End If
  Else
    EndTime = ""
    StartTime = ""
    EndTimeMins = 0
  End If

'Use temperature probe with highest reading (for safety)
  VesTemp = IO_VesselTemp
  
'Is temperature safe ?
  TempSafe = (VesTemp < 2050) And IO_TempInterlock And Not IO_ContactThermometer
  VesselHeatSafe = (Not IO_ContactThermometer) And IO_MainPumpRunning And (VesTemp < 2800)

'Is pressure safe ?
  PressSafe = IO_PressureInterlock And (VesselPressure < 50)
  
'Is temperature reading valid ?
  TempValid = (VesTemp > 300) And (VesTemp < 3000) And _
              IO_MainPumpStart And IO_MainPumpRunning And (Not IO_ContactThermometer)

' Run the safety control to handle pressurisation
  SafetyControl.Run VesTemp, TempSafe, PressSafe, PR.IsOn

'Machine Safe Flag
  MachineSafe = SafetyControl.IsDepressurised

'PumpLevel
  If (VesselLevel > Parameters_PumpMinimumLevel) Then
    PumpLevelLostTimer = 2
  End If
  If PumpLevelLostTimer.Finished Then
     PumpLevelTimer = Parameters_PumpMinimumLevelTime
  End If
  PumpLevel = PumpLevelTimer.Finished

'Some useful status flags for valve control
  Dim EStop As Boolean: EStop = IO_EStop_PB
  Dim Halt As Boolean: Halt = Parent.IsPaused Or IO_EStop_PB
  Dim NStop As Boolean: NStop = Not EStop
  Dim NHalt As Boolean: NHalt = Not Halt
  Dim NHSafe As Boolean: NHSafe = NHalt And MachineSafe
  
'Temperature Control - set enable timer
  If Halt Or (Not TempValid) Or (Not IO_MainPumpRunning) Then
    TemperatureControl.resetEnableTimer
    TemperatureControlContacts.resetEnableTimer
  End If
  
'Temperature control - set control Parameters
  TemperatureControl.EnableDelay = 10
  TemperatureControlContacts.EnableDelay = 10

'Temperature Control - run code
  TemperatureControl.Run VesTemp
  TemperatureControl.CheckErrorsAndMakeAlarms Me

'Temperature Control contacts - run code
  TemperatureControlContacts.Run VesTemp, TC.Gradient, Me
  TemperatureControlContacts.CheckErrorsAndMakeAlarms VesTemp, TC.Gradient, Me

' Only begin making alarms 5 seconds after start-up
  Static PowerUpTimer As acTimer
  If PowerUpTimer Is Nothing Then
    Set PowerUpTimer = New acTimer
    PowerUpTimer = 5
  ElseIf PowerUpTimer.Finished Then
    Alarms_MakeAlarms  ' ok to make alarms
  End If
 
'Vessel Level
  VesselLevel = IO_VesselLevelInput
  VesselLevelRange = Parameters_VesselLevelMax - Parameters_VesselLevelMin
  VesselLevelUncorrected = IO_VesselLevelInput - Parameters_VesselLevelMin
  If VesselLevelRange > 0 Then VesselLevel = (VesselLevelUncorrected * 1000) / VesselLevelRange
  If VesselLevel < 0 Then VesselLevel = 0
  If VesselLevel > 1000 Then VesselLevel = 1000

'Reserve Tank Level
  ReserveLevel = IO_ReserveLevelInput
  ReserveLevelRange = Parameters_ReserveLevelMax - Parameters_ReserveLevelMin
  ReserveLevelUncorrected = IO_ReserveLevelInput - Parameters_ReserveLevelMin
  If ReserveLevelRange > 0 Then ReserveLevel = (ReserveLevelUncorrected * 1000) / ReserveLevelRange
  If ReserveLevel < 0 Then ReserveLevel = 0
  If ReserveLevel > 1000 Then ReserveLevel = 1000

'Addition Tank Level
  AdditionLevel = IO_AdditionLevelInput
  AdditionLevelRange = Parameters_AdditionLevelMax - Parameters_AdditionLevelMin
  AdditionLevelUncorrected = IO_AdditionLevelInput - Parameters_AdditionLevelMin
  If AdditionLevelRange > 0 Then AdditionLevel = (AdditionLevelUncorrected * 1000) / AdditionLevelRange
  If AdditionLevel < 0 Then AdditionLevel = 0
  If AdditionLevel > 1000 Then AdditionLevel = 1000

'Differential Pressure across the packages
  If (IO_PackageDiffPressInput >= Parameters_PackageDiffPresMin) Then
    PackagePressureRange = Parameters_PackageDiffPresMax - Parameters_PackageDiffPresMin
    PackagePressureUncorrected = IO_PackageDiffPressInput - Parameters_PackageDiffPresMin
  Else
    PackagePressureRange = Parameters_PackageDiffPresMin
    PackagePressureUncorrected = Parameters_PackageDiffPresMin - IO_PackageDiffPressInput
  End If
  If PackagePressureRange > 0 Then PackageDifferentialPressure = ((PackagePressureUncorrected * Parameters_PackageDiffPresRange) / PackagePressureRange)
  If PackageDifferentialPressure < 0 Then PackageDifferentialPressure = 0
  If PackageDifferentialPressure > Parameters_PackageDiffPresRange Then PackageDifferentialPressure = Parameters_PackageDiffPresRange
 
'Vessel Pressure
  VesselPressure = IO_VesselPressInput
  VesselPressureRange = Parameters_VesselPressureMax - Parameters_VesselPressureMin
  VesselPressureUncorrected = IO_VesselPressInput - Parameters_VesselPressureMin
  If VesselPressureRange > 0 Then VesselPressure = (VesselPressureUncorrected * Parameters_VesselPressureRange) / VesselPressureRange
  If VesselPressure < 0 Then VesselPressure = 0
  If VesselPressure > Parameters_VesselPressureRange Then VesselPressure = Parameters_VesselPressureRange
 
'Flow Rate
  FlowRate = IO_FlowRateInput
  FlowRateRange = Parameters_FlowRateMax - Parameters_FlowRateMin
  FlowRateUncorrected = IO_FlowRateInput - Parameters_FlowRateMin
  If FlowRateRange > 0 Then FlowRate = (FlowRateUncorrected * Parameters_FlowRateRange) / FlowRateRange
  If FlowRate < 0 Then FlowRate = 0
  If FlowRate > Parameters_FlowRateRange Then FlowRate = Parameters_FlowRateRange
  If (Parameters_FlowRateRange > 0) Then
     FlowRatePercent = ((FlowRate / Parameters_FlowRateRange) * 1000)
  Else
     FlowRatePercent = 0
  End If
  
'Aninp8 - Test input for Maintenance (Parameters for user scaling & operation)
  TestValue = IO_TestInput
  TestValueRange = Parameters_TestValueMax - Parameters_TestValueMin
  TestValueUncorrected = IO_TestInput - Parameters_TestValueMin
  If TestValueRange > 0 Then TestValue = (TestValueUncorrected * Parameters_TestValueRange) / TestValueRange
  If TestValue < 0 Then TestValue = 0
  If TestValue > Parameters_TestValueRange Then TestValue = Parameters_TestValueRange
  If (Parameters_TestValueRange > 0) Then
     TestValuePercent = ((TestValue / Parameters_TestValueRange) * 1000)
  Else
     TestValuePercent = 0
  End If
    
'Tank 1
  Tank1Level = IO_Tank1LevelInput
  Tank1LevelRange = Parameters_Tank1LevelMax - Parameters_Tank1LevelMin
  Tank1LevelUncorrected = IO_Tank1LevelInput - Parameters_Tank1LevelMin
  If Tank1LevelRange > 0 Then Tank1Level = (Tank1LevelUncorrected * 1000) / Tank1LevelRange
  If Tank1Level < 0 Then Tank1Level = 0
  If Tank1Level > 1000 Then Tank1Level = 1000

'tank 1 time pot
  Tank1TimePot = IO_TimePotInput
  Tank1TimePotRange = Parameters_Tank1TimePotMax - Parameters_Tank1TimePotMin
  Tank1TimePotUncorrected = IO_TimePotInput - Parameters_Tank1TimePotMin
  If Tank1TimePotRange > 0 Then Tank1TimePot = (Tank1TimePotUncorrected * Parameters_Tank1TimePotRange) / Tank1TimePotRange
  If Tank1TimePot < 0 Then Tank1TimePot = 0
  If Tank1TimePot > Parameters_Tank1TimePotRange Then Tank1TimePot = Parameters_Tank1TimePotRange

'tank 1 temp pot
  Tank1TempPot = IO_TempPotInput
  Tank1TempPotRange = Parameters_Tank1TempPotMax - Parameters_Tank1TempPotMin
  Tank1TempPotUncorrected = IO_TempPotInput - Parameters_Tank1TempPotMin
  If Tank1TempPotRange > 0 Then Tank1TempPot = (Tank1TempPotUncorrected * Parameters_Tank1TempPotRange) / Tank1TempPotRange
  If Tank1TempPot < 0 Then Tank1TempPot = 0
  If Tank1TempPot > Parameters_Tank1TempPotRange Then Tank1TempPot = Parameters_Tank1TempPotRange

'Toggle Tank1Ready flag if add ready pushbutton is pressed
  Static PreviousTank1Ready_PB As Boolean
  If IO_Tank1Ready_PB And Not PreviousTank1Ready_PB Then Tank1Ready = Not Tank1Ready
  PreviousTank1Ready_PB = IO_Tank1Ready_PB

'Toggle AddReady flag if add ready pushbutton is pressed
  Static PreviousAddReady_PB As Boolean
  Static PreviousRun_PB As Boolean
  If (AddPreparePB And Not PreviousAddReady_PB) Or _
    (IO_RemoteRun And AP.IsOn And Not PreviousRun_PB) _
  Then AddReadyFromPB = Not AddReadyFromPB
  PreviousAddReady_PB = AddPreparePB
  
'Toggle ReserveReady flag if add ready pushbutton is pressed
  Static PreviousReserveReady_PB As Boolean
  If (ReservePreparePB And Not PreviousReserveReady_PB) Or _
    (IO_RemoteRun And RP.IsOn And Not PreviousRun_PB) _
  Then ReserveReadyFromPB = Not ReserveReadyFromPB
  PreviousReserveReady_PB = ReservePreparePB
  PreviousRun_PB = IO_RemoteRun

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

'airpad control
 If VesselPressure >= Parameters_AirpadCutoffPressure Then
    AirpadCutoff = True
 Else
    AirpadCutoff = False
 End If
 If VesselPressure > Parameters_AirpadBleedPressure Then
    AirpadBleed = True
 Else
    AirpadBleed = False
 End If
 
'drugroom mixer control
 If Tank1Level > Parameters_Tank1MixerOnLevel Then DrugroomMixerOn = True
 If Tank1Level <= Parameters_Tank1MixerOffLevel Then DrugroomMixerOn = False
 
'Toggle Tank 1 Mixer is pushbutton is pressed and tank 1 manual switch is in "Manual"
  Static PreviousTank1Mixer_PB As Boolean
  If IO_Tank1Mixer_PB And Not PreviousTank1Mixer_PB Then Tank1MixerRequest = Not Tank1MixerRequest
  PreviousTank1Mixer_PB = IO_Tank1Mixer_PB
  If Not (DrugroomMixerOn And IO_Tank1Manual_SW) Then Tank1MixerRequest = False

'reserve mixer control
 If ReserveLevel > Parameters_ReserveMixerOnLevel Then ReserveMixerOn = True
 If ReserveLevel <= Parameters_ReserveMixerOffLevel Then ReserveMixerOn = False
 
'addition tank mix control
  If AdditionLevel > Parameters_AddMixOnLevel Then AdditionMixOn = True
  If AdditionLevel <= Parameters_AddMixOffLevel Then AdditionMixOn = False
  
'volume based on level used for bath turnovers is in gallons
   VolumeBasedOnLevel = Parameters_SystemVolumeAt0Level + (((Parameters_SystemVolumeAt100Level - Parameters_SystemVolumeAt0Level) * VesselLevel) / 1000)
 
'figure out the system flow rate in seconds
  If FlowRateTimer.Finished Then
    If (BatchWeight > 0) Then
      GallonsPerMinPerPound = ((FlowRate / BatchWeight) * 10) 'in tenths
    Else: GallonsPerMinPerPound = 0
    End If
  
    'figure out number of contacts per min for tc command
    If IO_MainPumpRunning And (VolumeBasedOnLevel > 0) Then
      NumberofContactsPerMin = (FlowRate / VolumeBasedOnLevel)
    Else: NumberofContactsPerMin = 0
    End If
  
    'figure out number of contacts
    SystemVolume = SystemVolume + (FlowRate / 60) 'this is in gallons
    If SystemVolume >= (VolumeBasedOnLevel) Then
      SystemVolume = SystemVolume - (VolumeBasedOnLevel)
      NumberOfContacts = NumberOfContacts + 1
      TotalNumberOfContacts = TotalNumberOfContacts + 1
    End If
    FlowRateTimer = 1
  End If
 
 'figure out the recorded level stuff
 If GetRecordedLevel Then
   If GetRecordedLevelTimer.Finished And Not GotRecordedLevel Then
      GotRecordedLevel = True
      RecordedLevel = VesselLevel
   End If
 Else
   RecordedLevel = 0
   GotRecordedLevel = False
   GetRecordedLevelTimer = Parameters_RecordLevelTimer
 End If
 
'Digital ouputs in order ==================================================
'Outputs 1-16
'for siren and signal light instead of using signal stuff
  SignalOnRequest = LD.LDSignalOnRequest Or UL.ULSignalOnRequest Or _
                    SA.SASignalOnRequest Or PH.PHSignalOnRequest
  If SignalOnRequest Then
     If SignalRequestedWas = False Then
        SignalRequested = True
        SignalRequestedWas = True
     End If
     If (IO_RemoteRun = True) Then SignalRequested = False
  Else
     SignalRequested = False
     SignalRequestedWas = False
  End If

'Signal and delay lights and siren outputs
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
  IO_BlendFillOutput = 0
  If NHalt Then
    If FI.IsFilling Then IO_BlendFillOutput = FI.Output
    If RC.IsRinsing Or RC.IsFilling Then IO_BlendFillOutput = RC.BlendOutput
    If RI.IsRinsing Or RI.IsFilling Then IO_BlendFillOutput = RI.BlendOutput
    If RH.IsRinsing Or RH.IsFilling Then IO_BlendFillOutput = RH.BlendOutput
    If HD.IsFillingToCool Then IO_BlendFillOutput = Parameters_HDBlendFillPosition
    If IO_CityWater_SW Then IO_BlendFillOutput = 0
  End If
  If IO_BlendFillOutput < 0 Then IO_BlendFillOutput = 0
  If IO_BlendFillOutput > 1000 Then IO_BlendFillOutput = 1000
 
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
  
'outputs done=================================================================

' Press buttons
  Parent.PressButtons IO_RemoteRun, IO_RemoteHalt, False, False, False
  
'Set delay variable
  Delays_MakeDelays
 
'Utility usage calculations
'Assume all motors run at about 68% of rated capacity. 750W is equiv to HP.
'Power factor = 68% of 750 , convert to watts
  Dim PowerFactor As Long, HPNow As Long
  PowerFactor = (68 * 30) / 4
  HPNow = 0 'Reset power being used now uasge to 0
  If Utilities_Timer.Finished Then
    MainPumpHP = Parameters_MainPumpHP
   If IO_MainPumpRunning Then
      HPNow = HPNow + MainPumpHP
      
   End If
   PowerKWS = PowerKWS + ((HPNow * PowerFactor) / 100)
       
    Utilities_Timer = 10 'specifies how often this routine runs, in seconds
  End If
  If (PowerKWS >= 3600) Then
    PowerKWH = PowerKWH + 1
    PowerKWS = PowerKWS - 3600
  End If
  'Steam Consumption formula
  '120% * WorkingVolume * deg F of temp rise * weight of water (8.33lb/g) * 0.001 = lbs of steam used
  'Steam factor = 120% * 8.33 * 0.001 = 1/100
  'lbs of Steam used = working volume * Temp rise in F / 100
  'Do Steam Calculation
  If TC.IsOn Then
     FinalTemp = TemperatureControlContacts.TempFinalTemp
  Else
     FinalTemp = TemperatureControl.TempFinalTemp
  End If
  If FinalTemp <> FinalTempWas Then 'Changed final temp
    If VesTemp > FinalTempWas Then 'Choose highest of these, so we don't count steam twice
      StartTemp = VesTemp
    Else
      StartTemp = FinalTempWas
    End If
    If FinalTemp > StartTemp Then 'We're gonna heat
      TempRise = FinalTemp - StartTemp
      SteamNeeded = (VolumeBasedOnLevel * TempRise) / 1000
      SteamUsed = SteamUsed + SteamNeeded
    End If
    FinalTempWas = FinalTemp
  End If
  
'Graph logging    graph colors
'1- light red   2- pink   3- light blue   4- dark blue    5 - dark red    6- green    7- tan    8- light red    9- pink
 
 If TC.IsOn Then
  SetpointF = TemperatureControlContacts.TempSetpoint
 Else
  SetpointF = TemperatureControl.TempSetpoint
 End If
  TempFinalValue = Pad(SetpointF / 10, "0", 3) & "F"

'get current time
  CurrentTime = CStr(Now())

'this should halt the control if sleeping
  If Parent.IsSleeping And (Not IsSleepingWas) Then
     IsSleepingWas = True
     IO_RemoteHalt = True
  End If
  If (Not Parent.IsSleeping) And IsSleepingWas Then
    IsSleepingWas = False
  End If
  
'mimic buttons set to false
 Static MainPumpResetTimer As New acTimer
 If MainPumpResetPB = False Then MainPumpResetTimer = 5
 If MainPumpResetTimer.Finished And MainPumpResetPB Then MainPumpResetPB = False
  
'add tank mimic stuff
 'manual add filling
  ManualAddFill.Run ManualAddFillLevel, AdditionLevel, Parameters_AdditionFillDeadband, _
                   ManualAddFillHot, ManualAddFillCold
 
 'manual add draining
  If ManualAddDrain.IsTurnOff Then ManualAddDrainPB = False
  ManualAddDrain.Run AdditionLevel, ManualAddDrainPB, ManualAddDrainFinished, _
                       Parameters_AddTransferToDrainTime, Parameters_AddTransferRinseToDrainTime
 
 'manual add transferring
  ManualAddTransfer.Run AdditionLevel, ManualAddTransferPB, ManualAddTransferFinished, _
                       Parameters_AddTransferTimeBeforeRinse, Parameters_AddTransferRinseTime, _
                       Parameters_AddTransferTimeAfterRinse, Parameters_AddTransferRinseToDrainTime, Parameters_AddTransferToDrainTime
                       
 'manual reserve filling
  ManualReserveFill.Run ManualReserveFillLevel, ReserveLevel, ManualReserveFillHot, ManualReserveFillCold
 
 'manual reserve draining
  If ManualReserveDrain.IsTurnOff Then ManualReserveDrainPB = False
  ManualReserveDrain.Run ReserveLevel, ManualReserveDrainPB, ManualReserveDrainFinished, _
                       Parameters_ReserveDrainTime, Parameters_ReserveRinseToDrainTime

 'manual add transferring
  ManualReserveTransfer.Run Me
    
 'manual heat
  ManualReserveHeat.Run ManualReserveDesiredTemp, IO_ReserveTemp, ReserveLevel, ManualReserveHeatPB
   
 'Manual Steam Test
  ManualSteamOverride = (Parent.Mode = modeOverride)
  If Not ManualSteamOverride Then SteamTestPB = False
  ManualSteamTest.Run VesselLevel, ReserveLevel, Tank1Level, VesTemp, IO_Tank1Temp, IO_ReserveTemp, SteamTestPB, ManualSteamOverride
 
'Keep this at the end of this routine
  FirstScanDone = True
End Sub

Public Property Get Status() As String
  Status = ""
  If Parent.Signal <> "" Then
    Status = Parent.Signal
  ElseIf Alarms_LidNotLocked Then
    Status = "Lid Not Locked: Lock and Hold Run " & TimerString(LidLockTimer.TimeRemaining)
  ElseIf (AP.IsSlow Or AP.IsFast) And SlowFlash Then
    Status = "Local Add Addition Tank: Prepare"
  ElseIf (RP.IsSlow Or RP.IsFast) And SlowFlash Then
    Status = "Local Add Reserve Tank: Prepare"
  ElseIf AT.IsOn And AT.StateString <> "" Then
    Status = AT.StateString
  ElseIf AC.IsOn And AC.StateString <> "" Then
    Status = AC.StateString
  ElseIf CO.IsRamping And CO.StateString <> "" Then
    Status = CO.StateString
  ElseIf DR.IsOn And DR.StateString <> "" Then
    Status = DR.StateString
  ElseIf FI.IsOn And FI.StateString <> "" Then
    Status = FI.StateString
  ElseIf HD.IsOn And HD.StateString <> "" Then
    Status = HD.StateString
  ElseIf HE.IsRamping And HE.StateString <> "" Then
    Status = HE.StateString
  ElseIf HC.IsOn And HC.StateString <> "" Then
    Status = HC.StateString
  ElseIf LD.IsOn And LD.StateString <> "" Then
    Status = LD.StateString
  ElseIf PH.IsOn And PH.StateString <> "" Then
    Status = PH.StateString
  ElseIf RC.IsOn And RC.StateString <> "" Then
    Status = RC.StateString
  ElseIf RH.IsOn And RH.StateString <> "" Then
    Status = RH.StateString
  ElseIf RF.IsForeground And RF.StateString <> "" Then
    Status = RF.StateString
  ElseIf RI.IsOn And RI.StateString <> "" Then
    Status = RI.StateString
  ElseIf RT.IsOn And RT.StateString <> "" Then
    Status = RT.StateString
  ElseIf RW.IsOn And RW.StateString <> "" Then
    Status = RW.StateString
  ElseIf SA.IsOn And SA.StateString <> "" Then
    Status = SA.StateString
  ElseIf TC.IsRamping And TC.StateString <> "" Then
    Status = TC.StateString
  ElseIf TM.IsOn Then
    Status = "Holding " & TimerString(TimeToGo)
  ElseIf UL.IsOn And UL.StateString <> "" Then
    Status = UL.StateString
  ElseIf WK.IsOn Then
      If KP.KP1.IsInManual Then Status = "WK: Drugroom In Manual Mode "
      If KP.KP1.IsHeatTankOverrun Then Status = "WK: Drugroom Too Long to Heat "
      If KP.KP1.IsWaitIdle Then Status = "WK: Drugroom Waiting for RD to finish "
      If KP.KP1.IsDispensing Then Status = "WK: Waiting for Dispense "
      If KP.KP1.IsFill Then Status = "WK: Filling Drugroom Tank " & Pad(KP.KP1.AddFillLevel / 10, "0", 3) & "%"
      If KP.KP1.IsHeating Then Status = "WK: Heating Drugroom Tank " & Pad(KP.KP1.DesiredTemperature / 10, "0", 3) & "F"
      If KP.KP1.IsSlow Or KP.KP1.IsFast Then Status = "WK: Waiting for Drugroom Ready "
      If KP.KP1.IsMixing Then Status = "WK: Drugroom Mixing " & TimerString(KP.KP1.MixTimer)
      If KP.KP1.IsTransfer Then Status = "WK: Drugroom Transfering " & TimerString(KP.KP1.Timer)
      If KP.KP1.IsRinse Then Status = "WK: Drugroom Rinse to Machine " & TimerString(KP.KP1.Timer)
      If KP.KP1.IsRinseToDrain Then Status = "WK: Drugroom Rinse to Drain " & TimerString(KP.KP1.Timer)
      If KP.KP1.IsInterlocked Then Status = "WK: Add Tank Too High for Tank 1 Transfer "
      If Not KP.KP1.IsOn Then
        If WK.IsDelayOff Then
          Status = "Waiting for drugroom " & TimerString(WK.Timer)
        Else: Status = "Waiting for drugroom "
        End If
      End If
  ElseIf WT.IsOn Then
    Status = "Wait for Temp " & Pad(TemperatureControl.TempFinalTemp / 10, "0", 3) & "F"
  ElseIf Not Parent.IsProgramRunning Then
    Static ProgStopTime As Long
    ProgStopTime = ProgramStoppedTimer
    Status = "Machine Idle: " & TimerString(ProgStopTime)
  End If
End Property
Private Function TISValue(Index As Long) As Long
  On Error GoTo Err_TISValue
    Dim Values() As String
    Values = Split(Parent.TimeInStep, "/")
    TISValue = Values(Index)
  Exit Function
Err_TISValue:
  TISValue = 0
End Function

Private Property Get ACControlCode_DrawScreenButtons() As String
  ACControlCode_DrawScreenButtons = "Main,1;Data,7;Dispense,7;Tank 1,4;Add,4;Reserve,4;Flow,2;Temp,3;Lid,7;"
End Property

Public Sub ACControlCode_DrawScreen(ByVal CurrentScreen As Long, Row() As String)
  Dim i As Long
  For i = 1 To 15
    Row(i) = ""
  Next i
  Dim TextRow7 As String
  TextRow7 = ""
  Row(15) = TextRow7
  
  Select Case CurrentScreen
    Case 1: 'Main
      Row(1) = "Procedure " & Parent.ProgramNumber
      Dim PressureStateText As String
      PressureStateText = "Pressurized"
      If SafetyControl.IsDepressurised Then PressureStateText = "Depressurized"
      If SafetyControl.IsDepressurising Then PressureStateText = "Depressurizing"
      Row(2) = PressureStateText
      Row(3) = "Vessel Level " & Pad(VesselLevel / 10, "0.0", 4) & "%"
      Row(4) = "Vessel Volume " & Pad(VolumeBasedOnLevel, "0", 4) & "Gals"
      Row(5) = "Vessel Pressure " & Pad(VesselPressure / 10, "0.0", 4) & "psi"
      If IO_MainPumpRunning Then
        Row(6) = "Pump Running " & Pad(IO_PumpSpeedOutput / 10, "0", 3) & "%"
        If IO_FlowReverse Then
          If FR.IsOn Then
             Row(7) = "Flow Out to In " & TimerString(FR.Timer)
          ElseIf FC.IsOn Then
             Row(7) = "Flow Out to In " & (FC.OutToInTurns - NumberOfContacts)
          End If
          
          If DP.IsOn Then
             Row(11) = "Desired Pressure " & Pad(DP.OutToInPressure / 10, "0.0", 4) & "psi"
          End If
          If FL.IsOn Then
             Row(12) = "Desired Flowrate " & Pad(FL.OutToInFlowRate, "0", 4) & "gpm"
          End If
        Else
          If FR.IsOn Then
             Row(7) = "Flow In to Out " & TimerString(FR.Timer)
          ElseIf FC.IsOn Then
             Row(7) = "Flow In to Out " & (FC.InToOutTurns - NumberOfContacts)
          End If
          If DP.IsOn Then
             Row(11) = "Desired Pressure " & Pad(DP.InToOutPressure / 10, "0.0", 4) & "psi"
          End If
          If FL.IsOn Then
             Row(12) = "Desired Flowrate " & Pad(FL.InToOutFlowRate, "0", 4) & "gpm"
          End If
        End If
        Row(8) = "Package Differential " & Pad(PackageDifferentialPressure / 10, "0.0", 3) & " psi"
        Row(9) = "Flow Rate " & Pad(FlowRate, "0", 4) & " gpm"
        Row(10) = Pad(GallonsPerMinPerPound, "0.0", 4) & " gpm/lb"
     Else
        Row(6) = "Pump is off"
      End If
      If IO_Fill Then
         If FI.IsActive Then
          If (Parameters_EnableFillByWorkingLevel = 1) And (WorkingLevel > 0) Then
            Row(13) = "Desired Working Level " & Pad(WorkingLevel / 10, "0.0", 4) & "%"
          Else
            Row(13) = "Desired Fill Level " & Pad(FI.FillLevel / 10, "0.0", 4) & "%"
          End If
          Row(14) = "Desired Fill Temp. " & Pad(FI.FillTemperature / 10, "0.0", 4) & "F"
          Row(15) = "Actual Fill Temp. " & Pad(IO_BlendFillTemp / 10, "0.0", 4) & "F"
          Row(16) = "Blend Fill Valve " & Pad(IO_BlendFillOutput / 10, "0.0", 4) & "%"
         End If
         If RI.IsActive Then
           Row(13) = "Desired Rinse Temp. " & Pad(RI.RinseTemperature / 10, "0.0", 4) & "F"
           Row(14) = "Actual Rinse  Temp. " & Pad(IO_BlendFillTemp / 10, "0.0", 4) & "F"
           Row(15) = "Blend Fill Valve " & Pad(IO_BlendFillOutput / 10, "0.0", 4) & "%"
         End If
         If RH.IsActive Then
           Row(13) = "Desired Rinse Temp. " & Pad(RH.RinseTemperature / 10, "0.0", 4) & "F"
           Row(14) = "Actual Rinse Temp. " & Pad(IO_BlendFillTemp / 10, "0.0", 4) & "F"
           Row(15) = "Blend Fill Valve " & Pad(IO_BlendFillOutput / 10, "0.0", 4) & "%"
         End If
         If RC.IsActive Then
           Row(13) = "Desired Rinse Temp. " & Pad(RC.RinseTemperature / 10, "0.0", 4) & "F"
           Row(14) = "Actual Rinse Temp. " & Pad(IO_BlendFillTemp / 10, "0.0", 4) & "F"
           Row(15) = "Blend Fill Valve " & Pad(IO_BlendFillOutput / 10, "0.0", 4) & "%"
         End If
       End If
      
    Case 2: 'Data
      Dim Page6Text2 As String
      Dim Page6Text3 As String
      If Parent.IsProgramRunning Then
        Page6Text2 = "Cycle Time " & TimerString(CycleTime)
      Else
        Page6Text2 = "Cycle Time " & TimerString(LastProgramCycleTime)
      End If
      If Parent.IsProgramRunning Then
         Row(1) = Parent.ProgramNumber
         Row(2) = "Start Time    " & StartTime
         Row(3) = "End Time      " & EndTime
         Row(4) = "Standard Time " & TimerString(CStr(StandardTime))
         Row(5) = Page6Text2
         Row(6) = "Batch Weight: " & Pad(BatchWeight, "0", 4) & " lbs"
         Row(8) = "Package Height: " & Pad(PackageHeight, "0", 2) & " Type: " & PT.PackageType
         Row(9) = "Working Level: " & Pad(WorkingLevel / 10, "0.0", 4) & "%"
         Row(10) = "Total Number of "
         Row(11) = "Contacts: " & TotalNumberOfContacts
         If IO_CityWater_SW Then Row(12) = "Using City Water"
         Row(13) = "Version: " & Parent.ControlCodeFileVersion
      Else
         Row(1) = "Version: " & Parent.ControlCodeFileVersion
         Row(2) = ""
         Row(3) = ""
         Row(4) = ""
         Row(5) = ""
         Row(6) = ""
         Row(7) = ""
         Row(8) = ""
         Row(9) = ""
         Row(10) = ""
         Row(11) = ""
         Row(12) = ""
         Row(13) = ""
      End If

   Case 3:  'Dispense
      If DispenseTank = 2 Then
         Row(1) = "Dispensing to Reserve tank"
      ElseIf DispenseTank = 1 Then
         Row(1) = "Dispensing to add tank"
      Else
         Row(1) = ""
      End If
      Dim Products() As String, m As Long
      Products = Split(Dispenseproducts, "|")
      For m = 1 To UBound(Products)
        Row(m + 1) = Products(m)
      Next m
      
    Case 4: 'Tank 1
      Row(1) = "Tank 1 Level " & Pad(Tank1Level / 10, "0.0", 4) & "%"
      Row(2) = "Tank 1 Temperature " & Pad(IO_Tank1Temp / 10, "0.0", 4) & "F"
      If IO_Tank1Manual_SW Then
        Row(3) = "Drugroom Tank in Manual "
      Else
        If KP.KP1.IsFill Then
          Row(3) = "Tank 1 Filling to " & Pad((KP.KP1.AddFillLevel + Parameters_Tank1FillLevelDeadBand) / 10, "0.0", 4) & "%"
        ElseIf KP.KP1.IsHeating Then
          Row(3) = "Heating to " & Pad((KP.KP1.DesiredTemperature + Parameters_Tank1HeatDeadband) / 10, "0.0", 4) & "F"
        ElseIf KP.KP1.IsSlow Then
          Row(3) = "Prepare Tank 1"
        ElseIf KP.KP1.IsFast Then
          Row(3) = "Waiting for Tank 1"
        ElseIf KP.KP1.IsMixing Then
          Row(3) = "Mixing for " & TimerString(KP.KP1.MixTimer)
        ElseIf LA.KP1.IsFill Then
          Row(3) = "Tank 1 Filling to " & Pad((LA.KP1.AddFillLevel + Parameters_Tank1FillLevelDeadBand) / 10, "0.0", 4) & "%"
        ElseIf LA.KP1.IsHeating Then
          Row(3) = "Heating to " & Pad((LA.KP1.DesiredTemperature + Parameters_Tank1HeatDeadband) / 10, "0.0", 4) & "F"
        ElseIf LA.KP1.IsSlow Then
          Row(3) = "Prepare Tank 1"
        ElseIf LA.KP1.IsFast Then
          Row(3) = "Waiting for Tank 1"
        ElseIf LA.KP1.IsMixing Then
          Row(3) = "Mixing for " & TimerString(LA.KP1.MixTimer)
        ElseIf KP.KP1.IsPaused Or LA.KP1.IsPaused Or KA.IsPaused Then
          Row(3) = "Tank 1 Transfer Paused"
        ElseIf KA.IsRinse Then
          Row(3) = "Tank 1 Rinsing " & TimerString(KA.Timer)
        ElseIf KP.KP1.IsRinse Then
          Row(3) = "Tank 1 Rinsing " & TimerString(KP.KP1.Timer)
        ElseIf LA.KP1.IsRinse Then
          Row(3) = "Tank 1 Rinsing " & TimerString(LA.KP1.Timer)
        ElseIf KA.IsRinseToDrain Then
          Row(3) = "Tank 1 Rinsing to Drain " & TimerString(KA.Timer)
        ElseIf KP.KP1.IsRinseToDrain Then
          Row(3) = "Tank 1 Rinsing to Drain " & TimerString(KP.KP1.Timer)
        ElseIf LA.KP1.IsRinseToDrain Then
          Row(3) = "Tank 1 Rinsing to Drain " & TimerString(LA.KP1.Timer)
        ElseIf KA.IsTransferToDrain Then
          Row(3) = "Tank 1 Transfer to Drain " & TimerString(KA.Timer)
        ElseIf KP.KP1.IsTransferToDrain Then
          Row(3) = "Tank 1 Transfer to Drain " & TimerString(KP.KP1.Timer)
        ElseIf LA.KP1.IsTransferToDrain Then
          Row(3) = "Tank 1 Transfer to Drain " & TimerString(LA.KP1.Timer)
        ElseIf KA.IsTransfer Then
          Row(3) = "Tank 1 Transferring " & TimerString(KA.Timer)
        ElseIf KP.KP1.IsTransfer Then
          Row(3) = "Tank 1 Transferring " & TimerString(KP.KP1.Timer)
        ElseIf LA.KP1.IsTransfer Then
          Row(3) = "Tank 1 Transferring " & TimerString(LA.KP1.Timer)
        ElseIf KP.KP1.IsReady Then
          Row(3) = "Tank 1 Ready"
        ElseIf LA.KP1.IsReady Then
          Row(3) = "Tank 1 Ready"
        ElseIf CK.IsFilling Then
          Row(3) = "Tank 1 Filling to " & Pad((Parameters_CKFillLevel) / 10, "0.0", 4) & "%"
        ElseIf CK.IsTransferToDrain Then
          Row(3) = "Tank 1 Transferer to Drain " & TimerString(CK.Timer)
        ElseIf CK.IsRinseToDrain Then
          Row(3) = "Rinse to Drain"
        Else
          Row(3) = "Tank 1 idle"
        End If
      End If
      If LA.KP1.IsOn Then Row(4) = "Looking ahead to lot " & LA.Job
      
    Case 5:  'Add
      Row(1) = "Addition Level " & Pad(AdditionLevel / 10, "0.0", 4) & "%"
      If AdditionMixOn Then Row(2) = "Addition mixing"
      If AF.IsActive And AF.StateString <> "" Then
        Row(3) = AF.StateString
      ElseIf AT.IsActive Then
        Row(3) = AT.StateString
      ElseIf AC.IsActive Then
        Row(3) = AC.StateString
      ElseIf AD.IsActive Then
        Row(3) = AD.StateString
      Else
        Row(3) = "Addition Idle"
      End If
     
      If AT.IsDosePaused Or AT.IsDosing Then
         Row(4) = "Actual Level " & Pad(AdditionLevel / 10, "0.0", 4) & "%"
         Row(5) = "Desired Level " & Pad(AT.DesiredLevel / 10, "0.0", 4) & "%"
      ElseIf AC.IsDosePaused Or AC.IsDosing Then
         Row(4) = "Actual Level " & Pad(AdditionLevel / 10, "0.0", 4) & "%"
         Row(5) = "Desired Level " & Pad(AC.LevelDesired / 10, "0.0", 4) & "%"
      End If
      If AP.IsActive Then
         If (AP.AddPreparePrompt >= 1) Or (AP.AddPreparePrompt <= 99) Then
            Row(6) = Parent.message(AP.AddPreparePrompt)
            If AP.IsReady Then
               Row(7) = "Manual Add Ready"
            Else
               Row(7) = "Prepare Manual Add"
            End If
         End If
      End If

     Case 6: 'Reserve
      Row(1) = "Reserve Level " & Pad(ReserveLevel / 10, "0.0", 4) & "%"
      Row(2) = "Reserve Temp: " & Pad(IO_ReserveTemp / 10, "0.0", 4) & "F"
      If ReserveMixerOn Then Row(3) = "Reserve mixing"
      If RF.IsHeating Or RF.HeatOn Then
         Row(4) = "Heating to " & Pad(RF.DesiredTemperature / 10, "0.0", 4) & "F"
      End If
      If RF.IsOn Then
        Row(5) = RF.StateString
        If (Parameters_EnableReserveFillByWorkingLevel = 1) And WorkingLevel > 0 Then
          Row(9) = "RF: Working Level = " & Pad(WorkingLevel / 10, "0.0", 4) & "%"
          Row(10) = "   Fill Level = " & Pad(RF.FillLevel / 10, "0.0", 4) & "%"
        Else
          Row(9) = "RF: Fill Level = " & Pad(RF.FillLevel / 10, "0.0", 4) & "%"
          Row(10) = ""
        End If
      ElseIf RD.IsActive Then
         Row(5) = RD.StateString
      ElseIf RW.IsActive Then
         Row(5) = RW.StateString
      ElseIf RT.IsActive Then
         Row(5) = RT.StateString
      ElseIf KP.KP1.IsOn And KP.KP1.DispenseTank = 2 Then
         Row(5) = "Reserve Active with KP " & KP.KP1.StateString
      ElseIf LA.KP1.IsOn And LA.KP1.DispenseTank = 2 Then
         Row(5) = "Reserve Active with LA " & LA.KP1.StateString
      Else
         Row(5) = "Reserve Idle"
      End If
      If RP.IsActive Then
        If (RP.ReservePreparePrompt >= 1) Or (RP.ReservePreparePrompt <= 99) Then
           Row(6) = Parent.message(RP.ReservePreparePrompt)
           If RP.IsReady Then
             Row(7) = "Manual Add Ready"
           Else
             Row(7) = "Prepare Manual Add"
           End If
        End If
      End If
      
    Case 7: 'Flow
      If IO_MainPumpRunning Then
        Row(1) = "Pump Running " & Pad(IO_PumpSpeedOutput / 10, "0", 3) & "%"
        If FR.IsOn Then
          Row(2) = FR.StateString
        ElseIf FC.IsOn Then
          Row(2) = FC.StateString
        End If
        Row(3) = "Package Differential " & Pad(PackageDifferentialPressure / 10, "0.0", 3) & " psi"
        Row(4) = "Flow Rate " & Pad(FlowRate, "0", 4) & " gpm"
        Row(5) = "Flow Percent " & Pad(FlowRatePercent / 10, "0.0", 4) & "%"
        Row(6) = "Contacts/Minute " & Pad(NumberofContactsPerMin, "0.0", 4)
        Row(7) = Pad(GallonsPerMinPerPound, "0.00", 4) & " gpm/lb"
        If IO_FlowReverse Then
          If DP.IsOn Then
             Row(8) = "Desired Pressure " & Pad(DP.OutToInPressure / 10, "0.0", 4) & "psi"
          End If
          If FL.IsOn Then
             Row(9) = "Desired Flowrate " & Pad(FL.OutToInFlowRate, "0", 4) & "gpm"
          End If
        Else
          If DP.IsOn Then
             Row(8) = "Desired Pressure " & Pad(DP.InToOutPressure / 10, "0.0", 4) & "psi"
          End If
          If FL.IsOn Then
             Row(9) = "Desired Flowrate " & Pad(FL.InToOutFlowRate, "0", 4) & "gpm"
          End If
        End If
     Else
        Row(1) = "Pump is off"
      End If
      If Parent.IsProgramRunning Then
        Row(9) = "Flow In-Out: " & Pad((FR.InToOut + FC.InToOut), "0", 4)
        Row(10) = "Flow Out-In: " & Pad((FR.OutToIn + FC.OutToIn), "0", 4)
      Else
        Row(9) = "Previous Batch Flow In-Out: " & Pad((FR.InToOut + FC.InToOut), "0", 4)
        Row(10) = "Previous Batch Flow Out-In: " & Pad((FR.OutToIn + FC.OutToIn), "0", 4)
      End If
       
    Case 8:
     Dim StateText3 As String
     StateText3 = ""
     If TC.IsOn Then
        Row(1) = "Gradient " & Pad(TemperatureControlContacts.TempGradient / 10, "0.0", 4) & "F/m"
        Row(2) = "Final Temp " & Pad(TemperatureControlContacts.TempFinalTemp / 10, "0", 3) & "F"
        Row(3) = "Setpoint " & Pad(TemperatureControlContacts.TempSetpoint / 10, "0.0", 5) & "F"
        If TemperatureControlContacts.IsHeating Or TemperatureControlContacts.IsCooling Then
          Row(4) = "Valve Output " & Pad(IO_HeatCoolOutput / 10, "0", 3) & "%"
        End If
        
        If TemperatureControlContacts.IsHeating Then StateText3 = "Heating"
        If TemperatureControlContacts.IsCooling Then StateText3 = "Cooling"
        If TemperatureControlContacts.IsPreHeatVent Then StateText3 = "Vent before heat"
        If TemperatureControlContacts.IsPostHeatVent Then StateText3 = "Vent after heat"
        If TemperatureControlContacts.IsPreCoolVent Then StateText3 = "Vent before cool"
        If TemperatureControlContacts.IsPostCoolVent Then StateText3 = "Vent after cool"
        Row(5) = StateText3
      Else
        Row(1) = "Gradient " & Pad(TemperatureControl.TempGradient / 10, "0.0", 4) & "F/m"
        Row(2) = "Final Temp " & Pad(TemperatureControl.TempFinalTemp / 10, "0", 3) & "F"
        Row(3) = "Setpoint " & Pad(TemperatureControl.TempSetpoint / 10, "0.0", 5) & "F"
        If TemperatureControl.IsHeating Or TemperatureControl.IsCooling Then
          Row(4) = "Valve Output " & Pad(IO_HeatCoolOutput / 10, "0", 3) & "%"
        End If
        
        If TemperatureControl.IsHeating Then StateText3 = "Heating"
        If TemperatureControl.IsCooling Then StateText3 = "Cooling"
        If TemperatureControl.IsPreHeatVent Then StateText3 = "Vent before heat"
        If TemperatureControl.IsPostHeatVent Then StateText3 = "Vent after heat"
        If TemperatureControl.IsPreCoolVent Then StateText3 = "Vent before cool"
        If TemperatureControl.IsPostCoolVent Then StateText3 = "Vent after cool"
        Row(5) = StateText3
      End If
      
      If (IO_CondensateTemp > 320) And (IO_CondensateTemp < 3000) Then
        Row(6) = "Condensate Temp " & Pad(IO_CondensateTemp / 10, "0", 3) & "F"
      ElseIf Parameters_TempCondensateLimit > 0 Then
        Row(6) = "Condensate Temp " & Pad(IO_CondensateTemp / 10, "0", 3) & "F"
      Else
        Row(6) = ""
      End If

    Case 9: 'Lid
      If IO_OpenLid Or Not IO_LidLoweredSwitch Then
         Row(1) = "Lid Open"
      ElseIf IO_CloseLid Or IO_LidLoweredSwitch Then
         Row(1) = "Lid Closed"
      End If
      If IO_OpenLockingBand Then
         Row(2) = "Locking Band Open "
      ElseIf IO_CloseLockingBand Then
         Row(2) = "Locking Band Closed "
      End If
      If IO_LockingPinRaised Then
         Row(3) = "Locking Pin Raised "
      Else
         Row(3) = "Locking Pin Lowered "
      End If
      Row(4) = LidControl.StateString
  
  End Select
End Sub

Private Function ACControlCode_ReadInputs(Dinp() As Boolean, Aninp() As Integer, Temp() As Integer) As Boolean
  'NoResponse used to flag comms failure, i used to iterate through arrays
  Dim NoResponse As Boolean, i As Long, inputNumber As Long

'==================================================================================
'READ plc 1
'==================================================================================

  'Read digital inputs main (1-32)
  If Not IO_Modbus1 Is Nothing Then
    Dim DinpsRack1(1 To 32) As Boolean
    If IO_Modbus1.ReadBits(40001 + 16640, DinpsRack1, NoResponse) Then '416641
     If NoResponse Then
      Else
        ' Make sure we stay inbounds on both arrays
        For i = LBound(DinpsRack1) To UBound(DinpsRack1)
          inputNumber = i
          If inputNumber <= UBound(Dinp) Then Dinp(inputNumber) = DinpsRack1(i)
        Next i
        ACControlCode_ReadInputs = True
        PLC1WatchdogTimer = PLC1WatchdogTime
      End If
    End If
  End If

'Used to smooth analog inputs
  Static SmoothRate As Long
  SmoothRate = Parameters_SmoothRate
  Static Temp1 As New acSmoothing, Temp2 As New acSmoothing
  Static Temp3 As New acSmoothing, temp4 As New acSmoothing
  Static Temp5 As New acSmoothing
  Static Aninp1 As New acSmoothing, Aninp2 As New acSmoothing
  Static Aninp3 As New acSmoothing, Aninp4 As New acSmoothing
  Static Aninp5 As New acSmoothing, Aninp6 As New acSmoothing
  Static Aninp7 As New acSmoothing, Aninp8 As New acSmoothing
  Static Aninp9 As New acSmoothing, Aninp10 As New acSmoothing
  Static Aninp11 As New acSmoothing

'Read analog inputs in main
  If Not IO_Modbus1 Is Nothing Then
    Static Inputs(1 To 16) As Integer
   '8 inputs two words per input use only the first word
   '8 inputs =  V2040 = (1056 + 40001) = 41056
   
    If IO_Modbus1.ReadWords(40001 + 1056, Inputs, NoResponse) Then
      If NoResponse Then
          Aninp(1) = 0
          Aninp(2) = 0
          Aninp(3) = 0
          Aninp(4) = 0
          Aninp(5) = 0
          Aninp(6) = 0
          Aninp(7) = 0
          Aninp(8) = 0
      Else
      'Read, smooth and scale analog inputs
        IO_ANINP1Raw = ((Inputs(1) - 819) / (4095 - 819)) * 1000
        IO_ANINP2Raw = ((Inputs(3) - 819) / (4095 - 819)) * 1000
        IO_ANINP3Raw = ((Inputs(5) - 819) / (4095 - 819)) * 1000
        IO_ANINP4Raw = ((Inputs(7) - 819) / (4095 - 819)) * 1000
        IO_ANINP5Raw = ((Inputs(9) - 819) / (4095 - 819)) * 1000
        IO_ANINP6Raw = ((Inputs(11) - 819) / (4095 - 819)) * 1000
        IO_ANINP7Raw = ((Inputs(13) - 819) / (4095 - 819)) * 1000
        IO_ANINP8Raw = ((Inputs(15) - 819) / (4095 - 819)) * 1000
        IO_ANINP1Smooth = Aninp1.SmoothInt(IO_ANINP1Raw, SmoothRate)
        IO_ANINP2Smooth = Aninp2.SmoothInt(IO_ANINP2Raw, SmoothRate)
        IO_ANINP3Smooth = Aninp3.SmoothInt(IO_ANINP3Raw, SmoothRate)
        IO_ANINP4Smooth = Aninp4.SmoothInt(IO_ANINP4Raw, SmoothRate)
        IO_ANINP5Smooth = Aninp5.SmoothInt(IO_ANINP5Raw, SmoothRate)
        IO_ANINP6Smooth = Aninp6.SmoothInt(IO_ANINP6Raw, SmoothRate)
        IO_ANINP7Smooth = Aninp7.SmoothInt(IO_ANINP7Raw, SmoothRate)
        IO_ANINP8Smooth = Aninp8.SmoothInt(IO_ANINP8Raw, SmoothRate)
        IO_ANINP1Smooth = IO_ANINP1Smooth
        IO_ANINP2Smooth = IO_ANINP2Smooth
        IO_ANINP3Smooth = IO_ANINP3Smooth
        IO_ANINP4Smooth = IO_ANINP4Smooth
        IO_ANINP5Smooth = IO_ANINP5Smooth
        IO_ANINP6Smooth = IO_ANINP6Smooth
        IO_ANINP7Smooth = IO_ANINP7Smooth
        IO_ANINP8Smooth = IO_ANINP8Smooth
        Aninp(1) = IO_ANINP1Smooth
        Aninp(2) = IO_ANINP2Smooth
        Aninp(3) = IO_ANINP3Smooth
        Aninp(4) = IO_ANINP4Smooth
        Aninp(5) = IO_ANINP5Smooth
        Aninp(6) = IO_ANINP6Smooth
        Aninp(7) = IO_ANINP7Smooth
        Aninp(8) = IO_ANINP8Smooth
        
        If Parameters_TempTransmitterRange > 0 Then
          Dim Temp1InputRange As Long
          Dim Temp1Uncorrected As Long
          Dim Temp1TransmitterRange As Long '300.0F = 3000
          
          Temp1InputRange = MiniMaxi(Parameters_TempTransmitterMax, 900, 1000) - MiniMaxi(Parameters_TempTransmitterMin, 0, 100)
          Temp1Uncorrected = Aninp(7) - MiniMaxi(Parameters_TempTransmitterMin, 0, 100)
          Temp1TransmitterRange = MiniMaxi(Parameters_TempTransmitterRange, 2900, 5050)
        
          Temp(1) = (Temp1Uncorrected / Temp1InputRange) * Temp1TransmitterRange
                
        End If
      
      End If
    End If
  End If

'Read Temps in the control panel
  If Not IO_Modbus1 Is Nothing Then
    '4 temps uses two words per temp
    Static Temps(1 To 8) As Integer
        'Standard Ladder Logic Uses Memory Location V2070 octal = 1080 decimal (No Local PLC Noise Filtration)
        If IO_Modbus1.ReadWords(40001 + 1080, Temps, NoResponse) Then
          If NoResponse Then
            Temp(1) = 0
            Temp(2) = 0
            Temp(3) = 0
            Temp(4) = 0
          Else
          'Read and smooth temperatures
            IO_TEMP1FromPLC = Temps(1)
            IO_TEMP2FromPLC = Temps(3)
            IO_TEMP3FromPLC = Temps(5)
            IO_TEMP4FromPLC = Temps(7)
            IO_TEMP1Smooth = Temp1.SmoothInt(IO_TEMP1FromPLC, Parameters_SmoothRateRTD)
            IO_TEMP2Smooth = Temp2.SmoothInt(IO_TEMP2FromPLC, Parameters_SmoothRateRTD)
            IO_TEMP3Smooth = Temp3.SmoothInt(IO_TEMP3FromPLC, Parameters_SmoothRateRTD)
            IO_TEMP4Smooth = temp4.SmoothInt(IO_TEMP4FromPLC, Parameters_SmoothRateRTD)
            
            If Parameters_TempTransmitterRange = 0 Then
              'Only set Temp(1) = RTD value if we're not using a 4-20mA transmitter
              Temp(1) = IO_TEMP1Smooth
            End If

            Temp(2) = IO_TEMP2Smooth
            Temp(3) = IO_TEMP3Smooth
            Temp(4) = IO_TEMP4Smooth
          End If
        End If
    End If

  'Clear analog inputs here as a result of communication timeout.
  If Alarms_MainPLCNotResponding Then
    Aninp(1) = 0
    Aninp(2) = 0
    Aninp(3) = 0
    Aninp(4) = 0
    Aninp(5) = 0
    Aninp(6) = 0
    Temp(1) = 0
    Temp(2) = 0
    Temp(3) = 0
  End If

 
'==================================================================================
'READ plc 2
'==================================================================================

  'Read digital inputs drugroom (1-20)
  If Not IO_Modbus2 Is Nothing Then
    Dim DinpsRack2(1 To 20) As Boolean
    If IO_Modbus2.ReadBits(40001 + 16640, DinpsRack2, NoResponse) Then '416641
      If NoResponse Then
      Else
        ' Make sure we stay inbounds on both arrays
        For i = LBound(DinpsRack2) To UBound(DinpsRack2)
          inputNumber = i + 64
          If inputNumber <= UBound(Dinp) Then Dinp(inputNumber) = DinpsRack2(i)
        Next i
        ACControlCode_ReadInputs = True
        PLC2WatchdogTimer = PLC2WatchdogTime
      End If
    End If
  End If
  
  'Read analog inputs in drugroom
  If Not IO_Modbus2 Is Nothing Then
    Static InputsDrugroom(1 To 4) As Integer
   '4 inputs =  V2000 = (1024 + 40001) = 41025
    If IO_Modbus2.ReadWords(40001 + 1024, InputsDrugroom, NoResponse) Then
      If NoResponse Then
          Aninp(9) = 0
          Aninp(10) = 0
          Aninp(11) = 0
      Else
      'Read, smooth and scale analog inputs
        IO_ANINP9Raw = InputsDrugroom(1)
        IO_ANINP10Raw = InputsDrugroom(2)
        IO_ANINP11Raw = InputsDrugroom(3)
        IO_ANINP9Smooth = Aninp9.SmoothInt(IO_ANINP9Raw, SmoothRate)
        IO_ANINP10Smooth = Aninp10.SmoothInt(IO_ANINP10Raw, SmoothRate)
        IO_ANINP11Smooth = Aninp11.SmoothInt(IO_ANINP11Raw, SmoothRate)
        Aninp(9) = GetPercent(IO_ANINP9Smooth, 0, 4095)
        Aninp(10) = GetPercent(IO_ANINP10Smooth, 0, 4095)
        Aninp(11) = GetPercent(IO_ANINP11Smooth, 0, 4095)
      End If
    End If
  End If
 
  'Read temperatures in the drugroom panel
  If Not IO_Modbus2 Is Nothing Then
    Static TempsDrugroom(1 To 4) As Integer
    '1032 =2010 octal
    If IO_Modbus2.ReadWords(40001 + 1032, TempsDrugroom, NoResponse) Then
      If NoResponse Then
        Temp(5) = 0
      Else
      'Read and smooth temperatures
        IO_TEMP5FromPLC = TempsDrugroom(1)
        IO_TEMP5Smooth = Temp5.SmoothInt(IO_TEMP5FromPLC, SmoothRate)
        Temp(5) = IO_TEMP5Smooth
      End If
    End If
  End If

'Clear analog inputs here as a result of communication timeout.
  If Alarms_DrugroomPLCNotResponding Then
    Aninp(9) = 0
    Aninp(10) = 0
    Aninp(11) = 0
    Temp(5) = 0
  End If
    
End Function
Public Property Get Temperature() As String
  Temperature = VesTemp \ 10 & "F"
End Property

'==================================================================================
' WRITE OUTPUTS
'==================================================================================
Private Sub ACControlCode_WriteOutputs(Dout() As Boolean, Anout() As Integer)
'NoResponse used to flag comms failure, i used to iterate through arrays
  Dim NoResponse As Boolean, i As Long
    
  '==================================================================================
  'WRITE plc 1
  '==================================================================================
    
  'If the system is stopping or watchdog has timed out then turn all outputs off.
  If SystemShuttingDown Or PLC1WatchdogTimer.Finished Then
    For i = 1 To 62: Dout(i) = False: Next i
    For i = LBound(Anout) To UBound(Anout): Anout(i) = 0: Next i
  End If

  'Digital Outputs 1-12 in main
  Dim DoutMain(1 To 62) As Boolean
  For i = 1 To 6: DoutMain(i) = Dout(i): Next i                 'Card 1: group A (1-6)
  For i = 9 To 14: DoutMain(i) = Dout(i - 2): Next i            'Card 1: group B (7-12)
  For i = 17 To 22: DoutMain(i) = Dout(i - 4): Next i           'Card 2: group A (13-18)
  For i = 25 To 30: DoutMain(i) = Dout(i - 6): Next i           'Card 2: group B (19-24)
  For i = 33 To 38: DoutMain(i) = Dout(i - 8): Next i           'Card 3: group A (25-30)
  For i = 41 To 46: DoutMain(i) = Dout(i - 10): Next i          'Card 3: group B (31-36)
  For i = 49 To 54: DoutMain(i) = Dout(i - 12): Next i          'Card 4: group A (37-42)
  For i = 57 To 62: DoutMain(i) = Dout(i - 14): Next i          'Card 4: group B (43-48)
  
  'Write PLC 1 output array
  If IO_Modbus1.WriteBits(40001 + 16704, DoutMain, NoResponse) Then
    If NoResponse Then
    End If
  End If
  
  'Write PLC 1 watch dog
  Dim plc1watchdogoutput(1 To 1) As Boolean
  plc1watchdogoutput(1) = Dout(57)
  If IO_Modbus1.WriteBits(40001 + 16708, plc1watchdogoutput, NoResponse) Then
    If NoResponse Then
    End If
  End If

 
'==================================================================================
'Analog Outputs Plc1 8 outputs
'==================================================================================
  'Scale analog outputs (0-1000 =4mA-20mA=819-4095)
  Dim tempAnout(1 To 8) As Integer
  'only four outputs but uses two words for each output.
  '1072 dec is 2060 octal
  'output 1
  TempANOUT1 = ((Anout(1) / 1000) * 65535)
  If TempANOUT1 < 0 Then TempANOUT1 = 0
  If TempANOUT1 > 65535 Then TempANOUT1 = 65535
  'If (Not Dout(7)) Then TempANOUT1 = 0
  tempAnout(1) = TempANOUT1
  'output 2
  TempANOUT2 = ((Anout(2) / 1000) * 65535)
  If TempANOUT2 < 0 Then TempANOUT2 = 0
  If TempANOUT2 > 65535 Then TempANOUT2 = 65535
  tempAnout(3) = TempANOUT2
  'output 3
  TempANOUT3 = ((Anout(3) / 1000) * 65535)
  If TempANOUT3 < 0 Then TempANOUT3 = 0
  If TempANOUT3 > 65535 Then TempANOUT3 = 65535
  tempAnout(5) = TempANOUT3
  'output 4
  TempANOUT4 = ((Anout(4) / 1000) * 65535)
  If TempANOUT4 < 0 Then TempANOUT4 = 0
  If TempANOUT4 > 65535 Then TempANOUT4 = 65535
  tempAnout(7) = TempANOUT4
  If Not IO_Modbus1 Is Nothing Then
    If IO_Modbus1.WriteWords(40001 + 1072, tempAnout, NoResponse) Then
    End If
  End If
  
  
'==================================================================================
'WRITE plc 2
'==================================================================================
  
  'If the system is stopping or watchdog has timed out then turn all outputs off.
  If SystemShuttingDown Or PLC2WatchdogTimer.Finished Then
    For i = 65 To 80: Dout(i) = False: Next i
    For i = LBound(Anout) To UBound(Anout): Anout(i) = 0: Next i
  End If

  'Digital Outputs 65 to 80
  Dim DoutDrugroom(65 To 80) As Boolean
  For i = 65 To 80: DoutDrugroom(i) = Dout(i): Next i
  If IO_Modbus2.WriteBits(40001 + 16704, DoutDrugroom, NoResponse) Then
    If NoResponse Then
    End If
  End If

  'Write plc 2 watch dog
  Dim plc2watchdogoutput(1 To 1) As Boolean
  plc2watchdogoutput(1) = Dout(81)
  If IO_Modbus2.WriteBits(40001 + 16705, plc2watchdogoutput, NoResponse) Then
    If NoResponse Then
    End If
  End If
  
End Sub

Private Sub Alarms_MakeAlarms()

'Safety temp error
  Alarms_TemperatureTooHigh = (IO_ContactThermometer) Or (VesTemp > 2800)
  
'Temperature probe errors
  Alarms_VesselTempProbeError = IO_VesselTemp < 320 Or IO_VesselTemp > 3000
  Alarms_BlendFillTempProbeError = IO_BlendFillTemp < 320 Or IO_BlendFillTemp > 3000
  Alarms_ReserveTempProbeError = IO_ReserveTemp < 320 Or IO_ReserveTemp > 3000
   
'Condensate Return OverTemp Alarm
  Static condensateAlarmTimer As New acTimer
  If (IO_CondensateTemp > 320) And (IO_CondensateTemp < Parameters_TempCondensateLimit) Then condensateAlarmTimer = Parameters_TempCondensateLimitTime
  Alarms_CondensateTempTooHigh = condensateAlarmTimer.Finished And (Parameters_TempCondensateLimit > 0) And (Parameters_TempCondensateLimitTime > 0)
  
'Test Value Input - User Scaled Spare Analog input w/ parameters for signalling over limit
  Static testInputAlarmTimer As New acTimer
  If (TestValue < Parameters_TestValueLimit) Then testInputAlarmTimer = Parameters_TestValueLimitTime
  Alarms_TestValueTooHigh = testInputAlarmTimer.Finished And (Parameters_TestValueLimit > 0) And (Parameters_TestValueLimitTime > 0)
    
'Main pump tripped
  Static PumpAlarmTimer As New acTimer
  If (Not IO_MainPumpStart) Or IO_MainPumpRunning Then PumpAlarmTimer = 5
  Alarms_MainPumpTripped = PumpAlarmTimer.Finished
   
'add pump tripped
  Static AddPumpAlarmTimer As New acTimer
  If (Not IO_AddPumpStart) Or IO_AddPumpRunning Then AddPumpAlarmTimer = 5
  Alarms_AddPumpTripped = AddPumpAlarmTimer.Finished
  
'Lid Locked
  If Parent.IsProgramRunning And PumpRequest And Not (SA.IsSampling Or LD.IsOn Or UL.IsUnloading) And Not LidLocked Then
    PumpRequest = False
    Alarms_LidNotLocked = True
  End If
  If Not (IO_RemoteRun And LidLocked) Then LidLockTimer.TimeRemaining = 2
  If LidLockTimer.Finished Then
    If Parent.IsProgramRunning And (PumpLevel Or DR.IsInterlocked Or DR.IsDraining Or RT.IsTransfer) Then PumpRequest = True
    Alarms_LidNotLocked = False
  End If

'level to low
  Alarms_LevelToLowForPump = PumpRequest And (VesselLevel < Parameters_PumpMinimumLevel) And _
                            Not (RI.IsActive Or RH.IsActive Or DR.IsActive Or HD.IsActive Or RT.IsActive)
  
'Estop alarm
  Alarms_EmergencyStopPB = IO_EStop_PB
  
'Power problems?
  Alarms_ControlPowerFailure = Not IO_PowerOn
  
'Comms Problems?
  Alarms_MainPLCNotResponding = PLC1WatchdogTimer.Finished And (Parent.Mode <> modeDebug)
  Alarms_DrugroomPLCNotResponding = PLC2WatchdogTimer.Finished And (Parent.Mode <> modeDebug)
  
'Manual override modes
  Alarms_TestModeSelected = (Parent.Mode = modeTest)
  Alarms_OverrideModeSelected = (Parent.Mode = modeOverride)
  Alarms_DebugModeSelected = (Parent.Mode = modeDebug)
   
'vent valve not open after five mins on hot drain
  Alarms_VentNotOpenOnHotDrain = (HD.IsActive And HD.VentAlarmTimer.Finished And Not IO_Vent)

'lid alarms
  Alarms_LidNotClosed = (LidControl.IsLowerLid) And LidControl.Timer.Finished _
                        And Not IO_LidLoweredSwitch
                        
  Static AddLevelTooHighTimer As New acTimer
  If Not KA.IsInterlocked Then AddLevelTooHighTimer = 5
  Alarms_LevelTooHighInAddTank = AddLevelTooHighTimer.Finished
  
  If GetRecordedLevel Then
     If VesselLevel <= (RecordedLevel - Parameters_MaxLevelLossAmount) Then
        Alarms_LosingLevelInTheVessel = True
     Else
        Alarms_LosingLevelInTheVessel = False
     End If
  End If

'Level transmitter
 Alarms_CheckLevelTransmitter = (DR.IsActive And DR.LevelGaugeFlushCount > 5) Or _
                                (HD.IsActive And HD.LevelGaugeFlushCount > 5)


'pump on off alarms
  Alarms_MainPumpTurnedOnOffTooManyTimes = PumpOnCountTooHigh
  
'reserve tank high temp alarm
  Alarms_ReserveTempTooHigh = IO_ReserveTemp > Parameters_ReserveTankHighTempLimit
  
'temperature to high and low
  Alarms_TemperatureHigh = TemperatureControl.TemperatureHigh Or TemperatureControlContacts.TemperatureHigh _
                          Or RC.TemperatureHigh Or RH.TemperatureHigh Or RI.TemperatureHigh
  Alarms_TemperatureLow = TemperatureControl.TemperatureLow Or TemperatureControlContacts.TemperatureLow _
                          Or RC.TemperatureLow Or RH.TemperatureLow Or RI.TemperatureLow
                                                        
  Alarms_TooLongToHeat = HE.IsOverrun Or TC.IsOverrun
                            
'Low Flow Rate
  Dim disregardFlow_ As Boolean
  disregardFlow_ = DR.IsOn Or FI.IsOn Or FR.IsReversing Or FC.IsReversing Or RT.IsOn Or _
                   RI.IsFilling Or RH.IsFilling Or RW.IsOn Or _
                   (Not IO_MainPumpRunning)
  
  If disregardFlow_ Or ((FlowRatePercent / 10) >= Parameters_LowFlowRatePercent) Then
     LowFlowRateTimer = Parameters_LowFlowRateTime
  End If
  Alarms_LowFlowRate = LowFlowRateTimer.Finished

'Top Wash Alarm
  Alarms_TopWashBlocked = (RI.IsRinsing And RI.TopWashBlocked) Or (RH.IsRinsing And RH.TopWashBlocked)

'DRUGROOM ALARMS:
'Drugroom Temp Input is occasionally dropping to 0 for a second, resulting in alarm....
  Static TempProbeTimer As New acTimer
  If (IO_Tank1Temp >= 320) Or (IO_Tank1Temp <= 3000) Then TempProbeTimer = 5
  Alarms_Tank1TempProbeError = (IO_Tank1Temp < 320 Or IO_Tank1Temp > 3000) And TempProbeTimer.Finished
  
  Alarms_Tank1TooLongToHeat = KP.KP1.IsHeatTankOverrun Or LA.KP1.IsHeatTankOverrun
  Alarms_Tank1TempTooHigh = IO_Tank1Temp > Parameters_Tank1HighTempLimit

  If KP.Tank1DispenseError Or LA.Tank1DispenseError Then
     Alarms_Tank1DispenseError = True
  Else
     Alarms_Tank1DispenseError = False
  End If

'Alarms for auto-dispenser ready and dispenser response
  Alarms_AutoDispenserDelay = ((Parameters_DispenseEnabled = 1) And (Parameters_DispenseReadyDelayTime <> 0)) And _
                              (KP.KP1.IsDispenseReadyOverrun Or LA.KP1.IsDispenseReadyOverrun)
  
  Alarms_DispenserCompleteDelay = ((Parameters_DispenseEnabled = 1) And (Parameters_DispenseResponseDelayTime <> 0)) And _
                              (KP.KP1.IsDispenseResponseOverrun Or LA.KP1.IsDispenseResponseOverrun)
  
'Alarm for @1 bug
  Alarms_RedyeIssueCheckMachineTank = (Parameters_EnableRedyeIssueAlarm = 1) And (KP.KP1.AlarmRedyeIssue)
  
End Sub
Public Sub Parameters_MaybeSetDefaults()
  If (Parameters_InitializeParameters = 9387) Or _
     (Parameters_InitializeParameters = 8741) Or _
     (Parameters_InitializeParameters = 8742) Or _
     (Parameters_InitializeParameters = 8743) Or _
     (Parameters_InitializeParameters = 8744) Then
    Parameters_SetDefaults
  End If
End Sub

Public Sub Parameters_SetDefaults()
'Safety
  SafetyControl.Parameters.PressurizationTemperature = 1900
  SafetyControl.Parameters.DePressurizationTemperature = 1850
  SafetyControl.Parameters.DePressurizationTime = 30
  
'Temperature Control
  TemperatureControl.Parameters.HeatExitDeadband = 20
  TemperatureControl.Parameters.HeatPropBand = 100
  TemperatureControl.Parameters.HeatIntegral = 250
  TemperatureControl.Parameters.HeatMaxGradient = 150
  TemperatureControl.Parameters.HeatStepMargin = 10
  TemperatureControl.Parameters.CoolPropBand = 50
  TemperatureControl.Parameters.CoolIntegral = 250
  TemperatureControl.Parameters.CoolMaxGradient = 100
  TemperatureControl.Parameters.CoolStepMargin = 10
  TemperatureControl.Parameters.BalanceTemperature = 750
  TemperatureControl.Parameters.BalPercent = 100
  
  TemperatureControl.Parameters.HeatCoolModeChange = 0
  TemperatureControl.Parameters.HeatCoolModeChangeDelay = 60
  TemperatureControl.Parameters.TemperatureAlarmBand = 90
  TemperatureControl.Parameters.TemperatureAlarmDelay = 120
  TemperatureControl.Parameters.TemperaturePidPause = 200
  TemperatureControl.Parameters.TemperaturePidRestart = 30
  TemperatureControl.Parameters.TemperaturePidReset = 250
  
  TemperatureControl.Parameters.HeatVentTime = 10
  TemperatureControl.Parameters.CoolVentTime = 10
  
'Motor Control
  Parameters_PumpMinimumLevelTime = 1
  
'Blend Control
  Parameters_BlendFactor = 30
  Parameters_BlendDeadBand = 10
  Parameters_HotWaterTemperature = 1700
  Parameters_ColdWaterTemperature = 550
  
'Command Specific
  Parameters_DRDrainTime = 60
  Parameters_InitializeParameters = 0
  
'SETUP to Improve Startup Efficiency:
  Parameters_AddMixOffLevel = 100
  Parameters_AddMixOnLevel = 150
  Parameters_AddRinseFillLevel = 200
  Parameters_AddTransferRinseTime = 1
  Parameters_AddTransferRinseToDrainTime = 2
  Parameters_AddTransferTimeAfterRinse = 10
  Parameters_AddTransferTimeBeforeRinse = 10
  Parameters_AddTransferToDrainTime = 12
  Parameters_AdditionLevelMax = 100
  Parameters_AdditionMaxTransferLevel = 200
  Parameters_AdditionFillDeadband = 30
  Parameters_AddRinseMixPulseTime = 20
  Parameters_AddRinseMixTime = 120
  
  Parameters_AirpadBleedPressure = 550
  Parameters_AirpadCutoffPressure = 350
  
  Parameters_BlendDeadBand = 10
  Parameters_BlendFactor = 30
  Parameters_ColdWaterTemperature = 550
  Parameters_HotWaterTemperature = 1500
  
  Parameters_DPAdjustFactor = 200
  Parameters_DPAdjustMaximum = 100
  Parameters_MaxDifferentialPressure = 150
  
  Parameters_CKFillLevel = 800
  Parameters_CKRinseTime = 3
  Parameters_CKTransferToDrainTime1 = 20
  Parameters_CKTransferToDrainTime2 = 40
  Parameters_DrugroomRinses = 1
  Parameters_Tank1DrainTime = 10
  Parameters_Tank1FillLevelDeadBand = 30
  Parameters_Tank1HeatDeadband = 30
  Parameters_Tank1MixerOffLevel = 350
  Parameters_Tank1MixerOnLevel = 400
  Parameters_Tank1RinseLevel = 200
  Parameters_Tank1RinseTime = 3
  Parameters_Tank1RinseToDrainTime = 15
  Parameters_Tank1TimeAfterRinse = 10
  Parameters_Tank1TimeBeforeRinse = 10
  Parameters_Tank1HeatPrepTimer = 300
  Parameters_Tank1HighTempLimit = 1950
  
  Parameters_Tank1TempPotMax = 1000
  Parameters_Tank1TempPotMin = 10
  Parameters_Tank1TempPotRange = 1800
  Parameters_Tank1TimePotMax = 1000
  Parameters_Tank1TimePotMin = 10
  Parameters_Tank1TimePotRange = 1800
  
  Parameters_FillPrimePumpTime = 10
  Parameters_FillSettleTime = 10
  Parameters_MaxLevelLossAmount = 150
  Parameters_RecordLevelTimer = 180
  
  Parameters_FLAdjustFactor = 20
  Parameters_FLAdjustMaximum = 50
  Parameters_FlowRateMax = 1000
  Parameters_FlowRateMin = 10
  Parameters_FlowRateRange = 107
  Parameters_LowFlowRatePercent = 12
  Parameters_LowFlowRateTime = 30
  Parameters_SystemVolumeAt0Level = 40
  Parameters_SystemVolumeAt100Level = 330
  
  Parameters_AdditionLevelMax = 555
  Parameters_AdditionLevelMin = 10
  Parameters_ReserveLevelMax = 870
  Parameters_ReserveLevelMin = 10
  Parameters_Tank1LevelMax = 350
  Parameters_Tank1LevelMin = 10
  Parameters_VesselLevelMax = 1000
  Parameters_VesselLevelMin = 10
  
  Parameters_LevelOkToOpenLid = 750
  Parameters_LidRaisingTime = 10
  Parameters_LockingBandClosingTime = 5
  Parameters_LockingBandOpeningTime = 5
  
  Parameters_DRDrainTime = 45
  Parameters_HDBlendFillPosition = 750
  Parameters_HDDrainTime = 45
  Parameters_HDHiFillLevel = 250
  Parameters_HDVentAlarmTime = 600
  Parameters_LevelGaugeFlushTime = 5
  Parameters_LevelGaugeSettleTime = 5
  
  Parameters_PackageDiffPresMax = 1000
  Parameters_PackageDiffPresMin = 495
  Parameters_PackageDiffPresRange = 362
  Parameters_VesselPressureMax = 1000
  Parameters_VesselPressureMin = 10
  Parameters_VesselPressureRange = 1450
    
  Parameters_StandardOperatorTime = 180
  
  Parameters_DrainPumpSpeed = 500
  Parameters_FillSettlePumpSpeed = 500
  Parameters_FlowReverseTime = 10
  Parameters_PumpAccelerationTime = 15
  Parameters_PumpMinimumLevel = 10
  Parameters_PumpMinimumLevelTime = 1
  Parameters_PumpOnOffCountTooHigh = 5
  Parameters_PumpReversalSpeed = 350
  Parameters_PumpSpeedDecelTime = 15
  Parameters_PumpSpeedDefault = 500
  Parameters_PumpSpeedStart = 500
  Parameters_RinsePumpSpeed = 500
  
  Parameters_ReserveDrainTime = 15
  Parameters_ReserveHeatDeadband = 50
  Parameters_ReserveMixerOffLevel = 350
  Parameters_ReserveMixerOnLevel = 400
  Parameters_ReserveRinseTime = 5
  Parameters_ReserveRinseToDrainTime = 10
  Parameters_ReserveTimeAfterRinse = 7
  Parameters_ReserveTimeBeforeRinse = 4
  Parameters_ReserveTankHighTempLimit = 1950
  
  Parameters_RinseTemperatureAlarmBand = 150
  
  Parameters_InitializeParameters = 0
  Parameters_SmoothRate = 10
  
End Sub
Private Sub Delays_MakeDelays()
'Handy routine for setting variable

  Const ACNormalRunning As Long = 0
  Const ACDelayBoilout As Long = 1
  Const ACDelayCorrection As Long = 2
  Const ACDelayStepOverrun As Long = 3
  Const ACDelaySample As Long = 4
  Const ACDelayLoad As Long = 5
  Const ACDelayUnload As Long = 6
  Const ACDelayCheckPH As Long = 7
  Const ACDelayOperator As Long = 8
  Const ACDelayFill As Long = 9
  Const ACDelayDrain As Long = 10
  Const ACDelayHeating As Long = 11
  Const ACDelayCooling As Long = 12
  Const ACDelayPaused As Long = 13
  Const ACDelayMachineError As Long = 14
  Const ACDelayExpansionAdd As Long = 15
  Const ACDelayTank1Prepare As Long = 16
  Const ACDelayReserveTransfer As Long = 17
  Const ACDelayDispenseWaitReady As Long = 18
  Const ACDelayDispenseWaitResultDyes As Long = 19
  Const ACDelayDispenseWaitResultChems As Long = 20
  Const ACDelayDispenseWaitResultCombined As Long = 21
  Const ACDelayLocalPrepare As Long = 22
  Const ACDelayTank1Operator As Long = 23
  Const ACDelayTank1DispenseError As Long = 24
   
  Const ACSleeping As Long = 101

'Set delay to 0 = Normal Running
  Delay = ACNormalRunning

' ph check delay
  If PH.IsOverrun Then Delay = ACDelayCheckPH
  
'Step overrun stuff
  If RI.IsOverrun Then Delay = ACDelayStepOverrun
  If TM.IsOverrun Then Delay = ACDelayStepOverrun

'Reserve Tank Delay Stuff
  'RF using runback is a foreground function - so RT will not be on (waitready or otherwise)
  If RF.IsForeground And RF.IsOverrun Then Delay = ACDelayLocalPrepare
  
  'RT delay due to:
  ' - reserve mechanical <fill/heat/mix>
  ' - tank 1 dispenser <which one>
  ' - tank 1 mechanical <fill/heat/mix>
  ' - tank 1 operator <due to dispenser error or manual>
  If RT.IsDelayedWaitReady Then
    'Only signal a delay for a background function associated with RT if the RT is actively waiting after a 5min delay period
    If RP.IsWaitReady Then Delay = ACDelayOperator                                 'waiting for local operator to signal tank ready
    If RF.IsOverrun Then Delay = ACDelayLocalPrepare                               'waiting for local tank to fill & heat - charged against operator
    If KA.IsActive And (KA.Destination = 82) Then Delay = ACDelayTank1Operator     'KA command waiting for tank 1 operator to signal tank ready
    If KP.KP1.IsOn And (KP.KP1.DispenseTank = 2) Then
      With KP
        If .KP1.IsDispenseWaitReady Then Delay = ACDelayDispenseWaitReady          'waiting for autodispenser to signal ready
        If .KP1.IsDispenseWaitResponse Then
          If .KP1.DispenseChemsOnly Then
            Delay = ACDelayDispenseWaitResultChems                                 'waiting for response from chemical dispenser
          ElseIf .KP1.DispenseDyesOnly Then
            Delay = ACDelayDispenseWaitResultDyes                                  'waiting for response from dye dispenser
          Else
            'Verify that DyeDispenser is enabled - if not and combined, must be waiting for chemical dispenser
            If (Parameters_DDSEnabled <> 1) Then
              Delay = ACDelayDispenseWaitResultChems                               'waiting for response from chemical dispenser
            Else
              Delay = ACDelayDispenseWaitResultCombined                            'dyes & chems in drop - charged against combined delay (non specific)
            End If
          End If
        End If
        If .KP1.IsHeatTankOverrun Then Delay = ACDelayTank1Prepare                 'tank 1 too long to fill & heat (regardless of dispense error)
        If .KP1.IsSlow Or .KP1.IsFast Then                                         'Don't use IsWaitReady = don't want to include tank mix time against operator
          If .KP1.AddDispenseError Then
            Delay = ACDelayTank1DispenseError                                      'tank 1 operator too long to signal ready - after dispense error
          Else: Delay = ACDelayTank1Operator                                       'tank 1 operator too long to signal ready - manual drop
          End If
        End If
        If (KP.KP1.IsInterlocked) Then Delay = ACDelayTank1Prepare                 'tank 1 cannot transfer due to destination too high
        'remaining tank 1 transfer is not considered a delay
      End With
    End If
    If LA.KP1.IsOn And (LA.KP1.DispenseTank = 2) Then
      With LA
        If .KP1.IsDispenseWaitReady Then Delay = ACDelayDispenseWaitReady          'waiting for autodispenser to signal ready
        If .KP1.IsDispenseWaitResponse Then
          If .KP1.DispenseChemsOnly Then
            Delay = ACDelayDispenseWaitResultChems                                 'waiting for response from chemical dispenser
          ElseIf .KP1.DispenseDyesOnly Then
            Delay = ACDelayDispenseWaitResultDyes                                  'waiting for response from dye dispenser
          Else
            'Verify that DyeDispenser is enabled - if not and combined, must be waiting for chemical dispenser
            If (Parameters_DDSEnabled <> 1) Then
              Delay = ACDelayDispenseWaitResultChems                               'waiting for response from chemical dispenser
            Else
              Delay = ACDelayDispenseWaitResultCombined                            'dyes & chems in drop - charged against combined delay (non specific)
            End If
          End If
        End If
        If .KP1.IsHeatTankOverrun Then Delay = ACDelayTank1Prepare                 'tank 1 too long to fill & heat (regardless of dispense error)
        If .KP1.IsSlow Or .KP1.IsFast Then                                         'Don't use IsWaitReady = don't want to include tank mix time against operator
          If .KP1.AddDispenseError Then
            Delay = ACDelayTank1DispenseError                                      'tank 1 operator too long to signal ready - after dispense error
          Else: Delay = ACDelayTank1Operator                                       'tank 1 operator too long to signal ready - manual drop
          End If
        End If
        If (KP.KP1.IsInterlocked) Then Delay = ACDelayTank1Prepare                 'tank 1 cannot transfer due to destination too high
        'remaining tank 1 transfer is not considered a delay
      End With
    End If
  Else
    'this delay is associated with mechanical issues where tank doesn't transfer in specified parameter time
    If RT.IsDelayedTransfer Then Delay = ACDelayReserveTransfer
  End If

'Addition Tank Delay Stuff
  'AT/AC delay due to:
  ' - addition mechanical <fill/heat/mix>
  ' - tank 1 dispenser <which one>
  ' - tank 1 mechanical <fill/heat/mix>
  ' - tank 1 operator <due to dispenser error or manual>
  If AT.IsDelayedWaitReady Or AC.IsDelayedWaitReady Then
    If AP.IsWaitReady Then Delay = ACDelayLocalPrepare
    If AF.IsActive And (AF.State <> AFState.Fill) Then Delay = ACDelayLocalPrepare             'waiting for local tank to fill - charged against operator
    If KA.IsActive And (KA.Destination = 65) Then Delay = ACDelayTank1Operator        'KA command waiting for tank 1 operator to signal tank ready
    If KP.KP1.IsOn And (KP.KP1.DispenseTank = 1) Then
      With KP
        If .KP1.IsDispenseWaitReady Then Delay = ACDelayDispenseWaitReady           'waiting for autodispenser to signal ready
        If .KP1.IsDispenseWaitResponse Then
          If .KP1.DispenseChemsOnly Then
            Delay = ACDelayDispenseWaitResultChems                                  'waiting for response from chemical dispenser
          ElseIf .KP1.DispenseDyesOnly Then
            Delay = ACDelayDispenseWaitResultDyes                                   'waiting for response from dye dispenser
          Else
            'Verify that DyeDispenser is enabled - if not and combined, must be waiting for chemical dispenser
            If (Parameters_DDSEnabled <> 1) Then
              Delay = ACDelayDispenseWaitResultChems                                'waiting for response from chemical dispenser
            Else
              Delay = ACDelayDispenseWaitResultCombined                             'dyes & chems in drop - charged against combined delay (non specific)
            End If
          End If
        End If
        If .KP1.IsHeatTankOverrun Then Delay = ACDelayTank1Prepare                  'tank 1 too long to fill & heat (regardless of dispense error)
        If .KP1.IsSlow Or .KP1.IsFast Then                                          'Don't use IsWaitReady = don't want to include tank mix time against operator
          If .KP1.AddDispenseError Then
            Delay = ACDelayTank1DispenseError                                       'tank 1 operator too long to signal ready - after dispense error
          Else: Delay = ACDelayTank1Operator                                        'tank 1 operator too long to signal ready - manual drop
          End If
        End If
        If (.KP1.IsInterlocked) Then Delay = ACDelayTank1Prepare                    'tank 1 cannot transfer due to destination too high
        'remaining tank 1 transfer is not considered a delay
      End With
    End If
    If LA.KP1.IsOn And (LA.KP1.DispenseTank = 1) Then
      With LA
        If .KP1.IsDispenseWaitReady Then Delay = ACDelayDispenseWaitReady           'waiting for autodispenser to signal ready
        If .KP1.IsDispenseWaitResponse Then
          If .KP1.DispenseChemsOnly Then
            Delay = ACDelayDispenseWaitResultChems                                  'waiting for response from chemical dispenser
          ElseIf .KP1.DispenseDyesOnly Then
            Delay = ACDelayDispenseWaitResultDyes                                   'waiting for response from dye dispenser
          Else
            'Verify that DyeDispenser is enabled - if not and combined, must be waiting for chemical dispenser
            If (Parameters_DDSEnabled <> 1) Then
              Delay = ACDelayDispenseWaitResultChems                                'waiting for response from chemical dispenser
            Else
              Delay = ACDelayDispenseWaitResultCombined                             'dyes & chems in drop - charged against combined delay (non specific)
            End If
          End If
        End If
        If .KP1.IsHeatTankOverrun Then Delay = ACDelayTank1Prepare                  'tank 1 too long to fill & heat (regardless of dispense error)
        If .KP1.IsSlow Or .KP1.IsFast Then                                          'Don't use IsWaitReady = don't want to include tank mix time against operator
          If .KP1.AddDispenseError Then
            Delay = ACDelayTank1DispenseError                                       'tank 1 operator too long to signal ready - after dispense error
          Else: Delay = ACDelayTank1Operator                                        'tank 1 operator too long to signal ready - manual drop
          End If
        End If
        If (.KP1.IsInterlocked) Then Delay = ACDelayTank1Prepare                    'tank 1 cannot transfer due to destination too high
        'remaining tank 1 transfer is not considered a delay
      End With
    End If
  Else
    'no delay associated with addition tank mechanical transfer issues
  End If
        
'Look at Sample and Begin sample command for sample delays
  If SA.IsOverrun Then Delay = ACDelaySample
  
'Look at Load command for load delays
' If not running a full week, assume startup on sunday, and dsregard LD delays on sunday
'   (allen michael says always assume sunday startup and disregard rare holiday startups)
  If Parameters_SevenDaySchedule <> 1 Then
    'use function to return day of week and if sunday (0) then disregard delay
    If Int(DOW(Now)) <> 0 Then
      If LD.IsOverrun Then Delay = ACDelayLoad
    End If
  Else
    'Running seven days a week, always record load delays
    If LD.IsOverrun Then Delay = ACDelayLoad
  End If

'Look at Unload command for unload
  If UL.IsOverrun Then Delay = ACDelayUnload

'Look at Fill command for delays
  If FI.IsOverrun Then Delay = ACDelayFill

'Look at Drain Command for delays
  If DR.IsOverrun Then Delay = ACDelayDrain
  If HD.IsOverrun Then Delay = ACDelayDrain
  
'Look at heating
  If HE.IsOverrun Then Delay = ACDelayHeating

'Look at cooling
  If CO.IsOverrun Then Delay = ACDelayCooling

'Look at rinsing
  If RI.TopWashBlocked And RI.IsRinsing Then Delay = ACDelayMachineError
  If RH.TopWashBlocked And RH.IsRinsing Then Delay = ACDelayMachineError
  
'operator delay
  If Parent.ButtonText = "" Then OperatorDelayTimer = Parameters_StandardOperatorTime
  If OperatorDelayTimer.Finished Then Delay = ACDelayOperator
  
'Look at if control is paused
  If Parent.IsPaused Then Delay = ACDelayPaused

'Look at alarms and set machine delay
  If Alarms_BlendFillTempProbeError Then Delay = ACDelayMachineError
  If Alarms_ControlPowerFailure Then Delay = ACDelayMachineError
'  If Alarms_DebugModeSelected Then Delay = ACDelayMachineError
  If Alarms_EmergencyStopPB Then Delay = ACDelayMachineError
  If Alarms_LidNotLocked Then Delay = ACDelayMachineError
'  If Alarms_OverrideModeSelected Then Delay = ACDelayMachineError
  If Alarms_MainPLCNotResponding Then Delay = ACDelayMachineError
  If Alarms_DrugroomPLCNotResponding Then Delay = ACDelayMachineError
  If Alarms_MainPumpTripped Then Delay = ACDelayMachineError
  If Alarms_VesselTempProbeError Then Delay = ACDelayMachineError
  If Alarms_TemperatureTooHigh Then Delay = ACDelayMachineError
  If Alarms_TestModeSelected Then Delay = ACDelayMachineError
  
'Look at Begin Reprocess command for reprocessing (corrective add) delays
  If BR.IsOn Then Delay = ACDelayCorrection

'Look at boilout command for Boilout Delay
  If BO.IsOn Then Delay = ACDelayBoilout

'Look at sleeping flag and set delay
  If Parent.IsSleeping Then Delay = ACSleeping
  
'Set Boolean Delays
  Delay_NormalRunning = (Delay = ACNormalRunning)
  Delay_Boilout = (Delay = ACDelayBoilout)
  Delay_Correction = (Delay = ACDelayCorrection)
  Delay_Sample = (Delay = ACDelaySample)
  Delay_Load = (Delay = ACDelayLoad)
  Delay_Unload = (Delay = ACDelayUnload)
  Delay_CheckPH = (Delay = ACDelayCheckPH)
  Delay_Operator = (Delay = ACDelayOperator)
  Delay_Fill = (Delay = ACDelayFill)
  Delay_Drain = (Delay = ACDelayDrain)
  Delay_Heating = (Delay = ACDelayHeating)
  Delay_Cooling = (Delay = ACDelayCooling)
  Delay_Paused = (Delay = ACDelayPaused)
  Delay_MachineError = (Delay = ACDelayMachineError)
  Delay_ExpansionAdd = (Delay = ACDelayExpansionAdd)
  Delay_Tank1Prepare = (Delay = ACDelayTank1Prepare)
  Delay_ReserveTransfer = (Delay = ACDelayReserveTransfer)
  Delay_DispenseWaitReady = (Delay = ACDelayDispenseWaitReady)
  Delay_DispenseWaitResultDyes = (Delay = ACDelayDispenseWaitResultDyes)
  Delay_DispenseWaitResultChems = (Delay = ACDelayDispenseWaitResultChems)
  Delay_DispenseWaitResultCombined = (Delay = ACDelayDispenseWaitResultCombined)
  Delay_LocalPrepare = (Delay = ACDelayLocalPrepare)
  Delay_Tank1Operator = (Delay = ACDelayTank1Operator)
  Delay_Tank1DispenseError = (Delay = ACDelayTank1DispenseError)
    
  Delay_Sleeping = (Delay = ACSleeping)
  Delay_IsDelayed = (Delay <> 0)

End Sub

'Added 2009/06/15 to use with Parameters_SevenDaySchedule to disregard Load Delays upon sunday night startup
' requested by allen-michael due to delay compensation for initial week startup delays
' a&e schedules a machine unblocked at startup, and the operator may take as long as 120minutes to actually load and complete
Public Function DOW(ByVal GregDate As Date) As String
' Return values:
' 0 = Sunday
' 1 = Monday
' 2 = Tuesday
' 3 = Wednesday
' 4 = Thursday
' 5 = Friday
' 6 = Saturday
    Dim y As Integer
    Dim m As Integer
    Dim d As Integer
    
    ' monthdays:
    ' This is a "template" for a year. Each number
    ' stands for a day of the week. The general idea
    ' is that, in a standard year, if Jan 1 is on a
    ' Friday, then Feb 1 will be a Monday, Mar 1
    ' will be a Monday, April 1 will be a Thursday,
    ' May 1 will be Saturday, etc..
    Dim mcode As String
    Dim monthdays() As String
    monthdays = Split("5 1 1 4 6 2 4 0 3 5 1 3")
    
    ' Grab our date info
    y = val(Format(GregDate, "yyyy"))
    m = val(Format(GregDate, "mm"))
    d = val(Format(GregDate, "dd"))
    
    ' Snatch the corresponding month code
    mcode = val(monthdays(m - 1))
    
    ' Multiplying by 1.25 takes care of leap years,
    ' but not completely. Jan and Feb of a leap year
    ' will end up a day extra.
    ' The 'mod 7' gives us our day.
    DOW = ((Int(y * 1.25) + mcode + d) Mod 7)
    
    ' This takes care of leap year Jan and Feb days.
    If y Mod 4 = 0 And m < 3 Then DOW = (DOW + 6) Mod 7
    
End Function

Private Sub ACControlCode_ShuttingDown()
  SystemShuttingDown = True
End Sub

Private Sub ACControlCode_ProgramStart()
  Delay = 0
  FR.InToOut = 0
  FR.OutToIn = 0
  FC.InToOut = 0
  FC.OutToIn = 0
  DR.LevelGaugeFlushCount = 0
  HD.LevelGaugeFlushCount = 0
  
'Look in database for Recipe Code
' All this DB stuff is super cautious so as to be hopefully bomb-proof
  On Error Resume Next
  'Dim con As ADODB.Connection:   Set con = New ADODB.Connection
  'With con
  '  .ConnectionString = "BatchDyeing": .Open
  '  Dim rs As ADODB.Recordset
  '  Set rs = .Execute("SELECT StandardTime FROM Dyelots WHERE State=2")
  '  If Not rs.EOF Then
  '    Dim val As Variant: val = rs(0).Value
  '    If Not IsNull(val) Then StandardTime = val  ' got it
  '  End If
  '  Set rs = Nothing
  '  con.Close: Set con = Nothing  ' leave no litter
  'End With
  'On Error GoTo 0
  
  ' Use Late Binding
  Dim con As Object: con = CreateObject("ADODB.Connection")
  If Not (con Is Nothing) Then
    With con
      .ConnectionString = "BatchDyeing": .Open
      Dim rs As Object: rs = CreateObject("ADODB.Recordset")
      If Not (rs Is Nothing) Then
        Set rs = .Execute("SELECT StandardTime FROM Dyelots WHERE State=2")
        If Not rs.EOF Then
          Dim val As Variant: val = rs(0).Value
          If Not IsNull(val) Then StandardTime = val  ' got it
        End If
        Set rs = Nothing
      End If
      con.Close: Set con = Nothing  ' leave no litter
    End With
    On Error GoTo 0
  End If
  
   
  StandardTime = StandardTime * 86400
  StartTime = Now
  DrugroomPreview.LoadRecipeSteps        'Load the Recipe associated with this job
  DrugroomPreview.WaitingAtTransferTimer.Start
  DrugroomPreview.WaitingAtTransferTimer.Pause
  KP.KP1.WaitReadyTimer.Start
  KP.KP1.WaitReadyTimer.Pause

End Sub

Private Sub ACControlCode_ProgramStop()
  Alarms_LidNotLocked = False
End Sub

Private Sub Class_Initialize()
  Set Parent = New ACCommandSequencer
  
'Setup modbus comms
  ConnectModbus1

'Setup modbus comms
  ConnectModbus2
  
End Sub

Private Sub ConnectModbus1()
  Set IO_Modbus1 = New ACSerialDevices.ACModbus
  IO_Modbus1Open = IO_Modbus1.Open("COM3", 38400, "E", 8, 1)
  PLC1WatchdogTimer = PLC1WatchdogTime
End Sub

Private Sub ConnectModbus2()
  Set IO_Modbus2 = New ACSerialDevices.ACModbus
  IO_Modbus2Open = IO_Modbus2.Open("COM4", 38400, "E", 8, 1)
  PLC2WatchdogTimer = PLC2WatchdogTime
End Sub

#End If
End Class
