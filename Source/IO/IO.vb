Public Class IO : Inherits MarshalByRefObject
  Public Plc1 As Ports.Modbus
  Public PLC1Timer As New Timer    'Local PLC
  Public PLC1Fault As Boolean
  Public PLC1FaultTimer As New Timer

  Public Plc2 As Ports.Modbus      'Kitchen
  Public PLC2Timer As New Timer
  Public PLC2Fault As Boolean
  Public PLC2FaultTimer As New Timer

  Public IOScanDone As Boolean
  Public WatchDogTimeout As Short

  Public TrioUps As PhoenixTrio


#Region " DIGITAL INPUTS "

  ' CARD #1 
  <IO(IOType.Dinp, 1), Description("Acknowledge Pushbutton"), TranslateDescription("es", "")> Public AdvancePb As Boolean
  <IO(IOType.Dinp, 2), Description("Halt Pushbutton"), TranslateDescription("es", "")> Public HaltPb As Boolean
  <IO(IOType.Dinp, 3), Description("24vdc Control Power On"), TranslateDescription("es", "")> Public PowerOn As Boolean
  <IO(IOType.Dinp, 4), Description(" "), TranslateDescription("es", "")> Public Dinp4 As Boolean
  <IO(IOType.Dinp, 5), Description("Emergency stop Pushbutton"), TranslateDescription("es", "")> Public EmergencyStop As Boolean
  <IO(IOType.Dinp, 6), Description("Temperature Interlock"), TranslateDescription("es", "")> Public TempInterlock As Boolean
  <IO(IOType.Dinp, 7), Description("Kier pressure safe interlock"), TranslateDescription("es", "")> Public PressureInterlock As Boolean
  <IO(IOType.Dinp, 8), Description("Contact Temperature Switch"), TranslateDescription("es", "")> Public ContactTemp As Boolean
  <IO(IOType.Dinp, 9), Description(" "), TranslateDescription("es", "")> Public PressBelow1Psi1 As Boolean    ' Not used in vb6 code?
  <IO(IOType.Dinp, 10), Description(" "), TranslateDescription("es", "")> Public PressBelow1Psi2 As Boolean   ' Not used in vb6 code?
  <IO(IOType.Dinp, 11), Description(" "), TranslateDescription("es", "")> Public Dinp11 As Boolean
  <IO(IOType.Dinp, 12), Description("Main Pump Running Feedback Contact"), TranslateDescription("es", "")> Public PumpRunning As Boolean
  <IO(IOType.Dinp, 13), Description("Main Pump Inverter Fault Contact"), TranslateDescription("es", "")> Public PumpAlarm As Boolean
  <IO(IOType.Dinp, 14), Description("Add Pump Running Contact"), TranslateDescription("es", "")> Public AddPumpRunning As Boolean
  <IO(IOType.Dinp, 15), Description("Locking Pin Raised"), TranslateDescription("es", "")> Public LockingPinRaised As Boolean
  <IO(IOType.Dinp, 16), Description("Lid Closed Limit Switches"), TranslateDescription("es", "")> Public KierLidClosed As Boolean

  ' CARD #2
  <IO(IOType.Dinp, 17), Description(" "), TranslateDescription("es", "")> Public OpenLidPb As Boolean
  <IO(IOType.Dinp, 18), Description(" "), TranslateDescription("es", "")> Public CloseLidPb As Boolean
  <IO(IOType.Dinp, 19), Description(" "), TranslateDescription("es", "DDS Active to reserve tank")> Public DyeDispenseReserve As Boolean
  <IO(IOType.Dinp, 20), Description(" "), TranslateDescription("es", "DDS Active to add tank")> Public DyeDispenseAdd As Boolean
  <IO(IOType.Dinp, 21), Description(" "), TranslateDescription("es", "")> Public LidRaisedSw As Boolean
  <IO(IOType.Dinp, 22), Description(" ")> Public Dinp22 As Boolean
  <IO(IOType.Dinp, 23), Description(" ")> Public Dinp23 As Boolean
  <IO(IOType.Dinp, 24), Description(" ")> Public Dinp24 As Boolean
  <IO(IOType.Dinp, 25), Description(" ")> Public Dinp25 As Boolean  ' Dinp25-31 not defined in vb6
  <IO(IOType.Dinp, 26), Description(" ")> Public Dinp26 As Boolean
  <IO(IOType.Dinp, 27), Description(" ")> Public Dinp27 As Boolean
  <IO(IOType.Dinp, 28), Description(" ")> Public Dinp28 As Boolean
  <IO(IOType.Dinp, 29), Description(" ")> Public Dinp29 As Boolean
  <IO(IOType.Dinp, 30), Description(" ")> Public Dinp30 As Boolean
  <IO(IOType.Dinp, 31), Description(" ")> Public Dinp31 As Boolean
  <IO(IOType.Dinp, 32), Description(" ")> Public Dinp32 As Boolean

  ' DRUGROOM INPUTS 1-20 >> vb6 Dinp 65-73
  ' C0
  <IO(IOType.Dinp, 33), Description("X0")> Public Tank1MixerRequestPb As Boolean
  <IO(IOType.Dinp, 34), Description("X1"), TranslateDescription("es", "")> Public Tank1FillColdPb As Boolean
  <IO(IOType.Dinp, 35), Description("X2"), TranslateDescription("es", "")> Public Tank1FillHotPb As Boolean
  '<IO(IOType.Dinp, 36), Description(" "), TranslateDescription("es", "")> Public Tank1MixerRunning As Boolean
  <IO(IOType.Dinp, 36), Description("X3"), TranslateDescription("es", "")> Public Tank1HeatPb As Boolean
  ' C1
  <IO(IOType.Dinp, 37), Description("X4"), TranslateDescription("es", "")> Public Tank1ReadyPb As Boolean
  <IO(IOType.Dinp, 38), Description("X5")> Public Tank1ManualSw As Boolean
  <IO(IOType.Dinp, 39), Description("X6")> Public Dinp37 As Boolean
  <IO(IOType.Dinp, 40), Description("X7")> Public Dinp38 As Boolean
  ' C2
  <IO(IOType.Dinp, 41), Description("X10")> Public Dinp39 As Boolean  ' Maybe this is CityWaterSw?
  <IO(IOType.Dinp, 42), Description("X11")> Public CityWaterSw As Boolean
  <IO(IOType.Dinp, 43), Description("X12")> Public KitchenX12 As Boolean
  <IO(IOType.Dinp, 44), Description("X13")> Public KitchenX13 As Boolean
  ' C3
  <IO(IOType.Dinp, 45), Description("X14")> Public K2FillHotSw As Boolean
  <IO(IOType.Dinp, 46), Description("X15")> Public KitchenX15 As Boolean
  <IO(IOType.Dinp, 47), Description("X16")> Public KitchenX16 As Boolean
  <IO(IOType.Dinp, 48), Description("X17")> Public KitchenX17 As Boolean
  ' C4
  <IO(IOType.Dinp, 49), Description("X20")> Public KitchenX20 As Boolean
  <IO(IOType.Dinp, 50), Description("X21")> Public KitchenX21 As Boolean
  <IO(IOType.Dinp, 51), Description("X22")> Public KitchenX22 As Boolean
  <IO(IOType.Dinp, 52), Description("X23")> Public KitchenX23 As Boolean


  <IO(IOType.Dinp, 61), Description(" ")> Public UpsBatteryMode As Boolean
  <IO(IOType.Dinp, 62), Description(" ")> Public UpsBatteryOk As Boolean

#If 0 Then
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
  
#End If


#End Region

#Region " DIGITAL OUTPUTS "

  ' CARD #1
  <IO(IOType.Dout, 1, Override.Allow), Description("Stacklight Siren"), TranslateDescription("es", "Stacklight sirena")> Public Siren As Boolean
  <IO(IOType.Dout, 2, Override.Allow), Description("Red Stacklight Lamp"), TranslateDescription("es", "Luz rojo Stacklight")> Public LampAlarm As Boolean
  <IO(IOType.Dout, 3, Override.Allow), Description("Yellow Stacklight Lamp"), TranslateDescription("es", "Luz amarilla Stacklight")> Public LampSignal As Boolean
  <IO(IOType.Dout, 4, Override.Allow), Description("Blue Stacklight Lamp"), TranslateDescription("es", "Luz azul Stacklight")> Public LampDelay As Boolean
  <IO(IOType.Dout, 5, Override.Allow), Description("Machine Safe Lamp"), TranslateDescription("es", "")> Public LampMachineSafe As Boolean
  <IO(IOType.Dout, 6, Override.Allow), Description("Heat Exchanger Drain"), TranslateDescription("es", "")> Public HxDrain As Boolean         ' Added 
  <IO(IOType.Dout, 7, Override.Allow), Description("Main pump start"), TranslateDescription("es", "")> Public MainPump As Boolean
  <IO(IOType.Dout, 8, Override.Allow), Description("Main pump reset"), TranslateDescription("es", "")> Public MainPumpReset As Boolean
  <IO(IOType.Dout, 9, Override.Allow), Description("Add pump start"), TranslateDescription("es", "")> Public AddPumpStart As Boolean
  <IO(IOType.Dout, 10, Override.Allow), Description("Hold Down"), TranslateDescription("es", "")> Public HoldDown As Boolean
  <IO(IOType.Dout, 11, Override.Allow), Description("System Block"), TranslateDescription("es", "")> Public SystemBlock As Boolean
  <IO(IOType.Dout, 12, Override.Allow), Description("Level transducer flush"), TranslateDescription("es", "")> Public LevelFlush As Boolean

  ' CARD #2
  <IO(IOType.Dout, 13, Override.Allow), Description("Vent valve - VE2"), TranslateDescription("es", "")> Public Vent As Boolean
  <IO(IOType.Dout, 14, Override.Allow), Description("Air pad valve - VEL"), TranslateDescription("es", "")> Public AirPad As Boolean
  <IO(IOType.Dout, 15, Override.Allow), Description("Steam valve - VH"), TranslateDescription("es", "")> Public HeatSelect As Boolean
  <IO(IOType.Dout, 16, Override.Allow), Description("Cooling valve - VK"), TranslateDescription("es", "")> Public CoolSelect As Boolean
  <IO(IOType.Dout, 17, Override.Allow), Description("Condensate valve - VH1"), TranslateDescription("es", "")> Public Condensate As Boolean
  <IO(IOType.Dout, 18, Override.Allow), Description("Cooling return valve - VK1"), TranslateDescription("es", "")> Public WaterDump As Boolean
  <IO(IOType.Dout, 19, Override.Allow), Description("Machine Fill Valve - VR"), TranslateDescription("es", "")> Public Fill As Boolean
  <IO(IOType.Dout, 20, Override.Allow), Description("Top wash valve - VE"), TranslateDescription("es", "")> Public TopWash As Boolean
  <IO(IOType.Dout, 21, Override.Allow), Description("Drain valve - VA1"), TranslateDescription("es", "")> Public MachineDrain As Boolean
  <IO(IOType.Dout, 22, Override.Allow), Description("Hot Drain valve - VA2"), TranslateDescription("es", "")> Public MachineDrainHot As Boolean
  <IO(IOType.Dout, 23, Override.Allow), Description("Flow Reverse"), TranslateDescription("es", "")> Public FlowReverse As Boolean
  <IO(IOType.Dout, 24, Override.Allow), Description("Lid Locking Band Open"), TranslateDescription("es", "")> Public OpenLockingBand As Boolean

  ' CARD #3
  <IO(IOType.Dout, 25, Override.Allow), Description("Flow Outlet - In to Out"), TranslateDescription("es", "")> Public OutletInToOut As Boolean
  <IO(IOType.Dout, 26, Override.Allow), Description("Flow Outlet - Out to In"), TranslateDescription("es", "")> Public OutletOutToIn As Boolean
  <IO(IOType.Dout, 27, Override.Allow), Description("Flow Inlet - In to Out"), TranslateDescription("es", "")> Public InletInToOut As Boolean
  <IO(IOType.Dout, 28, Override.Allow), Description("Flow Inlet - Out to In"), TranslateDescription("es", "")> Public InletOutToIn As Boolean
  <IO(IOType.Dout, 29, Override.Allow), Description("Lid Locking Band Close"), TranslateDescription("es", "")> Public CloseLockingBand As Boolean
  <IO(IOType.Dout, 30, Override.Allow), Description("Add Fill"), TranslateDescription("es", "")> Public AddFillCold As Boolean
  <IO(IOType.Dout, 31, Override.Allow), Description("Add Fill Hot"), TranslateDescription("es", "")> Public AddFillHot As Boolean
  <IO(IOType.Dout, 32, Override.Allow), Description("Add drain"), TranslateDescription("es", "")> Public AddDrain As Boolean
  <IO(IOType.Dout, 33, Override.Allow), Description("Add transfer"), TranslateDescription("es", "")> Public AddTransfer As Boolean
  <IO(IOType.Dout, 34, Override.Allow), Description("Add mixing"), TranslateDescription("es", "")> Public AddMixing As Boolean
  <IO(IOType.Dout, 35, Override.Allow), Description("Add runback"), TranslateDescription("es", "")> Public AddRunback As Boolean
  <IO(IOType.Dout, 36, Override.Allow), Description("Reserve Fill"), TranslateDescription("es", "")> Public ReserveFillCold As Boolean

  ' CARD #4
  <IO(IOType.Dout, 37, Override.Allow), Description(""), TranslateDescription("es", "")> Public ReserveFillHot As Boolean
  <IO(IOType.Dout, 38, Override.Allow), Description(""), TranslateDescription("es", "")> Public LowerLockingPin As Boolean
  <IO(IOType.Dout, 39, Override.Allow), Description("Reserve drain - VJ"), TranslateDescription("es", "")> Public ReserveDrain As Boolean
  <IO(IOType.Dout, 40, Override.Allow), Description("Reserve transfer - VNH"), TranslateDescription("es", "")> Public ReserveTransfer As Boolean
  <IO(IOType.Dout, 41, Override.Allow), Description("Reserve Runback - VN"), TranslateDescription("es", "")> Public ReserveRunback As Boolean
  <IO(IOType.Dout, 42, Override.Allow), Description("Reserve Heating - VL"), TranslateDescription("es", "")> Public ReserveHeat As Boolean
  <IO(IOType.Dout, 43, Override.Allow), Description("Reserve Mixer - NB"), TranslateDescription("es", "")> Public ReserveMixerStart As Boolean
  <IO(IOType.Dout, 44, Override.Allow), Description("Tank 1 add valve - DV3"), TranslateDescription("es", "")> Public Tank1TransferToDrain As Boolean
  <IO(IOType.Dout, 45, Override.Allow), Description("Tank 1 add valve - DV1"), TranslateDescription("es", "")> Public Tank1TransferToAdd As Boolean
  <IO(IOType.Dout, 46, Override.Allow), Description("Tank 1 add valve - DV2"), TranslateDescription("es", "")> Public Tank1TransferToReserve As Boolean
  <IO(IOType.Dout, 47, Override.Allow), Description(""), TranslateDescription("es", "")> Public OpenLid As Boolean
  <IO(IOType.Dout, 48, Override.Allow), Description(""), TranslateDescription("es", "")> Public CloseLid As Boolean

  '  VB6
  '  Public IO_PlcWatchDog1 As Boolean            'DOUT 57 not really an output
  '  Attribute IO_PlcWatchDog1.VB_VarDescription = "IO=DOUT 57\r\nAllowOverride=True"



  ' DRUGROOM Outputs 1-16
  ' C0
  <IO(IOType.Dout, 51, Override.Allow), Description("Y0 - Tank 1 mixer start"), TranslateDescription("es", "")> Public Tank1Mixer As Boolean
  <IO(IOType.Dout, 52, Override.Allow), Description("Y1 - Tank 1 fill cold valve"), TranslateDescription("es", "")> Public Tank1FillCold As Boolean
  <IO(IOType.Dout, 53, Override.Allow), Description("Y2 - Tank 1 steam valve"), TranslateDescription("es", "")> Public Tank1Steam As Boolean
  <IO(IOType.Dout, 54, Override.Allow), Description("Y3 - Tank 1 prepare lamp"), TranslateDescription("es", "")> Public Tank1PrepareLamp As Boolean
  ' C1
  <IO(IOType.Dout, 55, Override.Allow), Description("Y4 - Tank 1 drugroom transfer valve"), TranslateDescription("es", "")> Public Tank1Transfer As Boolean
  <IO(IOType.Dout, 56, Override.Allow), Description("Y5 - Tank 1 fill hot valve"), TranslateDescription("es", "")> Public Tank1FillHot As Boolean
  <IO(IOType.Dout, 57, Override.Allow), Description("Y6"), TranslateDescription("es", "")> Public CityWaterFill As Boolean
  <IO(IOType.Dout, 58, Override.Allow), Description("Y7"), TranslateDescription("es", "")> Public DispenseStateLamp As Boolean
  ' C2 - 24vdc EStop
  <IO(IOType.Dout, 59, Override.Allow), Description("Y10")> Public KitchenY10 As Boolean
  <IO(IOType.Dout, 60, Override.Allow), Description("Y11")> Public KitchenY11 As Boolean
  <IO(IOType.Dout, 61, Override.Allow), Description("Y12")> Public KitchenY12 As Boolean
  <IO(IOType.Dout, 62, Override.Allow), Description("Y13")> Public KitchenY13 As Boolean
  ' C4 - 24vdc EStop
  <IO(IOType.Dout, 63, Override.Allow), Description("Y14")> Public KitchenY14 As Boolean
  <IO(IOType.Dout, 64, Override.Allow), Description("Y15")> Public KitchenY15 As Boolean
  <IO(IOType.Dout, 65, Override.Allow), Description("Y16")> Public KitchenY16 As Boolean
  <IO(IOType.Dout, 66, Override.Allow), Description("Y17")> Public KitchenY17 As Boolean





  '<IO(IOType.Dout, 54, Override.Allow), Description(""), TranslateDescription("es", "")> Public Dout54 As Boolean
  '<IO(IOType.Dout, 55, Override.Allow), Description(""), TranslateDescription("es", "")> Public Dout55 As Boolean
  '<IO(IOType.Dout, 57, Override.Allow), Description(""), TranslateDescription("es", "")> Public Dout57 As Boolean
  '<IO(IOType.Dout, 58, Override.Allow), Description(""), TranslateDescription("es", "")> Public Dout58 As Boolean
  '<IO(IOType.Dout, 59, Override.Allow), Description(""), TranslateDescription("es", "")> Public Dout59 As Boolean
  '<IO(IOType.Dout, 60, Override.Allow), Description(""), TranslateDescription("es", "")> Public Dout60 As Boolean
  '<IO(IOType.Dout, 61, Override.Allow), Description(""), TranslateDescription("es", "")> Public Dout61 As Boolean
  '<IO(IOType.Dout, 62, Override.Allow), Description(""), TranslateDescription("es", "")> Public Dout62 As Boolean
  '<IO(IOType.Dout, 63, Override.Allow), Description(""), TranslateDescription("es", "")> Public Dout63 As Boolean
  '<IO(IOType.Dout, 64, Override.Allow), Description(""), TranslateDescription("es", "")> Public Dout64 As Boolean
  '<IO(IOType.Dout, 65, Override.Allow), Description(""), TranslateDescription("es", "")> Public Dout65 As Boolean
  '<IO(IOType.Dout, 66, Override.Allow), Description(""), TranslateDescription("es", "")> Public Dout66 As Boolean


#End Region

#Region " TEMPERATURE INPUTS"
  ' NOTE:  20220408 Add, Blend, And Tank 1 RTD removed for Tina14 as they were not installed
  <IO(IOType.Temp, 1), Description("Vessel Temp (RTD)"), TranslateDescription("es", "")> Public VesselTemp As Short
  <IO(IOType.Temp, 2), Description(""), TranslateDescription("es", "")> Public ReserveTemp As Short
  <IO(IOType.Temp, 3), Description(""), TranslateDescription("es", "")> Public FillTemp As Short
  <IO(IOType.Temp, 4), Description(""), TranslateDescription("es", "")> Public CondensateTemp As Short

  <IO(IOType.Temp, 5), Description("Kitchen Temp (RTD)"), TranslateDescription("es", "")> Public Tank1Temp As Short

  Public Temp1Raw, Temp1Smooth, Temp1Fahrenheit, Temp1Offset As Short
  Public Temp2Raw, Temp2Smooth, Temp2Fahrenheit As Short
  Public Temp3Raw, Temp3Smooth, Temp3Fahrenheit As Short
  Public Temp4Raw, Temp4Smooth, Temp4Fahrenheit As Short
  Public Temp5Raw, Temp5Smooth, Temp5Fahrenheit As Short

  Public TempRaw(4) As Short
  Public TempSmooth(4) As Short
  Public TempSmoothing(4) As Smoothing
  Friend TempFarenheit(4) As Short
  Friend TempCelcius(4) As Short

#End Region

#Region " ANALOG INPUTS "
  <IO(IOType.Aninp, 1), Description("Vessel Level Input"), TranslateDescription("es", "")> Public VesselLevelInput As Short
  <IO(IOType.Aninp, 2), Description("Reserve Level Input"), TranslateDescription("es", "")> Public ReserveLevelInput As Short
  <IO(IOType.Aninp, 3), Description(""), TranslateDescription("es", "")> Public AddLevelInput As Short
  <IO(IOType.Aninp, 4), Description("Differential Pressure Input"), TranslateDescription("es", "")> Public DiffPressInput As Short
  <IO(IOType.Aninp, 5), Description("Vessel Pressure Input"), TranslateDescription("es", "")> Public VesselPressInput As Short
  <IO(IOType.Aninp, 6), Description("Circulation Flow Rate Input"), TranslateDescription("es", "")> Public FlowRateInput As Short
  <IO(IOType.Aninp, 7), Description(""), TranslateDescription("es", "")> Public VesTempInput As Short
  <IO(IOType.Aninp, 8), Description(""), TranslateDescription("es", "")> Public TestInput As Short

  'Raw analog inputs (before scaling or smoothing) - for calibration work
  Public AninpRaw(8) As Short
  Friend AninpScaled(8) As Short
  Public AninpSmooth(8) As Smoothing


  ' KITCHEN PLC ANALOG INPUTS
  ' 4-CHANNEL INPUT CARD 
  <IO(IOType.Aninp, 9), Description("Tank 1 Level Input"), TranslateDescription("es", "")> Public Tank1LevelInput As Short
  <IO(IOType.Aninp, 10), Description(""), TranslateDescription("es", "")> Public TimePotInput As Short
  <IO(IOType.Aninp, 11), Description(""), TranslateDescription("es", "")> Public TempPotInput As Short
  <IO(IOType.Aninp, 12), Description(""), TranslateDescription("es", "")> Public AnalogInput12 As Short

  Public Aninp9Raw, Aninp9Scaled, Aninp9Smooth As Short
  Public Aninp10Raw, Aninp10Scaled, Aninp10Smooth As Short
  Public Aninp11Raw, Aninp11Scaled, Aninp11Smooth As Short
  Public Aninp12Raw, Aninp12Scaled, Aninp12Smooth As Short


#End Region

#Region " ANALOG OUTPUTS "

  <IO(IOType.Anout, 1, Override.Allow), Description("Heat/Cool output(4-20mA)"), TranslateDescription("es", "")> Public HeatCoolOutput As Short
  <IO(IOType.Anout, 2, Override.Allow), Description(""), TranslateDescription("es", "")> Public BlendFillOutput As Short
  <IO(IOType.Anout, 3, Override.Allow), Description(""), TranslateDescription("es", "")> Public AnalogOutput3 As Short
  <IO(IOType.Anout, 4, Override.Allow), Description(""), TranslateDescription("es", "")> Public PumpSpeedOutput As Short

#End Region

  Public Sub New(ByVal controlCode As ControlCode)
    'Setup Communication Ports 
    '>> Use this method so that if the xml file doesn't declare the port, no "is nothing" exceptions will be thrown <<

    Dim plc1Address As String = controlCode.Parent.Setting("Plc1")
    If Not String.IsNullOrEmpty(plc1Address) Then
      Try : Plc1 = New Ports.Modbus(New Ports.ModbusTcp(plc1Address, 502)) : Catch : End Try
    End If

    Dim plc2Address As String = controlCode.Parent.Setting("Plc2")
    If Not String.IsNullOrEmpty(plc2Address) Then
      Try : Plc2 = New Ports.Modbus(New Ports.ModbusTcp(plc2Address, 502)) : Catch : End Try
    End If

    'Reset Communication Port Alarm Timeout
    PLC1Timer.Seconds = MinMax(controlCode.Parameters.PLCComsTime, 10, 100000)
    PLC2Timer.Seconds = MinMax(controlCode.Parameters.PLCComsTime, 10, 100000)


    Dim portTrioUps As String = controlCode.Parent.Setting("UPS")     ' COM1
    If Not String.IsNullOrEmpty(portTrioUps) Then TrioUps = New PhoenixTrio(portTrioUps)


    For i As Integer = 0 To AninpSmooth.GetUpperBound(0)
      AninpSmooth(i) = New Smoothing(0, 2000, 100)
    Next i
    For i As Integer = 0 To TempSmoothing.GetUpperBound(0)
      TempSmoothing(i) = New Smoothing(0, 3000, 50)
    Next i

  End Sub

  Public Function ReadInputs(ByVal parent As ACParent, ByVal dinp() As Boolean, ByVal aninp() As Short, ByVal temp() As Short, ByVal controlCode As ControlCode) As Boolean
    Try
      'Return true if any read succeeds
      Dim returnValue As Boolean = False


      '------------------------------------------------------------------------------------------
      ' PLC1  -  Automation Direct DL205 - H2-EBC100
      '------------------------------------------------------------------------------------------
      If Plc1 IsNot Nothing Then
        ' 32 Digital Inputs Available
        Dim DinpMax As Integer = 32
        Dim DinpMain(DinpMax) As Boolean
        Select Case Plc1.Read(1, 10001, DinpMain)
          Case Ports.Modbus.Result.Fault
            PLC1Fault = True
            If PLC1FaultTimer.Finished Then
              For i As Integer = 1 To DinpMax
                If i <= dinp.GetUpperBound(0) Then dinp(i) = False
              Next i
            End If
          Case Ports.Modbus.Result.OK
            PLC1Fault = False
            ' An array to configure the hardware inputs so that they appear in order in the control code IO Screen
            For i As Integer = 1 To DinpMax
              If i <= dinp.GetUpperBound(0) Then dinp(i) = DinpMain(i)
            Next i

            'Refresh PLC Coms Timer
            PLC1Timer.Seconds = MinMax(controlCode.Parameters.PLCComsTime, 10, 100000)
            IOScanDone = True
            returnValue = True
        End Select

        'Read (8) analog inputs inputs from PLC
        'SLOT6 - F2-08DA04AD-1 - 8-channel analog input 4-20mA <Note that card reads not 0-20ma>  30001-30008
        'SLOT7 - F2-04RTD - 4-channel RTD inputs 300009-3000012
        ' TODO Smooth
        Dim valueSetAnalogInputs(12) As Short
        Select Case Plc1.Read(1, 30001, valueSetAnalogInputs) ' Plc1.Read(1, 30001, valueSetAnalogInputs)
          Case Ports.Modbus.Result.Fault
          Case Ports.Modbus.Result.OK

            'F2-08DA04AD-1: 8-channel 0-20mA = 0-1000 (100.0%)
            For i As Integer = 1 To 8
              AninpRaw(i) = valueSetAnalogInputs(i)
              AninpScaled(i) = ReScale(AninpRaw(i), 819, 4095, 0, 1000)
              '   AninpSmooth(i) = CShort(SmoothAninp(i).Smooth(AninpScaled(i), controlCode.Parameters.SmoothRate))
              aninp(i) = CShort(AninpSmooth(i).Smooth(AninpScaled(i), controlCode.Parameters.SmoothRate))
            Next i

            ' F2-04RTD: 4-channel Temp inputs
            For i As Integer = 9 To 12
              TempRaw(i - 8) = valueSetAnalogInputs(i)
              TempSmooth(i - 8) = CShort(TempSmoothing(i - 8).Smooth(TempRaw(i - 8), controlCode.Parameters.SmoothRate))
              TempFarenheit(i - 8) = TempSmooth(i - 8)
              temp(i - 8) = TempFarenheit(i - 8)
            Next i
            'Temp1Raw = valueSetAnalogInputs(9)
            'Temp2Raw = valueSetAnalogInputs(10)
            'Temp3Raw = valueSetAnalogInputs(11)
            'Temp4Raw = valueSetAnalogInputs(12)


            If controlCode.Parameters.TempTransmitterRange > 0 Then
              temp(1) = ReScale(aninp(1), controlCode.Parameters.TempTransmitterMin, controlCode.Parameters.TempTransmitterMax, 0, controlCode.Parameters.TempTransmitterRange)
            End If

#If 0 Then
            'If Smooth Rate set, then use smoothing value:
            If controlCode.Parameters.SmoothRate > 0 Then
              Dim smoothRate As Integer = controlCode.Parameters.SmoothRate
                       
              ' Temp Input 1
              'Static temp1Smoothing() As Smoothing = {Nothing, New Smoothing}
              'Temp1Smooth = temp1Smoothing(1).Smooth(Temp1Raw, smoothRate)
              'Temp1Fahrenheit = Temp1Smooth

              '' Temp Offset correction
              'If controlCode.Parameters.VesselTempProbeOffset > 100 Then controlCode.Parameters.VesselTempProbeOffset = 100
              'Temp1Offset = CShort(controlCode.Parameters.VesselTempProbeOffset - 50)
              'temp(1) = CShort((((Temp1Fahrenheit - 320) / 9) * 5) + Temp1Offset)

            
              ' Temp Input 2
              Static temp2Smoothing() As SmoothingOld = {Nothing, New SmoothingOld}
              Temp2Smooth = temp2Smoothing(1).Smooth(Temp2Raw, smoothRate)
              Temp2Fahrenheit = Temp2Smooth
              temp(2) = CShort(((Temp2Fahrenheit - 320) / 9) * 5)

              ' Temp Input 3
              Static temp3Smoothing() As SmoothingOld = {Nothing, New SmoothingOld}
              Temp3Smooth = temp3Smoothing(1).Smooth(Temp3Raw, smoothRate)
              Temp3Fahrenheit = Temp3Smooth
              temp(3) = CShort(((Temp3Fahrenheit - 320) / 9) * 5)

              ' Temp Input 4
              Static temp4Smoothing() As SmoothingOld = {Nothing, New SmoothingOld}
              Temp4Smooth = temp4Smoothing(1).Smooth(Temp4Raw, smoothRate)
              Temp4Fahrenheit = Temp4Smooth
              temp(4) = CShort(((Temp4Fahrenheit - 320) / 9) * 5)

            Else
              'Use Raw Values
              aninp(1) = Aninp1Scaled
              aninp(2) = Aninp2Scaled
              aninp(3) = Aninp3Scaled
              aninp(4) = Aninp4Scaled
              aninp(5) = Aninp5Scaled
              aninp(6) = Aninp6Scaled
              aninp(7) = Aninp7Scaled
              aninp(8) = Aninp8Scaled

              Temp1Fahrenheit = Temp1Raw
              Temp2Fahrenheit = Temp2Raw
              Temp3Fahrenheit = Temp3Raw
              Temp4Fahrenheit = Temp4Raw
              temp(1) = CShort(((Temp1Fahrenheit - 320) / 9) * 5)
              temp(2) = CShort(((Temp2Fahrenheit - 320) / 9) * 5)
              temp(3) = CShort(((Temp3Fahrenheit - 320) / 9) * 5)
              temp(4) = CShort(((Temp4Fahrenheit - 320) / 9) * 5)
            End If

#End If

            If controlCode.Parameters.TempTransmitterRange > 0 Then
              temp(1) = ReScale(aninp(7), controlCode.Parameters.TempTransmitterMin, controlCode.Parameters.TempTransmitterMax, 0, controlCode.Parameters.TempTransmitterRange)
            End If


            'Refresh PLC Coms Timer
            PLC1Timer.Seconds = MinMax(controlCode.Parameters.PLCComsTime, 10, 100000)
            returnValue = True
        End Select



      End If


      '------------------------------------------------------------------------------------------
      ' PLC2 - KITCHEN CONTROL PANEL INPUTS
      '------------------------------------------------------------------------------------------
      If Plc2 IsNot Nothing Then

        ' Automation Direct DL06 Input Points: X0-X777 = V40400-V40437 (40400 octal = 16640 decimal) (d06uservol1.pdf page 3-32)
        Dim DinpPlc2(20) As Boolean
        Select Case Plc2.Read(1, 40001 + 16640, DinpPlc2)
          Case Ports.Modbus.Result.Fault
            PLC2Fault = True
          Case Ports.Modbus.Result.OK
            For i As Integer = 1 To 20
              If i <= dinp.GetUpperBound(0) Then dinp(i + 32) = DinpPlc2(i)
            Next i

            'Refresh PLC Coms Timer
            PLC2Timer.Seconds = MinMax(controlCode.Parameters.PLCComsTime, 10, 100000)
            PLC2Fault = False
            returnValue = True
        End Select


        ' Automation Direct DL06 Analog Input Points:  V2000 = (1024 + 40001) = 41025
        Dim valueSetAnalogInputs(4) As Short
        Select Case Plc2.Read(1, 40001 + 1024, valueSetAnalogInputs)
          Case Ports.Modbus.Result.Fault
            ' Clear Inputs? TODO
          Case Ports.Modbus.Result.OK

            ' F0-4AD-1 4-channel 0-10vdc analog
            Aninp9Raw = valueSetAnalogInputs(1)
            Aninp10Raw = valueSetAnalogInputs(2)
            Aninp11Raw = valueSetAnalogInputs(3)
            Aninp12Raw = valueSetAnalogInputs(4)

            'Scaled for 0-10vdc
            Aninp9Scaled = ReScale(Aninp9Raw, 0, 4095, 0, 1000)
            aninp(9) = Aninp9Scaled
            aninp(10) = Aninp10Raw
            aninp(11) = Aninp11Raw
            aninp(12) = Aninp12Raw

            'Refresh PLC Coms Timer
            PLC2Timer.Seconds = MinMax(controlCode.Parameters.PLCComsTime, 10, 100000)
            returnValue = True
        End Select


        Dim valueSetPlc2RtdInputs(2) As Short
        Select Case Plc2.Read(1, 40001 + 1032, valueSetPlc2RtdInputs)
          Case Ports.Modbus.Result.Fault
            ' Clear Inputs? TODO
            ' Temp(5) = 0
          Case Ports.Modbus.Result.OK
            Temp5Raw = valueSetPlc2RtdInputs(1)
            ' Temp6Raw = valueSetPlc2RtdInputs(2)
            ' Temp7Raw = valueSetPlc2RtdInputs(3)
            ' Temp8Raw = valueSetPlc2RtdInputs(4)
            temp(5) = Temp5Raw

            'Refresh PLC Coms Timer
            PLC2Timer.Seconds = MinMax(controlCode.Parameters.PLCComsTime, 10, 100000)
        End Select

        ''Clear analog inputs here as a result of communication timeout.
        '  If Alarms_DrugroomPLCNotResponding Then
        '    Aninp(9) = 0
        '    Aninp(10) = 0
        '    Aninp(11) = 0
        '    Temp(5) = 0
        '  End If

      End If

      ' Place USB Communication feedback for USB Battery Mode
      If TrioUps IsNot Nothing Then
        If TrioUps.ReadTrio Then
          dinp(61) = TrioUps.StatusBatMode

          returnValue = True
        End If

      End If





      ' True if any read succeeds
      Return returnValue

    Catch ex As Exception
      controlCode.Parent.LogException(ex)
    End Try
  End Function

#If 0 Then  ' TODO - VB6 Check and remove
  



'Read analog inputs in main
  If Not IO_Modbus1 Is Nothing Then
    Static Inputs(1 To 16) As Integer
   '8 inputs two words per input use only the first word
   '8 inputs =  V2040 = (1056 + 40001) = 41056
   
    If IO_Modbus1.ReadWords(40001 + 1056, Inputs, NoResponse) Then
      'Read, smooth and scale analog inputs


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

 
End Function

#End If

  Public Sub WriteOutputs(ByVal dout() As Boolean, ByVal anout() As Short, ByVal controlCode As ControlCode)
    Try

      'If the system is stopping or watchdog has timed out then turn all outputs off.
      If controlCode.SystemShuttingDown Then
        For j As Integer = 1 To dout.GetUpperBound(0)
          dout(j) = False
        Next
        For j As Integer = 1 To anout.GetUpperBound(0)
          anout(j) = 0
        Next
      End If

      'Write to the PLC
      Dim i As Integer
      If Plc1 IsNot Nothing Then
        Dim DoutMax As Integer = 62
        Dim DoutValueSet(DoutMax) As Boolean
        ' Outputs are grouped by 12 but first 6 outputs are y0 -> y5 in plc
        ' second set of 6 are Y10-->y15 octal based numbers
        For i = 1 To 6 : DoutValueSet(i) = dout(i) : Next i                     'Card 1: group A (1-6)
        For i = 9 To 14 : DoutValueSet(i) = dout(i - 2) : Next i                'Card 1: group B (7-12)
        For i = 17 To 22 : DoutValueSet(i) = dout(i - 4) : Next i               'Card 2: group A (13-18)
        For i = 25 To 30 : DoutValueSet(i) = dout(i - 6) : Next i               'Card 2: group B (19-24)
        For i = 33 To 38 : DoutValueSet(i) = dout(i - 8) : Next i               'Card 3: group A (25-30)
        For i = 41 To 46 : DoutValueSet(i) = dout(i - 10) : Next i              'Card 3: group B (31-36)
        For i = 49 To 54 : DoutValueSet(i) = dout(i - 12) : Next i              'Card 4: group A (37-42)
        For i = 57 To 62 : DoutValueSet(i) = dout(i - 14) : Next i              'Card 4: group B (43-48)

        'Write the Digital Outputs
        Select Case Plc1.Write(1, 42001, DoutValueSet, Ports.WriteMode.Optimised)
          Case Ports.Modbus.Result.Fault
          Case Ports.Modbus.Result.OK
            PLC1Timer.Seconds = MinMax(controlCode.Parameters.PLCComsTime, 10, 100000)
        End Select

        'Write the analog output
        ' F2-08AD04DA-1 
        ' 4 Out 4-20mA
        ' Output resolution can be: 12-bit (0-4095), 14-bit (0-16383), 16-bit (0-65535)
        ' Use NetEdit - Show Base Contents and found that 4-channels use 8-words 
        Dim TempAnout(4) As Short
        For i = 1 To TempAnout.GetUpperBound(0)
          If i <= anout.GetUpperBound(0) Then TempAnout(i) = CType(Math.Min(MulDiv(anout(i), 65535, 1000), Short.MaxValue), Short)
        Next i
        Select Case Plc1.Write(1, 40001, TempAnout, Ports.WriteMode.Optimised)
          Case Ports.Modbus.Result.Fault
          Case Ports.Modbus.Result.OK
            PLC1Timer.Seconds = MinMax(controlCode.Parameters.PLCComsTime, 10, 100000)
        End Select

        ' Watchdog Timeout Set on PLC - Ethernet Base Controller Manual page 4-8
        ' Method of writing to 410007 causes PLC Faults
        Dim WatchdogTimeout(1) As Short
        WatchdogTimeout(1) = 10000
        Plc1.Write(1, 50007, WatchdogTimeout, Ports.WriteMode.Always)

      End If



      '------------------------------------------------------------------------------------------
      ' PLC2 - KITCHEN CONTROL PANEL OUTPUTS
      '------------------------------------------------------------------------------------------
      If Plc2 IsNot Nothing Then

        Dim DoutMax As Integer = 16
        Dim DoutPlcDl06(DoutMax) As Boolean

        'Update PLC Write Array
        For i = 1 To DoutMax : DoutPlcDl06(i) = dout(i + 50) : Next i

        'Write the Digital Outputs
        ' V-memory Octal V40500-V40537 (Octal 40500 = decimal 16704)
        Select Case Plc2.Write(1, 40001 + 16704, DoutPlcDl06, Ports.WriteMode.Optimised)
          Case Ports.Modbus.Result.Fault
          Case Ports.Modbus.Result.OK
            PLC2Timer.Seconds = MinMax(controlCode.Parameters.PLCComsTime, 10, 100000)
        End Select

        ' Keep writing a true to a register in the AD plc - this acts as a coms watchdog for that plc
        ' V-Memory is V40501 for discrete bits Y20 through Y37, with alias VY20 (Octal 40501 = decimal 16705)
        Dim watchdogOutput(2) As Boolean
        watchdogOutput(1) = True
        watchdogOutput(2) = True
        Plc2.Write(1, 40001 + 16705, watchdogOutput, Ports.WriteMode.Always)

      End If

    Catch ex As Exception
      controlCode.Parent.LogException(ex)
    End Try
  End Sub

#If 0 Then  ' TODO vb6 check & remove
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


#End If

End Class
