Public Class Parameters : Inherits MarshalByRefObject
  'Save a local reference for convenience
  Private controlCode As ControlCode

  Public Sub New(controlCode As ControlCode)
    Me.controlCode = controlCode
  End Sub



#Region " ADD TANK "
  Private Const section_AddTank As String = "Add Tank"

  <Parameter(0, 10), Category(section_AddTank),
  Description("Number of tank rinses during a add tank transfer."),
  TranslateDescription("es", "")>
  Public AddRinses As Integer

  <Parameter(0, 10), Category(section_AddTank),
  Description("Time, in seconds, to pause add transfer pump if level does not appear to be reducing."),
  TranslateDescription("es", "")>
  Public AddPumpPauseTime As Integer

  <Parameter(0, 60), Category(section_AddTank),
  Description("Time, in seconds, to stop transferring if taking more than this time.  Prevent damage to pump."),
  TranslateDescription("es", "")>
  Public AddTransferTimeOverrun As Integer


  <Parameter(500, 10000), Category(section_AddTank),
  Description("Minimum Time, in milliseconds, to transfer the add tank contents when dosing.  Prevent pulsing."),
  TranslateDescription("es", "")>
  Public AddTransferDoseMinTimeMs As Integer




  <Parameter(0, 600), Category(section_AddTank),
  Description("Time, in seconds, to continue transferring the add tank tank to the machine once emtpy before rinsing."),
  TranslateDescription("es", "")>
  Public AddTransferTimeBeforeRinse As Integer

  <Parameter(0, 60), Category(section_AddTank),
  Description("Time, in seconds, to rinse the add tank tank during a transfer to machine or drain."),
  TranslateDescription("es", "")>
  Public AddTimeToRinseToMachine As Integer

  <Parameter(0, 120), Category(section_AddTank),
  Description("Time, in seconds, to rinse the add tank tank during a transfer to drain."),
  TranslateDescription("es", "")>
  Public AddTimeRinseToDrain As Integer

  <Parameter(0, 300), Category(section_AddTank),
  Description("Time, in seconds, to continue transferring the add tank tank to the machine after rinsing to the machine."),
  TranslateDescription("es", "")>
  Public AddTimeAfterRinse As Integer

  <Parameter(0, 300), Category(section_AddTank),
  Description("Time, in seconds, to continue transferring the add tank tank to the drain after rinsing to the drain."),
  TranslateDescription("es", "")>
  Public AddTimeToDrain As Integer

  <Parameter(0, 1000), Category(section_AddTank),
  Description(" "),
    TranslateDescription("es", "")>
  Public AddMixOffLevel As Integer

  <Parameter(0, 1000), Category(section_AddTank),
  Description(" "),
  TranslateDescription("es", "")>
  Public AddMixOnLevel As Integer

  <Parameter(0, 1000), Category(section_AddTank),
  Description("Add Level deadband, in tenths, to consider desired level reached when filling."),
  TranslateDescription("es", "")>
  Public AdditionFillDeadband As Integer

  <Parameter(0, 1000), Category(section_AddTank),
  Description(" "),
  TranslateDescription("es", "")>
  Public AdditionMaxTransferLevel As Integer

  <Parameter(0, 1000), Category(section_AddTank),
  Description(" "),
  TranslateDescription("es", "")>
  Public AddRinseFillLevel As Integer

  <Parameter(0, 1000), Category(section_AddTank),
  Description(" "),
  TranslateDescription("es", "")>
  Public AddRinseMixTime As Integer

  <Parameter(0, 1000), Category(section_AddTank),
  Description(" "),
  TranslateDescription("es", "")>
  Public AddRinseMixPulseTime As Integer

  <Parameter(0, 1000), Category(section_AddTank),
  Description(" "),
  TranslateDescription("es", "")>
  Public AddTransferSettleTime As Integer

  <Parameter(0, 1000), Category(section_AddTank),
  Description(" "),
  TranslateDescription("es", "")>
  Public AddRunbackPulseTime As Integer

  <Parameter(0, 1000), Category(section_AddTank),
  Description(" "),
  TranslateDescription("es", "")>
  Public AddMixPulseTime As Integer

  '<Parameter(0, 1000), Category(section_AddTank), _
  'Description("Slow Add Transfer Proportional Control position, in tenths."),
  'TranslateDescription("es", "")>
  'Public AddTransferSlowOutput As Integer

  '<Parameter(0, 1000), Category(section_AddTank), _
  'Description("Fast Add Transfer Proportional Control position, in tenths."),
  'TranslateDescription("es", "")>
  'Public AddTransferFastOutput As Integer


#End Region


#Region " ALARMS "


  <Parameter(0, 1), Category("Alarms"),
  Description("Set to '1' to enable siren."),
  TranslateDescription("es", "Establece en '1' para activar la sirena.")>
  Public EnableSiren As Integer

  <Parameter(0, 1), Category("Alarms"),
  Description("Set to '1' to disregard communications timeout from network loss."),
  TranslateDescription("es", "Establece en '1' para ignorar el tiempo de espera de comunicaciones de la pérdida de la red.")>
  Public PLCComsLossDisregard As Integer

  <Parameter(1, 10000), Category("Alarms"),
  Description("Time, in seconds, to wait after communication lost before signalling and turning off all outputs."),
  TranslateDescription("es", "Tiempo, en segundos, esperar después de comunicación perdido antes de la señalización y apagar todas las salidas.")>
  Public PLCComsTime As Integer

  '<Parameter(1000, 10000), Category("Alarm"),
  'Description("Communications timeout, in seconds, to turn off all outputs."),
  'TranslateDescription("es", "Comunicaciones tiempo de espera, en segundos, para apagar todas las salidas.")>
  'Public Parameters_WatchdogTimeout As Integer


  <Parameter(0, 60), Category("Alarms"), Description("Alarm if load command runs longer than this, minutes.")>
  Public LoadAlarmTime As Integer

  <Parameter(0, 60), Category("Alarms"), Description("Alarm if unload command runs longer than this, minutes.")>
  Public UnloadAlarmTime As Integer

  <Parameter(0, 60), Category("Alarms"), Description("Alarm if a sample takes longer than this, minutes.")>
  Public SampleAlarmTime As Integer

  <Parameter(0, 60), Category("Alarms"), Description("Alarm if a sample approval takes longer than this, minutes.")>
  Public ApprovalAlarmTime As Integer

  <Parameter(0, 60), Category("Alarms"), Description("Alarm if kitchen prepare takes longer than this, minutes.")>
  Public KitchenPrepareAlarmTime As Integer

  <Parameter(0, 60), Category("Alarms"), Description("Alarm if fill command runs longer than this, minutes.")>
  Public FillAlarmTime As Integer

  <Parameter(0, 60), Category("Alarms"), Description("Alarm if drain command runs longer than this, minutes.")>
  Public DrainAlarmTime As Integer

  <Parameter(0, 60), Category("Alarms"), Description("Alarm if rinse command runs over the programmed time by this amount, minutes.")>
  Public RinseAlarmTime As Integer

  <Parameter(0, 60), Category("Alarms"), Description("Alarm if pH check command runs longer than this, minutes.")>
  Public PHAlarmTime As Integer

  <Parameter(0, 60), Category("Alarms"), Description("Alarm if add prepare takes longer than this, minutes.")>
  Public AddPrepareAlarmTime As Integer

  <Parameter(0, 60), Category("Alarms"), Description("Alarm if add transfer command runs longer than this, minutes.")>
  Public AddTransferAlarmTime As Integer

  <Parameter(0, 60), Category("Alarms"), Description("Alarm if kitchen transfer command runs longer than this, minutes.")>
  Public KitchenTransferAlarmTime As Integer
  

  <Parameter(0, 60), Category("Alarms"), Description("")>
  Public TempCondensateLimit As Integer

  <Parameter(0, 300), Category("Alarms"), Description("")>
  Public TempCondensateLimitTime As Integer


#End Region



#Region " Blend Control "

  <Parameter(0, 100), Category("Blend Control"), Description("If the blend fill temperature error from the setpoint is less than this, F, the valve doesn't adjust.")>
  Public BlendDeadBand As Integer

  <Parameter(0, 100), Category("Blend Control"), Description("Determines blend adjustment amount during a fill or rinse.")>
  Public BlendFactor As Integer

  <Parameter(0, 100), Category("Blend Control"), Description("Determines how quickly the blend valve adjusts during a fill or rinse, seconds.")>
  Public BlendSettleTime As Integer

  <Parameter(0, 180), Category("Blend Control"), Description("Set to the temperature of the cold water supplied to the machine, in F.")>
  Public ColdWaterTemperature As Integer

  <Parameter(0, 180), Category("Blend Control"), Description("Set to the temperature of the hot water supplied to the machine, in F.")>
  Public HotWaterTemperature As Integer

#End Region




#Region " CALIBRATION - FLOWRATE"
  Private Const section_CalFlowRate As String = "Calibration - Flowrate Transmitters"

  <Parameter(0, 1000), Category(section_CalFlowRate),
  Description("The equivalent flow value, in tenths of a gallons/min, for 100% system flow input.")>
  Public FlowRateMachineRange As Integer

  <Parameter(0, 100), Category(section_CalFlowRate),
  Description("The value that the controller reads from the system flow transmitter when the pump is stopped and there is not flow, in tenths of a percent.")>
  Public FlowRateMachineMin As Integer

  <Parameter(0, 100), Category(section_CalFlowRate),
  Description("The value that the controller reads from the system flow transmitter when the pump is running at 100%, in tenths of a percent.")>
  Public FlowRateMachineMax As Integer

#End Region

#Region " CALIBRATION - LEVEL TRANSMITTERS "

  <Parameter(0, 1000), Category("Calibration - Level Transmitters"),
  Description("Maximum Raw Analog Value for the add level input when full,in tenths. Used to calibrate Add Level Input."),
  TranslateDescription("es", "")>
  Public AddLevelTransmitterMax As Integer

  <Parameter(0, 1000), Category("Calibration - Level Transmitters"),
  Description("Minimum Raw Analog Value for the add level input when full,in tenths. Used to calibrate Add Level Input."),
  TranslateDescription("es", "")>
  Public AddLevelTransmitterMin As Integer

  <Parameter(0, 1000), Category("Calibration - Level Transmitters"),
  Description("Maximum Raw Analog Value for the reserve level input when full,in tenths. Used to calibrate reserve Level Input."),
  TranslateDescription("es", "")>
  Public ReserveLevelTransmitterMax As Integer

  <Parameter(0, 1000), Category("Calibration - Level Transmitters"),
  Description("Minimum Raw Analog Value for the reserve level input when full,in tenths. Used to calibrate reserve Level Input."),
  TranslateDescription("es", "")>
  Public ReserveLevelTransmitterMin As Integer

  <Parameter(0, 1000), Category("Calibration - Level Transmitters"),
  Description("Maximum Raw Analog Value for the vessel level input when full,in tenths. Used to calibrate Vessel Level Input."),
  TranslateDescription("es", "")>
  Public VesselLevelTransmitterMax As Integer

  <Parameter(0, 1000), Category("Calibration - Level Transmitters"),
  Description("Minimum Raw Analog Value for the vessel level input when full,in tenths. Used to calibrate Vessel Level Input."),
  TranslateDescription("es", "")>
  Public VesselLevelTransmitterMin As Integer

  <Parameter(0, 1000), Category("Calibration - Level Transmitters"),
  Description("Level signal bounce to disregard when calibrating level signal, in tenths.  Used with Tank1Mixer running to disregard some input changes.  Set to '0' to disable. Used to calibrate Tank1 Level Input."),
  TranslateDescription("es", "Rebote de la señal de nivel hacer caso omiso cuando calibrar señal de nivel, en décimas.  Utilizado con el funcionamiento de Tank1Mixer hacer caso omiso de algunos cambios de entrada.  Establece en '0' para desactivar. Utilizado para calibrar Tank1 nivel de entrada.")>
  Public Tank1LevelDisregardDelta As Integer

  <Parameter(0, 99), Category("Calibration - Level Transmitters"),
  Description("Time, in seconds, to advance during states related to tank 1 transferring.  Bouncing signal onsite requires frequent interaction to manually advance so use this instead."),
  TranslateDescription("es", "Tiempo, en segundos, para avanzar durante los Estados relacionados con la transferencia de depósito 1.  Que despide señal in situ requiere la interacción frecuente para avanzar manualmente así que uso esta en su lugar.")>
  Public Tank1LevelDisregardTimeTransfer As Integer

  <Parameter(0, 99), Category("Calibration - Level Transmitters"),
  Description("Time, in seconds, to advance during states related to tank 1 filling.  Bouncing signal onsite requires frequent interaction to manually advance so use this instead."),
  TranslateDescription("es", "Time, in seconds, to advance during states related to tank 1 filling.  Bouncing signal onsite requires frequent interaction to manually advance so use this instead.")>
  Public Tank1LevelDisregardTimeFill As Integer

  <Parameter(0, 1000), Category("Calibration - Level Transmitters"),
  Description("Minimum Raw Analog Value for the Tank1 level input when full,in tenths. Used to calibrate Tank1 Level Input."),
  TranslateDescription("es", "")>
  Public Tank1LevelTransmitterMin As Integer

  <Parameter(0, 1000), Category("Calibration - Level Transmitters"),
  Description("Maximum Raw Analog Value for the Tank1 level input when full,in tenths. Used to calibrate Tank1 Level Input."),
  TranslateDescription("es", "")>
  Public Tank1LevelTransmitterMax As Integer

#End Region

#Region " CALIBRATION - PRESSURE TRANSMITTERS "
  Private Const section_CalPressure As String = "Calibration - Pressure Transmitters"

  <Parameter(0, 1000), Category(section_CalPressure), Description("The minimum raw analog input provided by the vessel pressure transmitter, in tenths of a percent.  Used to calibrate vessel pressure.")>
  Public VesselPressureMin As Integer

  <Parameter(0, 1000), Category(section_CalPressure), Description("The maximum raw analog input provided by the vessel pressure transmitter, in tenths of a percent.  Used to calibrate vessel pressure.")>
  Public VesselPressureMax As Integer

  <Parameter(0, 9999), Category(section_CalPressure), Description("The 100% input pressure of the vessel pressure transmitter, in tenths of psi.")>
  Public VesselPressureRange As Integer

  <Parameter(0, 99), Category(section_CalPressure), Description("Set to '1' to enable differential pressure display.  Set to '0' for faulty DP Transmitter."),
  TranslateDescription("es", "")>
  Public PackageDiffPressEnable As Integer

  <Parameter(0, 1000), Category(section_CalPressure), Description("The minimum input provided by the differential pressure transmitter, in tenths of a percent."),
  TranslateDescription("es", "")>
  Public PackageDiffPressZero As Integer

  <Parameter(0, 1000), Category(section_CalPressure), Description("The maximum input provided by the differential pressure transmitter, in tenths of a percent."),
  TranslateDescription("es", "")>
  Public PackageDiffPress100 As Integer

  <Parameter(0, 100000), Category(section_CalPressure), Description("The 100% input pressure of the differential pressure transmitter, in tenths psi."),
  TranslateDescription("es", "")>
  Public PackageDiffPressRange As Integer

#End Region

#Region " CALIBRATION - TEMPERATURE TRANSMITTERS "
  Private Const section_CalTemperature As String = "Calibration - Temperature Transmitters"

  <Parameter(0, 1), Category(section_CalTemperature),
  Description("Set to '1' to indicate that Reserve tank has RTD installed and working.  Used to enable reserve tank heating functions."),
  TranslateDescription("es", "Establece en '1' para indicar que el tanque RTD instalado y funcionando.  Sirve para habilitar el tanque calefacción funciones.")>
  Public ReserveTempProbe As Integer

  <Parameter(0, 1), Category(section_CalTemperature),
  Description("Set to '1' to indicate that Kitchen tank has RTD installed and working.  Used to enable Kitchen tank heating functions."),
  TranslateDescription("es", "Establece en '1' para indicar que el tanque cocina RTD instalado y funcionando.  Sirve para habilitar el tanque cocina calefacción funciones.")>
  Public Tank1TempProbe As Integer

  <Parameter(0, 100), Category(section_CalTemperature),
  Description("Offset to calibrate the vessel temperature input. To enable negative offset, 50 represents 0 offset.  100 represents 5.0C and 0 represents -5.0C."),
  TranslateDescription("es", "Desplazamiento para calibrar la entrada de temperatura del recipiente. Para habilitar el desplazamiento negativo, 50 representa el desplazamiento 0.  100 representa 5.0C y 0 representa -5.0C.")>
  Public VesselTempProbeOffset As Integer


  <Parameter(0, 100), Category(section_CalTemperature),
  Description(""),
  TranslateDescription("es", ".")>
  Public TempTransmitterRange As Integer

  <Parameter(0, 100), Category(section_CalTemperature),
  Description(""),
  TranslateDescription("es", ".")>
  Public TempTransmitterMax As Integer

  <Parameter(0, 100), Category(section_CalTemperature),
  Description(""),
  TranslateDescription("es", ".")>
  Public TempTransmitterMin As Integer



#End Region

#Region "CALIBRATION - TEST TRANSMITTER"
  Private Const section_CalTestInput As String = "Calibration - Test Transmitter"

  <Parameter(0, 100), Category(section_CalTestInput),
  Description(""),
  TranslateDescription("es", ".")>
  Public TestTransmitterRange As Integer

  <Parameter(0, 100), Category(section_CalTestInput),
  Description(""),
  TranslateDescription("es", ".")>
  Public TestTransmitterMax As Integer

  <Parameter(0, 100), Category(section_CalTestInput),
  Description(""),
  TranslateDescription("es", ".")>
  Public TestTransmitterMin As Integer

#End Region



#Region "CALIBRATION - DRUGROOM CALIBRATION"
  Private Const section_Tank1Pots As String = "Calibration - Drugroom Calibration"

  <Parameter(0, 1000), Category(section_Tank1Pots), Description("The value that the controller reads from the temperature pot at the fully counter clockwise position. in tenths of a percent."),
  TranslateDescription("es", "")>
  Public Tank1TempPotMin As Integer

  <Parameter(0, 1000), Category(section_Tank1Pots), Description("The value that the controller reads from the temperature pot at the fully clockwise position. in tenths of a percent."),
  TranslateDescription("es", "")>
  Public Tank1TempPotMax As Integer

  <Parameter(0, 100000), Category(section_Tank1Pots), Description("The value for the fully clockwise position of the temperature pot.  in tenths of a degree F"),
  TranslateDescription("es", "")>
  Public Tank1TempPotRange As Integer


  <Parameter(0, 1000), Category(section_Tank1Pots), Description("The value that the controller reads from the time pot at the fully counter clockwise position. in tenths of a percent."),
  TranslateDescription("es", "")>
  Public Tank1TimePotMin As Integer

  <Parameter(0, 1000), Category(section_Tank1Pots), Description("The value that the controller reads from the time pot at the fully clockwise position. in tenths of a percent."),
  TranslateDescription("es", "")>
  Public Tank1TimePotMax As Integer

  <Parameter(0, 100000), Category(section_Tank1Pots), Description("The value for the fully clockwise position of the time pot, in minutes."),
  TranslateDescription("es", "")>
  Public Tank1TimePotRange As Integer


#End Region

#Region " DEMO MODE"

  <Parameter(0, 10000), Category("Demo Mode"), Description("Set to '1' to indicate Demo Mode for Batch Control."),
  TranslateDescription("es", "Establezca en '1' para indicar el modo de demostración para el lote Control.")>
  Public Demo As Integer

#End Region


#Region " DISPENSE "

  <Parameter(0, 999), Category("Dispense"), Description(" "),
  TranslateDescription("es", "")>
  Public DDSEnabled As Integer


  <Parameter(0, 999), Category("Dispense"), Description(" "),
  TranslateDescription("es", "")>
  Public DispenseEnabled As Integer

  <Parameter(0, 30), Category("Dispense"), Description(" "),
  TranslateDescription("es", "")>
  Public DispenseReadyDelayTime As Integer


  <Parameter(0, 30), Category("Dispense"), Description(" "),
  TranslateDescription("es", "")>
  Public DispenseResponseDelayTime As Integer



  <Parameter(0, 999), Category("Dispense"), Description(" "),
  TranslateDescription("es", "")>
  Public DispenseTestCalloff As Integer

  <Parameter(0, 999), Category("Dispense"), Description(" "),
  TranslateDescription("es", "")>
  Public DispenseTestState As Integer

#End Region



#Region " FLOW CONTROL "

  <Parameter(0, 100), Category("Flow Control"), Description(" "),
  TranslateDescription("es", "")>
  Public FLAdjustFactor As Integer

  <Parameter(0, 100), Category("Flow Control"), Description(" "),
  TranslateDescription("es", "")>
  Public FLAdjustMaximum As Integer

  <Parameter(0, 100), Category("Flow Control"), Description(" "),
  TranslateDescription("es", "")>
  Public SystemVolumeAt0Level As Integer

  <Parameter(0, 100), Category("Flow Control"), Description(" "),
  TranslateDescription("es", "")>
  Public SystemVolumeAt100Level As Integer

  <Parameter(0, 100), Category("Flow Control"), Description(" "),
  TranslateDescription("es", "")>
  Public LowFlowRatePercent As Integer

  <Parameter(0, 100), Category("Flow Control"), Description(" "),
  TranslateDescription("es", "")>
  Public LowFlowRateTime As Integer


#End Region





#Region " LEVEL CONTROL - DRAINING "

  <Parameter(0, 1000), Category("Level Control - Draining"),
  Description("Minutes to delay draining machine, once idle. Set to '0' to disable."),
  TranslateDescription("es", "Minutos para retrasar una vez drenaje máquina, inactivo. Establece en '0' para desactivar.")>
  Public IdleDrainTime As Integer

  <Parameter(0, 1000), Category("Level Control - Draining"),
  Description("Time, in seconds, to continue draining the machine once empty."),
  TranslateDescription("es", "Tiempo, en segundos, para continuar drenando la máquina una vez vacío.")>
  Public DrainMachineTime As Integer

  <Parameter(0, 1000), Category("Level Control - Draining"),
  Description("Time, in seconds, to preform a level flush while draining if the drain does not complete.  May be issue with faulty level transmitter."),
  TranslateDescription("es", "Tiempo en segundos, para preformas color nivel mientras drenaje si el desagüe no se completa.  Puede ser problema con el transmisor defectuoso del nivel.")>
  Public DrainMachineFlushTime As Integer

  <Parameter(0, 1000), Category("Level Control - Draining"),
  Description("Time, in seconds, to continue high temp draining once level is lost."),
  TranslateDescription("es", "Tiempo, en segundos, para continuar alta temp vaciado una vez perdido el nivel.")>
  Public HDDrainTime As Integer

  <Parameter(0, 1000), Category("Level Control - Draining"),
  Description("Time, in seconds, to continue high temp draining once level is lost."),
  TranslateDescription("es", "Tiempo, en segundos, para continuar alta temp vaciado una vez perdido el nivel.")>
  Public HDFillPosition As Integer

  <Parameter(0, 1000), Category("Level Control - Draining"),
  Description("Blend Fill position to fill during Hot Drain"),
  TranslateDescription("es", "")>
  Public HDVentAlarmTime As Integer

  <Parameter(0, 1000), Category("Level Control - Draining"),
  Description("Level, in tenths, to fill expansion to when filling to cool during hot drain, if greater than 'Fill Level Expansion 1'."),
  TranslateDescription("es", "Nivel, en décimas, para llenar expansión cuando relleno para enfriar durante caliente drenaje, si es mayor que 'Llenado nivel de expansión 1'.")>
  Public HDHiFillLevel As Integer

  <Parameter(0, 1000), Category("Level Control - Draining"),
  Description("Time, in seconds, to delay draining once Hot Drain has filled to cool and once machine temperature is safe.  Delay before atmospheric drain."),
  TranslateDescription("es", "Tiempo, en segundos, a retardo vaciado una vez caliente drenaje ha llenado para enfriar y una vez la temperatura de la máquina es segura.  Retraso antes de la descarga atmosférica.")>
  Public HDFillRunTime As Integer

  <Parameter(0, 1000), Category("Level Control - Draining"),
  Description("Time, in seconds, to perform a level flush while hot draining if the drain does not complete.  May be issue with faulty level transmitter."),
  TranslateDescription("es", "Tiempo en segundos, para preformas color nivel mientras drenaje si el desagüe no se completa.  Puede ser problema con el transmisor defectuoso del nivel.")>
  Public HotDrainMachineFlushTime As Integer

#End Region

#Region " LEVEL CONTROL - FILLING "
  Private Const section_LevelCntrFilling As String = "Level Control - Filling"

  <Parameter(0, 600), Category(section_LevelCntrFilling), Description("Set to '1' to enable filling with both hot and cold water.  Otherwise, fill will only consist of cold or hot, if Fill Hot is enabled."),
  TranslateDescription("es", "Tiempo, en segundos, para retrasar antes de llenar con agua fría.  Establece en 0 para desactivar.")>
  Public FillColdOverrideTime As Integer

  '<Parameter(0, 1), Category(section_LevelCntrFilling), Description("Set to '1' to enable filling with both hot and cold water.  Otherwise, fill will only consist of cold or hot, if Fill Hot is enabled."),
  'TranslateDescription("es", "Establezca en '1' para permitir el llenado con ambos caliente y agua fría.  De lo contrario, llene sólo consisten en frío o caliente, si el relleno caliente está activado.")>
  'Public FillEnableBlend As Integer

  '<Parameter(0, 1), Category(section_LevelCntrFilling), Description("Set to '1' to enable filling with hot water.  Otherwise, fill will only consist of cold."),
  'TranslateDescription("es", "Establezca en '1' para permitir el llenado con agua caliente.  De lo contrario, relleno consistirá sólo en frío.")>
  'Public FillEnableHot As Integer

  <Parameter(0, 1000), Category(section_LevelCntrFilling), Description("Maximum Fill Level, in tenths."),
  TranslateDescription("es", "Máximo nivel lleno, en décimas.")>
  Public FillLevelMaximum As Integer

  <Parameter(0, 1000), Category(section_LevelCntrFilling), Description("Minimum Fill Level, in tenths."),
  TranslateDescription("es", "Mínimo nivel lleno, en décimas.")>
  Public FillLevelMinimum As Integer

  <Parameter(0, 10000), Category(section_LevelCntrFilling), Description("Maximum Fill Volume, in tenths liters."),
  TranslateDescription("es", "")>
  Public FillVolumeMaximum As Integer

  <Parameter(0, 10000), Category(section_LevelCntrFilling), Description("Minimum Fill Volume, in tenths liters."),
  TranslateDescription("es", "")>
  Public FillVolumeMinimum As Integer

  <Parameter(0, 1000), Category(section_LevelCntrFilling), Description("Fill Level 1, in tenths of percent"),
  TranslateDescription("es", "")>
  Public FillLevel1 As Integer

  <Parameter(0, 1000), Category(section_LevelCntrFilling), Description("Fill Level 2, in tenths of percent"),
  TranslateDescription("es", "")>
  Public FillLevel2 As Integer

  <Parameter(5, 60), Category(section_LevelCntrFilling),
  Description("Time, in seconds, to delay opening Fill Valve while Top Wash valve opens.  Prevent excess pressure in kier due to water pressure."),
  TranslateDescription("es", "Tiempo en segundos, para retrasar la valvula de apertura mientras que abre la válvula de lavado superior.  Evitar el exceso de presión en kier debido a la presión del agua.")>
  Public FillOpenDelayTime As Integer

  <Parameter(0, 600), Category(section_LevelCntrFilling),
  Description("The time to turn on the pump and let the level settle before adding more water to the machine during a fill."),
  TranslateDescription("es", "El momento de encender la bomba y dejar que el nivel instalan antes de añadir más agua a la máquina durante un relleno.")>
  Public FillSettleTime As Integer

  <Parameter(0, 600), Category(section_LevelCntrFilling),
  Description("Time, in seconds, to turn the pump on to prime it."),
  TranslateDescription("es", "Tiempo, en segundos, para encender la bomba para cebarla.")>
  Public FillPrimePumpTime As Integer

  <Parameter(0, 600), Category(section_LevelCntrFilling),
  Description("Time, in seconds, to continue filling the machine after the fill level is reached."),
  TranslateDescription("es", "Tiempo en segundos, para seguir llenando la máquina después de alcanza el nivel de llenado.")>
  Public OverFillTime As Integer

  <Parameter(0, 1000), Category(section_LevelCntrFilling),
  Description("Time, in seconds, to flush the analog level transmitter to clean after draining."),
  TranslateDescription("es", "Tiempo, en segundos, para limpiar el transmisor de nivel analógico para limpiar después de drenar.")>
  Public LevelGaugeFlushTime As Integer

  <Parameter(0, 1000), Category(section_LevelCntrFilling),
   Description("Time, in seconds, to allow level transmitter to drain and settle after flushing."),
  TranslateDescription("es", "Tiempo en segundos, para permitir que el transmisor de nivel drenar y colocar después del lavado.")>
  Public LevelGaugeSettleTime As Integer

  <Parameter(0, 1000), Category(section_LevelCntrFilling),
  Description("Time, in seconds, to delay after filling before recording the level of the machine once filled.  Used to indicate level loss alarm."),
  TranslateDescription("es", "Tiempo, en segundos, para retrasar después de llenar antes de grabar el nivel de la máquina una vez.  Utiliza para indicar la alarma de pérdida de nivel.")>
  Public RecordLevelTimer As Integer

  <Parameter(0, 1000), Category(section_LevelCntrFilling),
  Description("Maximum level amount, in tenths, to allow the machine level to drop below the recorded level before indicating a level loss alarm.  Disable alarm with value of '0'."),
  TranslateDescription("es", "Máximo nivel, en décimas, para permitir que el nivel caiga por debajo del nivel registrado antes de indicar una alarma de pérdida de nivel del equipo.  Desactivar la alarma con el valor '0'.")>
  Public MaxLevelLossAmount As Integer

  '  <Parameter(0, 1000), Category(section_LevelCntrFilling), Description(" "),
  '  TranslateDescription("es", "")>
  '  Public RinseTemperatureAlarmBand As Integer

  '  <Parameter(0, 1000), Category(section_LevelCntrFilling), Description(" "),
  '  TranslateDescription("es", "")>
  '  Public RinseTemperatureAlarmDelay As Integer

  <Parameter(5, 60), Category(section_LevelCntrFilling),
  Description("Time, in seconds, to delay opening Fill Valve while Top Wash valve opens.  Prevent excess pressure in kier due to water pressure."),
  TranslateDescription("es", "Tiempo en segundos, para retrasar la valvula de apertura mientras que abre la válvula de lavado superior.  Evitar el exceso de presión en kier debido a la presión del agua.")>
  Public TopWashOpenDelay As Integer

  <Parameter(0, 1000), Category(section_LevelCntrFilling),
  Description("Time, in seconds, to indicate the top wash valve is blocked when rinsing.  Disable with value of '0'."),
  TranslateDescription("es", "Tiempo, en segundos, para indicar que la válvula de lavado superior está bloqueada cuando enjuague.  Desactivar con valor de '0'.")>
  Public TopWashBlockageTimer As Integer

#End Region

#Region " LID CONTROL "
  Private Const section_LidControl As String = "Lid Control"

  <Parameter(0, 600), Category(section_LevelCntrFilling), Description("The level in tenths of a percent that it is ok to open the lid at."),
    TranslateDescription("es", "El nivel en décimas de un porcentaje al que está bien abrir la tapa.")>
  Public LevelOkToOpenLid As Integer

  <Parameter(5, 30), Category(section_LevelCntrFilling), Description("The time in seconds to wait to ensure the lid is raised."),
    TranslateDescription("es", "El tiempo en segundos de espera para asegurarse de que la tapa esté levantada. ")>
  Public LidRaisingTime As Integer

  <Parameter(5, 30), Category(section_LevelCntrFilling), Description("The time in seconds to wait to ensure the locking band is open."),
    TranslateDescription("es", "El tiempo en segundos de espera para asegurarse de que la banda de bloqueo esté abierta.")>
  Public LockingBandOpeningTime As Integer

  <Parameter(5, 30), Category(section_LevelCntrFilling), Description("The time in seconds to wait to ensure the locking band is closed."),
    TranslateDescription("es", "El tiempo en segundos de espera para asegurarse de que la banda de bloqueo esté cerrada.")>
  Public LockingBandClosingTime As Integer

  <Parameter(5, 30), Category(section_LevelCntrFilling), Description("The time in seconds to wait to release hold down before pulling pin."),
    TranslateDescription("es", "El tiempo en segundos para esperar a soltar, mantenga presionado antes de tirar del pasador")>
  Public CarrierHoldDownTime As Integer

#End Region

#Region " LOOK AHEAD "

  <Parameter(0, 1), Category("Look Ahead"),
  Description("Set to '1' to enable the look ahead function."),
  TranslateDescription("es", "Establezca en '1' para activar la función de look ahead.")>
  Public LookAheadEnabled As Integer

  <Parameter(0, 1), Category("Look Ahead"),
  Description("Set to '1' to enable the look ahead function for blocked dyelots."),
  TranslateDescription("es", "Se establece en '1' para activar la función de mirada adelante para dyelots bloqueado.")>
  Public LookAheadIgnoreBlocked As Integer

  <Parameter(0, 1), Category("Look Ahead"),
  Description("Set to '1' to enable alarm indicating a redye issue during look ahead."),
  TranslateDescription("es", "Establece en '1' para activar la alarma que indica un problema reteñido durante mirada adelante.")>
  Public EnableRedyeIssueAlarm As Integer

#End Region

#Region " PRODUCTION REPORTS "

  <Parameter(0, 1000), Category("Production reports"),
  Description("The horse power rating of the main pump used to calculate the amount of power used."),
  TranslateDescription("es", "La potencia de caballos de la bomba principal para calcular la cantidad de energía utilizada.")>
  Public MainPumpMotorHP As Integer

  <Parameter(0, 1000), Category("Production reports"),
Description("The horse power rating of the addition feed pump used to calculate the amount of power used."),
  TranslateDescription("es", "La potencia de caballos de la adición de alimentación bomba usada para calcular la cantidad de energía utilizada.")>
  Public AddFeedPumpMotorHP As Integer

#End Region

#Region " PRESSURE CONTROL "

  <Parameter(0, 4000), Category("Pressure Control"),
  Description("The machine pressure, in tenths, above which to use the bleed valve to reduce."),
  TranslateDescription("es", "La presión de la máquina, en décimas, que utilizar la válvula de purga para reducir.")>
  Public AirpadBleedPressure As Integer

  <Parameter(250, 10000), Category("Pressure Control"),
    Description("Time, in milliseconds, to delay closing the vent valve when bleeding the system due to the pressure rising above the parameter 'Airpad Bleed Pressure.'  Minimum value 250ms."),
  TranslateDescription("es", "Tiempo, en milisegundos, para retrasar el cierre de la válvula de ventilación cuando se purga del sistema debido a la presión sobre el parámetro 'Airpad sangrar presión.'  Valor mínimo 250ms.")>
  Public AirpadBleedTime As Integer

  <Parameter(0, 2000), Category("Pressure Control"),
  Description("The machine pressure, in tenths, above which to turn off the air pressure supply valve."),
  TranslateDescription("es", "La presión de la máquina, en décimas, que apagar la válvula de suministro de presión de aire.")>
  Public AirpadCutoffPressure As Integer

  <Parameter(0, 4000), Category("Pressure Control"),
  TranslateDescription("es", "")>
  Public ATBleedPressure As Integer

#End Region

#Region " RESERVE TANK "
  Private Const section_ReserveTank As String = "Reserve Tank"

  <Parameter(0, 300), Category(section_ReserveTank),
  TranslateDescription("es", "")>
  Public ReserveHeatTime As Integer

  <Parameter(0, 750), Category(section_ReserveTank),
  TranslateDescription("es", "")>
  Public ReserveTankHighTempLimit As Integer

  <Parameter(0, 60), Category(section_ReserveTank),
  TranslateDescription("es", "")>
  Public ReserveDrainTime As Integer

  <Parameter(0, 100), Category(section_ReserveTank),
  TranslateDescription("es", "")>
  Public ReserveHeatDeadband As Long

  <Parameter(400, 750), Category(section_ReserveTank),
  Description("Minimum required reserve level, in tenths, before the Reserve live steam can activate to heat the Reserve Tank.  Minimum Value is 40.0%, 400.  Maximum value is 75.0%, 750."),
  TranslateDescription("es", "")>
  Public ReserveHeatEnableLevel As Integer

  <Parameter(0, 1000), Category(section_ReserveTank), Description(""),
  TranslateDescription("es", "")> Public ReserveMixerOnLevel As Integer

  <Parameter(0, 1000), Category(section_ReserveTank), Description(""),
  TranslateDescription("es", "")> Public ReserveMixerOffLevel As Integer

  <Parameter(0, 1000), Category(section_ReserveTank), Description(""),
  TranslateDescription("es", "")> Public ReserveRinseLevel As Integer

  <Parameter(0, 1000), Category(section_ReserveTank), Description(""),
  TranslateDescription("es", "")> Public ReserveCirculateTime As Integer

  <Parameter(0, 1000), Category(section_ReserveTank), Description(""),
  TranslateDescription("es", "")> Public ReserveCirculateStopTime As Integer

  <Parameter(0, 100), Category(section_ReserveTank), Description(""),
TranslateDescription("es", "")>
  Public ReserveOverfillTime As Integer

  <Parameter(0, 60), Category(section_ReserveTank), Description(""),
  TranslateDescription("es", "")> Public ReserveRinseTime As Integer

  <Parameter(0, 60), Category(section_ReserveTank), Description(""),
  TranslateDescription("es", "")> Public ReserveRinseToDrainTime As Integer

  <Parameter(0, 60), Category(section_ReserveTank), Description(""),
  TranslateDescription("es", "")> Public ReserveRinseToDrainNumber As Integer

  <Parameter(0, 60), Category(section_ReserveTank), Description(""),
  TranslateDescription("es", "")> Public ReserveTimeAfterRinse As Integer

  <Parameter(0, 60), Category(section_ReserveTank), Description(""),
  TranslateDescription("es", "")>
  Public ReserveTimeBeforeRinse As Integer

  <Parameter(0, 60), Category(section_ReserveTank), Description(""),
  TranslateDescription("es", "")>
  Public ReserveTimeGravityTransfer As Integer

  <Parameter(0, 750), Category(section_ReserveTank), Description(""),
  TranslateDescription("es", "")> Public ReserveBackfillLevel As Integer

  <Parameter(0, 30), Category(section_ReserveTank), Description(""),
  TranslateDescription("es", "")> Public ReserveBackFillStartTime As Integer


#End Region

#Region " SETUP"

  <Parameter(0, 1), Category("Setup"),
  Description("Set to '1' to enable program step parameter changes while running."),
  TranslateDescription("es", "Establezca en '1' para permitir cambios de parámetros de paso de programa mientras se está ejecutando.")>
  Public EnableCommandChg As Integer


  <Parameter(0, 1), Category("System"),
    Description("Hibernate the controller if all optumux channels show not responding alarms for more than hibernate delay seconds.")>
  Public HibernateEnable As Integer

  <Parameter(30, 120), Category("System"),
    Description("Delay hibernate for this amount of time if the system switches to battery, seconds")>
  Public HibernateDelay As Integer


  <Parameter(0, 1), Category("Setup"),
  Description("Set to '1' to enable mimic pushbuttons for manual control."),
  TranslateDescription("es", "Establecer a '1' para habilitar mímica pulsadores para el control manual.")>
  Public EnableMimicPushbuttons As Integer

  <Parameter(0, 10000), Category("Setup"),
  Description("Resets all parameters to default value if magic value is entered."),
  TranslateDescription("es", "Restablece todos los parámetros con valor por defecto si se introduce valor mágico")>
  Public InitializeParameters As Integer

  <Parameter(0, 200), Category("Setup"),
  Description("Smoothing Rate to average the temperature inputs and reduce signal noise."),
  TranslateDescription("es", "Suavizado de tasa promedio de las entradas de temperatura y reducir el ruido de la señal.")>
  Public SmoothRate As Integer

#End Region

#Region " STANDARD TIME "

  <Parameter(0, 1000), Category("Standard Time"),
  Description("Time, in minutes, to allow the operators to prepare an addition before setting delay conditions."),
  TranslateDescription("es", "Tiempo, en minutos, para permitir que los operadores preparar una adición antes de establecer las condiciones de retraso.")>
  Public StandardTimeAddPrepare As Integer

  <Parameter(0, 1000), Category("Standard Time"),
  Description("Time, in minutes, to allow the operators to unload the machine before setting delay conditions."),
  TranslateDescription("es", "Tiempo, en minutos, para permitir que los operadores descargar la máquina antes de establecer las condiciones de retraso.")>
  Public StandardTimeApproval As Integer

  <Parameter(0, 1000), Category("Standard Time"),
  Description("Time, in minutes, to allow the operators to check pH before setting delay conditions."),
  TranslateDescription("es", "Tiempo, en minutos, para permitir que los operadores para controlar el pH antes de establecer las condiciones de retraso.")>
  Public StandardTimeCheckPH As Integer

  <Parameter(0, 1000), Category("Standard Time"),
  Description("Time, in minutes, to allow the operators to check salt before setting delay conditions."),
  TranslateDescription("es", "Tiempo, en minutos, para permitir que los operadores revisar sal antes de establecer las condiciones de retraso.")>
  Public StandardTimeCheckSalt As Integer

  <Parameter(0, 1000), Category("Standard Time"),
  Description("Time, in minutes, to allow the machine to drain before setting delay conditions."),
  TranslateDescription("es", "Tiempo, en minutos, para permitir que la máquina de desagüe antes de fijar las condiciones de retraso.")>
  Public StandardTimeDrain As Integer

  <Parameter(0, 1000), Category("Standard Time"),
  Description("Time, in minutes, to allow the machine to hot drain before setting delay conditions."),
  TranslateDescription("es", "Tiempo, en minutos, para permitir que la máquina se agote en caliente antes de establecer condiciones de retraso.")>
  Public StandardTimeDrainHot As Integer

  <Parameter(0, 1000), Category("Standard Time"),
  Description("Time, in minutes, to allow the machine to fill before setting delay conditions."),
  TranslateDescription("es", "Tiempo, en minutos, para permitir que la máquina llenar antes de establecer las condiciones de retraso.")>
  Public StandardTimeFillMachine As Integer

  <Parameter(0, 1000), Category("Standard Time"),
  Description("Time, in minutes, to allow the operators to load the machine before setting delay conditions."),
  TranslateDescription("es", "Tiempo, en minutos, para permitir que los operadores de la máquina de carga antes de establecer las condiciones de retraso.")>
  Public StandardTimeLoad As Integer

  <Parameter(0, 1000), Category("Standard Time"),
  Description("Time, in minutes, to allow the operators to respond before setting delay conditions."),
  TranslateDescription("es", "Tiempo, en minutos, para permitir que los operadores responder antes de establecer las condiciones de retraso.")>
  Public StandardTimeOperator As Integer

  <Parameter(0, 1000), Category("Standard Time"),
  Description("Time, in minutes, to allow the operators for sample approval before setting delay conditions."),
  TranslateDescription("es", "Tiempo, en minutos, para permitir que a los operadores para la aprobación de la muestra antes de establecer las condiciones de retraso.")>
  Public StandardTimeSample As Integer

  <Parameter(0, 1000), Category("Standard Time"),
  Description("Time, in minutes, to allow the Reserve Drain command to function for each Drain/Wash/Mix sequence, before setting delay conditions."),
  TranslateDescription("es", "")>
  Public StandardTimeReserveDrain As Integer

  <Parameter(0, 1000), Category("Standard Time"),
  Description("Time, in minutes, to allow the Reserve Tranfser command to function for before setting delay conditions."),
  TranslateDescription("es", "Tiempo, en minutos, para permitir que el comando Reserve Tranfser funcione antes de establecer condiciones de retardo.")>
  Public StandardTimeReserveTransfer As Integer

  <Parameter(0, 1000), Category("Standard Time"),
  Description("Time, in minutes, to allow steps to contiue before setting delay conditions."),
  TranslateDescription("es", "Tiempo, en minutos, para permitir que los pasos que contiue antes de establecer las condiciones de retraso.")>
  Public StandardTimeStepExcess As Integer

  <Parameter(0, 1000), Category("Standard Time"),
  Description("Time, in minutes, to allow the operators to unload the machine before setting delay conditions."),
  TranslateDescription("es", "Tiempo, en minutos, para permitir que los operadores descargar la máquina antes de establecer las condiciones de retraso.")>
  Public StandardTimeUnload As Integer

  <Parameter(0, 1000), Category("Standard Time"),
  Description("Set to '1', to enable seven day scheduling (full work week), else Load delays will be disregarded on sunday startup."),
  TranslateDescription("es", "Establecer a '1', para habilitar la programación de (semana de trabajo), de siete días más retrasos de carga carecerá en el arranque del domingo.")>
  Public SevenDaySchedule As Integer

#End Region

#Region " TANK 1 "

  <Parameter(0, 1000), Category("Tank 1"),
  Description("Time, in seconds, to continue transferring kitchen tank to drain after level is lost."),
  TranslateDescription("es", "Tiempo, en segundos, para continuar la transferencia de depósito cocina para drenar después de nivel se pierde.")>
  Public CKTransferToDrainTime1 As Integer

  <Parameter(0, 1000), Category("Tank 1"),
  Description("Level, in tenths, to fill the kitchen tank to when filling to clean."),
  TranslateDescription("es", "Nivel, en décimas, llenar el tanque de la cocina para cuando relleno para limpiar.")>
  Public CKFillLevel As Integer

  <Parameter(0, 1000), Category("Tank 1"),
  Description("Time, in seconds, to rinse the kitchen tank for when rinsing to clean."),
  TranslateDescription("es", "Tiempo, en segundos, enjuagar el tanque de cocina cuando enjuague para limpieza.")>
  Public CKRinseTime As Integer

  <Parameter(0, 1000), Category("Tank 1"),
  Description("Time, in seconds, to continue transferring kitchen tank to drain after level is lost."),
  TranslateDescription("es", "Tiempo, en segundos, para continuar la transferencia de depósito cocina para drenar después de nivel se pierde.")>
  Public CKTransferToDrainTime2 As Integer

  <Parameter(0, 1000), Category("Tank 1"),
  Description("Tank 1 prefill level, in tenths."),
  TranslateDescription("es", "Nivel de banda muerta, en décimas, .")>
  Public Tank1PreFillLevel As Integer



  <Parameter(0, 1000), Category("Tank 1"),
  Description("Deadband level, in tenths, to fill within the desired fill level when filling to level."),
  TranslateDescription("es", "Nivel de banda muerta, en décimas, para llenar en el nivel de llenado deseado al llenar al nivel.")>
  Public Tank1FillLevelDeadBand As Integer

  <Parameter(0, 120), Category("Tank 1"),
  Description("Tank 1 delay time before heating."),
  TranslateDescription("es", "")>
  Public Tank1HeatPrepTime As Integer




  <Parameter(0, 1000), Category("Tank 1"),
  Description("Time, in seconds, to continue transferring the kitchen tank once empty, before rinsing."),
  TranslateDescription("es", "Tiempo, en segundos, para continuar la transferencia del depósito de cocina una vez vacíos, antes de enjuagar.")>
  Public Tank1TimeBeforeRinse As Integer

  <Parameter(0, 1000), Category("Tank 1"),
  Description("Time, in seconds, to rinse the kitchen tank for when rinsing to machine."),
  TranslateDescription("es", "Tiempo, en segundos, enjuagar el tanque de cocina al lavado a máquina.")>
  Public Tank1RinseTime As Integer

  <Parameter(0, 1000), Category("Tank 1"),
  Description("Number of rinses when transferring the kitchen tank to the machine."),
  TranslateDescription("es", "Número de enjuagues al transferir el depósito de la cocina a la máquina.")>
  Public Tank1RinseNumber As Integer

  <Parameter(0, 1000), Category("Tank 1"),
  Description("Time, in seconds, to continue transferring the kitchen tank once empty, after rinsing."),
  TranslateDescription("es", "Tiempo, en segundos, para continuar transfiriendo el tanque de la cocina una vez vacía, después de enjuagar.")>
  Public Tank1TimeAfterRinse As Integer

  <Parameter(0, 1000), Category("Tank 1"),
  Description("Time, in seconds, to rinse the kitchen tank for when rinsing to drain."),
  TranslateDescription("es", "Tiempo, en segundos, enjuagar el tanque de cocina cuando enjuague para drenar.")>
  Public Tank1RinseToDrainTime As Integer

  <Parameter(0, 1000), Category("Tank 1"),
  Description("Tank 1 Level, in tenths, to fill to during final Rinse to Drain sequence.  Used to clean tank and mixer after kitchen prepare."),
  TranslateDescription("es", "1 nivel, en décimas, para rellenar a durante el enjuague final a la secuencia de drenaje del tanque.  Utiliza para limpiar el depósito y mezclador después de preparar cocina.")>
  Public Tank1RinseToDrainLevel As Integer

  <Parameter(0, 1000), Category("Tank 1"),
  Description("Time, in seconds, to continue transferring kitchen tank to drain after level is lost."),
  TranslateDescription("es", "Tiempo, en segundos, para continuar la transferencia de depósito cocina para drenar después de nivel se pierde.")>
  Public Tank1DrainTime As Integer

  <Parameter(0, 1000), Category("Tank 1"),
  Description("Kitchen level, in tenths, to enable the mixer to turn on."),
  TranslateDescription("es", "Nivel de cocina, en décimas, para activar el mezclador encender.")>
  Public Tank1MixerOnLevel As Integer

  <Parameter(0, 1000), Category("Tank 1"),
  Description("Kitchen level, in tenths, below which to turn off the mixer."),
  TranslateDescription("es", "Nivel de cocina, en décimas, que apagar la batidora.")>
  Public Tank1MixerOffLevel As Integer



  <Parameter(0, 1000), Category("Tank 1"),
  Description("Kitchen level, in tenths, "),
  TranslateDescription("es", ".")>
  Public Tank1RinseLevel As Integer

#End Region

#Region " TEMPERATURE CONTROL"
  ' Note: Refer to Temperature.vb

  <Parameter(0, 1000), Category("Temperature Control"),
    Description("Percent, in tenths, of the balance term to use in the Pid."),
  TranslateDescription("es", "Por ciento, en décimas, de la expresión de equilibrio para el Pid.")>
  Public BalancePercent As Integer

  <Parameter(0, 1000), Category("Temperature Control"),
    Description("Ambient temperatur eof the machine in tenths with the pump running."),
  TranslateDescription("es", "Temperatura ambiente EF. la máquina en décimas con la bomba en funcionamiento.")>
  Public BalanceTemperature As Integer

  <Parameter(1, 1000), Category("Temperature Control"),
     Description("The rate of integral action, while cooling, in repeats per minute."),
   TranslateDescription("es", "La tasa de acción integral, mientras se enfría, en repeticiones por minuto.")>
  Public CoolIntegral As Integer

  <Parameter(0, 100), Category("Temperature Control"),
    Description("The proportional band value, while cooling, in tenths."),
  TranslateDescription("es", "El valor de banda proporcional, mientras se enfría, en décimas.")>
  Public CoolPropBand As Integer

  <Parameter(0, 100), Category("Temperature Control"),
    Description("The maximum cooling rate that the machine can maintain, measured in degrees per minute."),
  TranslateDescription("es", "La tasa de acción integral, mientras se enfría, en repeticiones por minuto.")>
  Public CoolMaxGradient As Integer

  <Parameter(0, 100), Category("Temperature Control"),
    Description("When measured temp is within this many degrees of setpoint, while cooling, then setpoint has been achieved."),
  TranslateDescription("es", "Cuando se mide temperatura es dentro de este muchos grados de punto de referencia, mientras se enfría, entonces punto de referencia se ha logrado.")>
  Public CoolStepMargin As Integer

  <Parameter(0, 2), Category("Temperature Control"),
    Description("Sets mode change during cooling: 0 = switch on ramp or hold, 1 = switch only when holding, 2 = switch disabled."),
  TranslateDescription("es", "Establece cambiar el modo durante la refrigeración: 0 = interruptor en rampa o bodega, 1 = interruptor solamente cuando se sostiene, 2 = Interruptor desactivado.")>
  Public CoolModeChange As Integer

  <Parameter(0, 120), Category("Temperature Control"),
    Description("Delay a mode change for this time in seconds while cooling."),
  TranslateDescription("es", "Demora un cambio de modo para este tiempo en segundos mientras se enfría")>
  Public CoolModeChangeDelay As Integer

  <Parameter(0, 30), Category("Temperature Control"),
  Description("Time, in seconds, to delay opening the Cooling valve to prevent valve pulsing."),
  TranslateDescription("es", "Tiempo en segundos, para retrasar la apertura de la válvula de enfriamiento para evitar la pulsación de la válvula.")>
  Public CoolOnDelay As Integer

  <Parameter(0, 1000), Category("Temperature Control"),
  Description("PID Output while cooling above which to activate cooling coil, in tenths."),
  TranslateDescription("es", "Salida del PID al enfriamiento por encima del cual activar el serpentín de enfriamiento, en décimas")>
  Public CoolOnOutput As Integer

  <Parameter(0, 1000), Category("Temperature Control"),
    Description("Output, in tenths, to use while venting at the beginning of a cooling step."),
  TranslateDescription("es", "Salida, en décimas, uso de ventilación al comienzo de una etapa de enfriamiento.")>
  Public CoolVentOutput As Integer

  <Parameter(0, 60), Category("Temperature Control"),
  Description("At the beginning of a cooling step, the heat-exchanger is vented / purged for this many seconds."),
  TranslateDescription("es", "Al principio de una etapa de enfriamiento, el intercambiador de calor con ventilación / purgado por tantos segundos.")>
  Public CoolVentTime As Integer

  <Parameter(0, 600), Category("Temperature Control"),
  Description("Temperature in tenths of a degree that the machine will cool to when the crash cool button is pushed."),
  TranslateDescription("es", "Temperatura en décimas de grado que la máquina se enfríe a cuando se presiona el botón accidente cool.")>
  Public CrashCoolTemperature As Integer

  <Parameter(0, 100), Category("Temperature Control"),
  Description("Time, in seconds, to reduce the Gradient term from the PID output when no longer ramping.  Used to reduce error when first stop ramping.  Limited range 0-120 seconds.  Disable with value of '0'."),
  TranslateDescription("es", "Time, in seconds, to reduce the Gradient term from the PID output when no longer ramping.  Used to reduce error when first stop ramping.  Limited range 0-120 seconds.  Disable with value of '0'.")>
  Public GradientRampDownTime As Integer

  <Parameter(0, 1000), Category("Temperature Control"),
  Description("Percent of Gradient Term, in tenths, to retain of the initial Gradient term when no longer ramping.  Used to reduce error first stop ramping. Value range 0-1000 = 0.0-100.0%.  Disable with value of '0'."),
  TranslateDescription("es", "% De gradiente, en décimas, para mantener el gradiente inicial de plazo cuando ya no en rampa.  Utilizado para reducir el error primer parada en rampa. Valor de rango 0-1000 = 0.0-100.0%.  Desactivar con valor de '0'.")>
  Public GradientRampDownPercent As Integer

  <Parameter(0, 1000), Category("Temperature Control"),
  Description("The rate of integral action, while heating, in repeats per minute."),
  TranslateDescription("es", "La tasa de acción integral, mientras la calefaccion, en repeticiones por minuto.")>
  Public HeatIntegral As Integer

  <Parameter(0, 30), Category("Temperature Control"),
  Description("Time, in seconds, to delay opening the steam valve to prevent valve pulsing."),
  TranslateDescription("es", "Tiempo en segundos, para retrasar la apertura de la válvula de vapor para evitar la pulsación de la válvula.")>
  Public HeatOnDelay As Integer

  <Parameter(0, 1000), Category("Temperature Control"),
  Description("PID Output while heating above which to activate heating coil, in tenths."),
  TranslateDescription("es", "Salida de PID mientras la calefacción encima de que activar la bobina de calefacción, en décimas.")>
  Public HeatOnOutput As Integer

  <Parameter(1, 1000), Category("Temperature Control"),
  Description("The proportional band value, while heating, in tenths."),
  TranslateDescription("es", "El valor de banda proporcional, mientras que la calefacción, en décimas.")>
  Public HeatPropBand As Integer

  <Parameter(0, 100), Category("Temperature Control"),
  Description("The maximum heating rate that the machine can maintain, measured in degrees per minute."),
  TranslateDescription("es", "La máxima velocidad que puede mantener la máquina, de calentamiento se mide en grados por minuto.")>
  Public HeatMaxGradient As Integer

  <Parameter(0, 1000), Category("Temperature Control"),
  Description("The maximum heating output, in tenths.  Set to 0 to disable."),
  TranslateDescription("es", "")>
  Public HeatMaxOutput As Integer

  <Parameter(0, 2), Category("Temperature Control"),
  Description("Sets mode change during heating: 0 = switch on ramp or hold, 1 = switch only when holding, 2 = switch disabled."),
  TranslateDescription("es", "Establece cambiar el modo durante la calefacción: 0 = interruptor en rampa o bodega, 1 = interruptor solamente cuando se sostiene, 2 = Interruptor desactivado.")>
  Public HeatModeChange As Integer

  <Parameter(0, 1000), Category("Temperature Control"),
  Description("Time, in seconds, to delay switching between heating and cooling modes when output drops to 0%."),
  TranslateDescription("es", "Tiempo, en segundos, para retrasar la conmutación entre calefacción y refrigeración modos cuando salida cae a 0%.")>
  Public HeatModeChangeDelay As Integer

  <Parameter(0, 100), Category("Temperature Control"),
  Description("When measured temp is within this many degrees of setpoint, while heating, then setpoint has been achieved."),
  TranslateDescription("es", "Cuando se mide temperatura es dentro de este muchos grados de punto de referencia, mientras que la calefacción, entonces punto de referencia se ha logrado.")>
  Public HeatStepMargin As Integer

  <Parameter(0, 1000), Category("Temperature Control"),
    Description("Output, in tenths, to use while venting at the beginning of a heating step."),
  TranslateDescription("es", "Salida, en décimas, uso de ventilación al comienzo de una etapa de calefacción.")>
  Public HeatVentOutput As Integer

  <Parameter(0, 60), Category("Temperature Control"),
    Description("At the beginning of a heating step, the heat-exchanger is vented / purged for this many seconds."),
  TranslateDescription("es", "Al principio de un paso de la calefacción, el intercambiador de calor con ventilación / purgado por tantos segundos.")>
  Public HeatVentTime As Integer

  <Parameter(0, 120), Category("Temperature Control"),
  Description("Time, in seconds, used to create Steam Output Factor to prevent Steam output from peaking at initial demand."),
  TranslateDescription("es", "Tiempo, en segundos, utilizado para crear el Factor de salida de vapor para evitar que el vapor de salida de un pico en la demanda inicial.")>
  Public SteamValveDelay As Integer

  <Parameter(0, 1000), Category("Temperature Control"),
    Description("The temperature bandwidth in tenths of a degree. If the actual temp is this many degrees higher/lower than the setpoint temp, an alarm will occur."),
  TranslateDescription("es", "El ancho de banda de temperatura en décimas de grado. Si la temperatura real es tantos grados superior e inferior que la temperatura de consigna, se producirá una alarma.")>
  Public TemperatureAlarmBand As Integer

  <Parameter(0, 600), Category("Temperature Control"),
    Description("The delay time in seconds before an alarm will occur if the Temp Alarm Band is exceed."),
  TranslateDescription("es", "Exceder el tiempo de retardo en segundos antes de que una alarma se producirá si la banda de la alarma de Temp.")>
  Public TemperatureAlarmDelay As Integer

  <Parameter(0, 1000), Category("Temperature Control"),
    Description("If the difference between the setpoint and the measured temp is greater than this value, then stop the setpoint changing."),
  TranslateDescription("es", "Si la diferencia entre la consigna y la temperatura medida es superior a este valor, entonces parar el cambio del punto de ajuste.")>
  Public TemperaturePidPause As Integer

  <Parameter(0, 1000), Category("Temperature Control"),
    Description("If the temp gradient has been paused, the measured value must be within this many degrees of the setpoint before the gradient can restart."),
  TranslateDescription("es", "Si el gradiente de temperatura ha sido detenido, el valor medido debe ser dentro de esta muchos grados de la consigna antes de reinicia el gradiente.")>
  Public TemperaturePidRestart As Integer

  <Parameter(0, 1000), Category("Temperature Control"),
    Description("If the difference between the setpoint and the measured temp is greater than this value, then reset the temp control."),
  TranslateDescription("es", "Si la diferencia entre la consigna y la temperatura medida es mayor que este valor, entonces restablecer el control de temperatura.")>
  Public TemperaturePidReset As Integer


#If 0 Then
  ' TODO - Temperature.vb vb6

  Public Parameters_HeatPidCalculateDelay As Long
Attribute HeatPidCalculateDelay.VB_VarDescription = "Category=Temperature Control\r\nHelp=Time inteveral, in seconds, to delay between PID calculations when heating.  Use to minimize output changing on small scale adjustments, increase valve life. Set to '0' to disable."
  Public HeatPidDelayTimer As New acTimer
Public Parameters_HeatCoolOutputDeltaRamp As Long
Attribute HeatCoolOutputDeltaRamp.VB_VarDescription = "Category=Temperature Control\r\nHelp=Necessary difference between existing setpoint and new setpoint before HeatCool output will adjust to new value, while heating or cooling to a final setpoint.  Set to '0' to disable function.  Maximum value 100 (10.0%)."
Public Parameters_HeatCoolOutputDeltaHold As Long
Attribute HeatCoolOutputDeltaHold.VB_VarDescription = "Category=Temperature Control\r\nHelp=Necessary difference between existing setpoint and new setpoint before HeatCool output will adjust to new value, while holding after reaching a final setpoint.  Set to '0' to disable function.  Maximum value 100 (10.0%)"

#End If


#End Region

#Region " WORKING LEVEL SETUP PACKAGE TYPE "

  ' TODO
  <Parameter(0, 1000), Category("Working Level")> Public NumberOfSpindales As Integer
  <Parameter(0, 1000), Category("Working Level")> Public EnableFillByWorkingLevel As Integer
  <Parameter(0, 1000), Category("Working Level")> Public EnableReserveFillByWorkingLevel As Integer

  ''' TYPE 0
  <Parameter(0, 1000), Category("Working Level"),
  Description("Package Type 0 - Working Level, in tenths percent, for package height."),
  TranslateDescription("es", "")>
  Public PackageType0Level(9) As Integer

  ''' TYPE 1
  <Parameter(0, 1000), Category("Working Level"),
  Description("Package Type 1 - Working Level, in tenths percent, for package height of 1."),
  TranslateDescription("es", "")>
  Public PackageType1Level(9) As Integer

  ''' TYPE 2
  <Parameter(0, 1000), Category("Working Level"),
  Description("Package Type 2 - Working Level, in tenths percent, for package height of 1."),
  TranslateDescription("es", "")>
  Public PackageType2Level(9) As Integer

  ''' TYPE 3
  <Parameter(0, 1000), Category("Working Level"),
  Description("Package Type 3 - Working Level, in tenths percent, for package height of 1."),
  TranslateDescription("es", "")>
  Public PackageType3Level(9) As Integer

  ''' TYPE 4
  <Parameter(0, 1000), Category("Working Level"),
  Description("Package Type 4 - Working Level, in tenths percent, for package height of 1.")>
  Public PackageType4Level(9) As Integer

  ''' TYPE 5
  <Parameter(0, 1000), Category("Working Level"),
    Description("Package Type 5 - Working Level, in tenths percent, for package height of 1.")>
  Public PackageType5Level(9) As Integer

  ''' TYPE 6
  <Parameter(0, 1000), Category("Working Level"),
    Description("Package Type 6 - Working Level, in tenths percent, for package height of 1.")>
  Public PackageType6Level(9) As Integer

  ''' TYPE 7
  <Parameter(0, 1000), Category("Working Level"),
    Description("Package Type 7 - Working Level, in tenths percent, for package height of 1.")>
  Public PackageType7Level(9) As Integer

  ''' TYPE 8
  <Parameter(0, 1000), Category("Working Level"),
    Description("Package Type 8 - Working Level, in tenths percent, for package height of 1.")>
  Public PackageType8Level(9) As Integer

  ''' TYPE 9
  <Parameter(0, 1000), Category("Working Level"),
    Description("Package Type 9 - Working Level, in tenths percent, for package height of 1.")>
  Public PackageType9Level(9) As Integer

#End Region




  Sub MaybeSetDefaults()
    If (InitializeParameters = 9387) Or (InitializeParameters = 8741) Or
      (InitializeParameters = 8742) Or (InitializeParameters = 8743) Or
      (InitializeParameters = 8744) Then
      SetDefaults()
    End If
  End Sub

  Private Sub SetDefaults()

    ' Alarm Configuration
    EnableSiren = 0
    PLCComsLossDisregard = 0
    PLCComsTime = 10
    '  WatchdogTimeout = 10

    ' Calibration - Level Transmitters
    VesselLevelTransmitterMax = 1000
    VesselLevelTransmitterMin = 10

    Tank1LevelTransmitterMax = 1000
    Tank1LevelTransmitterMin = 30

    ' Calibration - Pressure Transmitters
    PackageDiffPress100 = 0
    PackageDiffPressEnable = 0
    PackageDiffPressZero = 0
    PackageDiffPressRange = 0
    VesselPressureMax = 900
    VesselPressureMin = 0
    VesselPressureRange = 100000

    ' Calibration - Temperature Transmitters
    Tank1TempProbe = 0

    ' Demo Mode
    Demo = 0

    ' Flow Control
    SystemVolumeAt0Level = 100      ' 117
    SystemVolumeAt100Level = 500    ' 377

    ' Level Control - Draining
    DrainMachineTime = 120
    HDDrainTime = 120
    HDFillRunTime = 120
    HDHiFillLevel = 500
    HDVentAlarmTime = 120
    IdleDrainTime = 30

    ' Level Control - Filling



    FillLevelMaximum = 900
    FillLevelMinimum = 200
    FillOpenDelayTime = 5
    FillPrimePumpTime = 5
    FillSettleTime = 5
    LevelGaugeFlushTime = 10
    LevelGaugeSettleTime = 10
    MaxLevelLossAmount = 200
    OverFillTime = 0
    RecordLevelTimer = 120


    ' Look Ahead
    EnableRedyeIssueAlarm = 0
    LookAheadEnabled = 0
    LookAheadIgnoreBlocked = 0

    ' Pressure Control
    AirpadBleedPressure = 34000
    AirpadCutoffPressure = 15000

    ' Production Reports
    MainPumpMotorHP = 0

    ' Pump Control
    With controlCode.PumpControl
      .Parameters_FlowReverseSwitchingTime = 10
      .Parameters_PumpAccelerationTime = 10
      .Parameters_PumpEnableTime = 5
      .Parameters_PumpMinimumLevel = 150
      .Parameters_PumpMinimumLevelTime = 5
    End With

    ' Safey Control
    With controlCode.SafetyControl
      .Parameters_DePressurizationTemperature = 800
      .Parameters_DePressurizationTime = 30
      .Parameters_PressurizationTemperature = 820
    End With

    ' Setup
    EnableCommandChg = 1
    SmoothRate = 10

    ' Standard Times
    SevenDaySchedule = 0
    StandardTimeAddPrepare = 10
    StandardTimeApproval = 10
    StandardTimeCheckSalt = 10
    StandardTimeCheckPH = 10
    StandardTimeDrain = 10
    StandardTimeFillMachine = 10
    StandardTimeLoad = 10
    StandardTimeOperator = 10
    StandardTimeSample = 30
    StandardTimeStepExcess = 10
    StandardTimeUnload = 10

    ' Tank 1
    CKFillLevel = 100
    CKRinseTime = 10
    CKTransferToDrainTime1 = 15
    CKTransferToDrainTime2 = 15
    Tank1DrainTime = 15
    Tank1FillLevelDeadBand = 10
    Tank1MixerOffLevel = 100
    Tank1MixerOnLevel = 150
    Tank1RinseNumber = 2
    Tank1RinseTime = 10
    Tank1RinseToDrainTime = 20
    Tank1TimeAfterRinse = 10
    Tank1TimeBeforeRinse = 10

    ' Temperature Control
    BalancePercent = 100
    BalanceTemperature = 800
    CoolIntegral = 50
    CoolMaxGradient = 10
    CoolModeChange = 0
    CoolModeChangeDelay = 20
    CoolPropBand = 150
    CoolStepMargin = 30
    CoolVentTime = 10
    CrashCoolTemperature = 1600
    HeatIntegral = 100
    HeatMaxGradient = 10
    HeatMaxOutput = 500
    HeatModeChange = 0
    HeatModeChangeDelay = 30
    HeatPropBand = 100
    HeatStepMargin = 20
    HeatVentTime = 5
    SteamValveDelay = 5
    TemperatureAlarmBand = 120
    TemperatureAlarmDelay = 300
    TemperaturePidPause = 100
    TemperaturePidReset = 250
    TemperaturePidRestart = 400


    ' Clear Initialize Flag
    InitializeParameters = 0
  End Sub






#If 0 Then
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
  
#End If



End Class
