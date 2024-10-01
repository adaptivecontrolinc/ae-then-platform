Public Class Simulation
  Private ReadOnly controlCode As ControlCode
  Private demoSetup_ As Boolean = False

  Public FillRate As Integer = 10
  Public DrainRate As Integer = 12
  Public HeatRate As Integer = 12
  Public CoolRate As Integer = 10
  Public TransferRate As Integer = 10

  Private ReserveLevelDemo As Integer
  Private AddLevelDemo As Integer
  Private Tank1LevelDemo As Integer
  Private FlowRateDemo As Integer = 10

  Private MachinePressureDemo As Integer

  Public FillTick As Integer = 24

  Public MainTimer As New Timer With {.Seconds = 4}
  Sub New(controlCode As ControlCode)
    Me.controlCode = controlCode
  End Sub

  Sub Run()
    SimulateMachine(MainTimer.Finished)
    SimulateLevelTank1()

    If MainTimer.Finished Then MainTimer.Seconds = 1
  End Sub

  Private Sub SimulateMachine(tick As Boolean)
    With controlCode
      If tick Then

        .IO.AddPumpRunning = .IO.AddPumpStart AndAlso .IO.PowerOn AndAlso Not .IO.EmergencyStop
        .IO.PumpRunning = .IO.MainPump AndAlso .IO.PowerOn AndAlso Not .IO.EmergencyStop



        If .IO.Fill Then
          .IO.VesselLevelInput = CShort(.IO.VesselLevelInput + FillRate)
          ' Update temperature if not heating or below 75F
          If Not (.IO.HeatSelect) AndAlso (.IO.HeatCoolOutput > 150) Then    ' 0-500F: 0.15% x 500 = 75.0F
            .IO.VesselTemp -= CShort(5)
          End If
        End If

        ' Simulate Reserve Level
        ReserveLevelDemo = SimulateLevelReserve()
        .IO.ReserveLevelInput = CShort(MinMax(.IO.ReserveLevelInput + ReserveLevelDemo, 0, 1000))

        ' Simulate Addition Level
        AddLevelDemo += SimulateLevellAdd()
        .IO.AddLevelInput = CShort(MinMax(.IO.AddLevelInput + AddLevelDemo, 0, 1000))

        ' Tank 1 Level  
        Tank1LevelDemo = SimulateLevelTank1()
        .IO.Tank1LevelInput = CShort(MinMax(.IO.Tank1LevelInput + Tank1LevelDemo, 0, 1000))

        ' Simulate Tank 1 Heating
        If .IO.Tank1Steam Then
          .IO.Tank1Temp = CShort(.IO.Tank1Temp + 5)
        Else
          If .IO.Tank1Temp > 700 Then .IO.Tank1Temp = (CShort(.IO.Tank1Temp - 2))
        End If


        ' SIMULATE MACHINE PRESSURE
        If .IO.AirPad Then
          If .IO.Vent Then
            MachinePressureDemo -= 1
          Else : MachinePressureDemo += 5
          End If
        Else
          If .IO.Vent Then MachinePressureDemo -= 5
          If .IO.MachineDrainHot Then MachinePressureDemo -= 2
        End If
        If MachinePressureDemo < .Parameters.VesselPressureMin Then MachinePressureDemo = .Parameters.VesselPressureMin
        If MachinePressureDemo > .Parameters.VesselPressureMax Then MachinePressureDemo = .Parameters.VesselPressureMax
        .IO.VesselPressInput = Convert.ToInt16(MachinePressureDemo)


        'Simulate FlowRate based on Filling & Rinsing
        If .IO.PumpRunning Then
          If Not .IO.FlowReverse Then
            ' DP : Increases from Zero (500 > 900)
            .IO.DiffPressInput = CShort(.IO.DiffPressInput + Convert.ToInt16((1 * (.IO.PumpSpeedOutput / 1000))))
          Else
            .IO.DiffPressInput = CShort(.IO.DiffPressInput - Convert.ToInt16((1 * (.IO.PumpSpeedOutput / 1000))))
            ' DP : Decreases from Zero (500 > 100)
          End If
          ' Flowrate
          .IO.FlowRateInput = CType(Math.Min(10 * .IO.PumpSpeedOutput, Short.MaxValue), Short)
        Else
          ' DP : Bring back to zero
          If .IO.DiffPressInput >= 510 Then
            .IO.DiffPressInput -= CShort(10)
          ElseIf .IO.DiffPressInput <= 490 Then
            .IO.DiffPressInput += CShort(10)
          End If
          .IO.FlowRateInput = 0
        End If



        'Simulate Heating, Cooling Temp 
        Static rtdTemperature As Integer
        Static rtdWas As Integer
        Static randomNumber As New Random
        Dim wasSteamControl As New Timer
        Dim tempColdWater As Integer = 700
        Dim tempHotWater As Integer = 1400

        If .IO.VesselTemp <> rtdWas Then
          'We've manually adjusted the temperature 
          rtdTemperature = .IO.VesselTemp
        End If

        If .IO.HeatSelect Then
          If .IO.HeatCoolOutput > 0 Then

          End If
          If .TemperatureControl.IoOutput > 0 Then
            wasSteamControl.Seconds = 5
          Else
            rtdTemperature -= 2
          End If
          If Not wasSteamControl.Finished Then
            rtdTemperature += Convert.ToInt16((randomNumber.NextDouble * (.TemperatureControl.IoOutput / 1000)) + 1)
          Else
            rtdTemperature += Convert.ToInt16((randomNumber.NextDouble * (.TemperatureControl.IoOutput / 1000)))
          End If

        ElseIf Not wasSteamControl.Finished Then
          rtdTemperature += Convert.ToInt16(randomNumber.NextDouble / 1000)
        Else
        End If
        If .IO.CoolSelect Then
          rtdTemperature -= Convert.ToInt16((randomNumber.NextDouble * (.TemperatureControl.IoOutput / 1000)))
        End If

        If rtdTemperature < 600 Then rtdTemperature = 600
        If rtdTemperature > 2750 Then rtdTemperature = 2750
        .IO.VesselTemp = Convert.ToInt16(rtdTemperature + 5)
        rtdWas = .IO.VesselTemp



        ' First Demo 
        If Not demoSetup_ Then
          .IO.PowerOn = True
          .IO.EmergencyStop = False
          .IO.KierLidClosed = True
          .IO.LockingPinRaised = True

          'Initialize Temperature Inputs
          If .IO.VesselTemp < 650 Then .IO.VesselTemp = 650
          If .IO.ReserveTemp < 650 Then .IO.ReserveTemp = 650
          '      If IO.AddTemp < 650 Then .IO.AddTemp = 650
          If .IO.FillTemp < 650 Then .IO.FillTemp = 650
          If .IO.DiffPressInput = 0 Then .IO.DiffPressInput = 501

          If .IO.VesselTemp < 910 Then .IO.TempInterlock = True
          If .IO.VesselTemp > 1400 Then .IO.ContactTemp = True
          If .MachinePressure < (4 * Utilities.Conversions.PsiToBar) Then .IO.PressureInterlock = True


          demoSetup_ = True
        End If

      End If
    End With
  End Sub

  Private Function SimulateLevelReserve() As Integer
    Dim returnValue As Integer = 0
    With controlCode
      If .IO.ReserveDrain Then returnValue = -4
      If .IO.ReserveFillCold Then returnValue = +5
      If .IO.ReserveFillHot Then returnValue = +5
      If .IO.ReserveTransfer Then
        If .RT.IsTransfer AndAlso .IO.PumpRunning AndAlso (.ReserveLevel > 0) Then returnValue = -2
        If .RF.IsRunback AndAlso (.MachineLevel > .ReserveLevel) Then returnValue = +2
      End If
      If .IO.Tank1Transfer AndAlso .IO.Tank1TransferToReserve AndAlso (.Tank1Level > 0) Then returnValue = +2
    End With
    Return returnValue
  End Function

  Private Function SimulateLevellAdd() As Integer
    Dim returnValue As Integer = 0
    With controlCode
      If .IO.AddDrain Then returnValue = -4
      If .IO.AddFillCold Then returnValue = +5
      '  If .IO.AddFillCold AndAlso (IO.FillColdSelect OrElse IO.FillHotSelect) Then returnValue= +5
      '  If IO.AddRateOutput > 0 Then Return CInt(IO.AddRateOutput / 1000) * -1
      If .IO.AddTransfer AndAlso .IO.AddPumpRunning AndAlso (.AddLevel > 0) Then returnValue = -2
      If .IO.Tank1Transfer AndAlso .IO.Tank1TransferToAdd AndAlso (.Tank1Level > 0) Then returnValue = +2

    End With
    Return returnValue
  End Function

  Private Function SimulateLevelTank1() As Integer
    Dim returnValue As Integer = 0
    With controlCode
      If .IO.Tank1Transfer AndAlso (.Tank1Level > 0) Then
        If .IO.Tank1TransferToDrain Then returnValue = -4
        If .IO.Tank1TransferToAdd Then returnValue = -2
        If .IO.Tank1TransferToReserve Then returnValue = -2
      End If
      If .IO.Tank1FillCold Then returnValue = +5
      If .IO.Tank1FillHot Then returnValue = +3

    End With
    Return returnValue
  End Function


#If 0 Then

  
    Private Sub Demo()
      If (Parameters.Demo > 0) AndAlso (Parent.Mode = Mode.Debug) Then

        'Initialize Control Voltage & Safe Interlock inputs, as long as EStop is not active
        If Not (IO.PowerOn AndAlso IO.EmergencyStop) Then
          IO.PowerOn = True
        End If

        'Initialize Temperature Inputs


        Static machineLevel As Integer
        Static addLevel As Integer
        Static reserveLevel As Integer
        Static tank1Level As Integer
        Static randomNumber As New Random
        Dim wasSteamControl As New Timer

        'Run this code every 250ms i.e. four times a second      
        Static Timer As New Timer
        If Timer.Finished Then



          'Simulate FlowRate based on Filling & Rinsing
          Static flowRateBath As Integer
          If IO.PumpRunning Then
            Dim nextDouble As Double = randomNumber.NextDouble / 2
            flowRateBath = Convert.ToInt16((nextDouble * IO.PumpSpeedOutput * Parameters.FlowRateMachineMax / 1000))
            If flowRateBath < 0 Then flowRateBath = 0
            If flowRateBath > Parameters.FlowRateMachineMax Then flowRateBath = Parameters.FlowRateMachineMax
            IO.FlowRateInput = Convert.ToInt16((IO.FlowRateInput + flowRateBath) / 2)
          End If
  
          ' SIMULATE MACHINE PRESSURE
          Static machinePressure As Integer
          If IO.AirPad Then
            If IO.Vent Then
              machinePressure -= 1
            Else : machinePressure += 5
            End If
          Else
            If IO.Vent Then machinePressure -= 5
            If IO.MachineDrainHot Then machinePressure -= 2
          End If
          If machinePressure < 0 Then machinePressure = 0
          If machinePressure > Parameters.VesselPressureMax Then machinePressure = Parameters.VesselPressureMax
          IO.VesselPressInput = Convert.ToInt16(machinePressure)



          'Reset Demo Timer
          Timer.Milliseconds = 500
        End If

      End If
    End Sub



  
    Private Function SimulateLevelVessel() As Integer
      If IO.MachineDrain Then Return -4
      If IO.MachineDrainHot Then Return -10
      If (RC.IoFillCold OrElse RC.IoFillHot) Then
        If (MachineLevel > 750) Then Return -2
        If (MachineLevel < 750) Then Return +5
      End If
      If (RH.IoFillCold OrElse RH.IoFillHot) Then
        If (MachineLevel > 750) Then Return -2
        If (MachineLevel < 750) Then Return +5
      End If
      If (RI.IoFillCold OrElse RI.IoFillHot) Then
        If (MachineLevel > 750) Then Return -2
        If (MachineLevel < 750) Then Return +5
      End If
      If IO.Fill Then
        '  If IO.FillColdSelect Then Return +5
        '  If IO.FillHotSelect Then Return +2
        Return +5
      End If

      If IO.ReserveTransfer Then
        If RT.IsTransfer AndAlso (ReserveLevel > 0) Then
          If IO.PumpRunning Then Return +4
          Return +2
        End If
        If RF.IsRunback AndAlso (MachineLevel > ReserveLevel) Then
          Return -2
        End If
      End If

      If IO.AddTransfer AndAlso IO.AddPumpRunning AndAlso (AddLevel > 0) Then Return +2
      Return 0
    End Function

    Private Function SimulateLevelReserve() As Integer
      If IO.ReserveDrain Then Return -4
      If IO.ReserveFillCold Then Return +5
      'If IO.ReserveFillCold AndAlso (IO.FillColdSelect OrElse IO.FillHotSelect) Then Return +5
      If IO.ReserveTransfer Then
        If RT.IsTransfer AndAlso IO.PumpRunning AndAlso (ReserveLevel > 0) Then Return -2
        If RF.IsRunback AndAlso (MachineLevel > ReserveLevel) Then Return +2
      End If

      If IO.Tank1Transfer AndAlso IO.Tank1TransferToReserve AndAlso (Tank1Level > 0) Then Return +2
      Return 0
    End Function

    Private Function SimulateLevelAdd() As Integer
      If IO.AddDrain Then Return -4
      If IO.AddFillCold Then Return +5
      '  If IO.AddFillCold AndAlso (IO.FillColdSelect OrElse IO.FillHotSelect) Then Return +5
      '  If IO.AddRateOutput > 0 Then Return CInt(IO.AddRateOutput / 1000) * -1
      If IO.AddTransfer AndAlso IO.AddPumpRunning AndAlso (AddLevel > 0) Then Return -2
      If IO.Tank1Transfer AndAlso IO.Tank1TransferToAdd AndAlso (Tank1Level > 0) Then Return +2

      Return 0
    End Function

    Private Function SimulateLevelTank1() As Integer
      If IO.Tank1Transfer AndAlso (Tank1Level > 0) Then
        If IO.Tank1TransferToDrain Then Return -4
        If IO.Tank1TransferToAdd Then Return -2
        If IO.Tank1TransferToReserve Then Return -2
      End If
      If IO.Tank1FillCold Then Return +5
      Return 0
    End Function

    Private Function SimulateMachinePressure() As Integer
      If IO.AirPad Then Return +5
      If IO.Vent Then Return -2
      Return 0
    End Function


#End If
End Class
