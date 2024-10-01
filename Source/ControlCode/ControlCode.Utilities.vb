Partial Class ControlCode

  Private Sub CalculateUtilities()
    'Utility usage calculations

    Dim HPNow As Integer
    'PowerFactor = (750 * 68) / 100
    HPNow = 0 'Reset power being used now uasge to 0
    Static UtilitiesTimer As New Timer
    If UtilitiesTimer.Finished Then

      If IO.PumpRunning Then HPNow = +Parameters.MainPumpMotorHP
      If IO.AddPumpRunning Then HPNow = +Parameters.AddFeedPumpMotorHP

      'Assume all motors run at about 68% of rated capacity. 750W is equiv to HP.
      'Power factor = 68% of 750 , convert to watts
      'PowerFactor = 510 'watts which is .51kw
      PowerKWS += Convert.ToInt32(HPNow * 0.51)

      UtilitiesTimer.Seconds = 1
    End If
    If PowerKWS >= 3600 Then
      PowerKWH += 1
      PowerKWS -= 3600
    End If

    'Steam Consumption formula
    '120% * WorkingVolume * deg F of temp rise * weight of water (8.33lb/g) * 0.001 = lbs of steam used
    'Steam factor = 120% * 8.33 * 0.001 = 1/100
    'lbs of Steam used = working volume * Temp rise in F / 100
    FinalTemp = TemperatureControl.FinalTemp
    If FinalTemp <> FinalTempWas Then
      If Temp > FinalTempWas Then
        StartTemp = Temp
      Else
        StartTemp = FinalTempWas
      End If
      If FinalTemp > StartTemp Then
        TempRise = FinalTemp - StartTemp
        SteamNeeded = (VolumeBasedOnLevel * TempRise) \ 1000 ' (MachineVolume * TempRise) \ 1000
        SteamUsed += SteamNeeded
      End If
      FinalTempWas = FinalTemp
    End If
  End Sub


#If 0 Then ' TODO VB6 check

  
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



#End If

End Class
