Partial Class ControlCode
  'These variables are set by a remote client usually from a mimic
  '   we catch the variable going true immediately reset to false then call the appropriate function
  '   probably better this way rather than allowing direct remote calls 'cos then we can disable them here if necessary

  Private Sub RunRemoteFunctions()
    RunRemoteControlFuntions()
    RunRemoteAddFuntions()
    RunRemoteReserveFunctions()

  End Sub

#Region " LOCAL CONTROL FUNCTIONS "
  Public Remote_MainPumpReset As Boolean
  Public Remote_SteamTest As Boolean
  '  Public ManualSteamOverride As Boolean
  '  Public ManualSteamTest As New ACManualSteamTest

  Private Sub RunRemoteControlFuntions()

  End Sub
#End Region

#Region " ADD TANK FUNCTIONS "

  Public Remote_AddReady As Boolean
  Public Remote_AddFillHot As Boolean : Public Remote_AddFillLevel As Integer
  Public Remote_AddFillCold As Boolean
  Public Remote_AddDrain As Boolean
  Public Remote_AddTransfer As Boolean

  Private Sub RunRemoteAddFuntions()

    If Remote_AddReady Then
      Remote_AddReady = False : AddReady = Not AddReady
    End If

    With AddControl

      If Remote_AddFillCold Then
        Remote_AddFillCold = False : .FillManual(EFillType.Cold, Remote_AddFillLevel)
      End If
      If Remote_AddFillHot Then
        Remote_AddFillHot = False : .FillManual(EFillType.Hot, Remote_AddFillLevel)
      End If

    End With
  End Sub

#End Region

#Region " RESERVE TANK FUNCTIONS "

  Public Remote_ReserveReady As Boolean
  Public Remote_ReserveFillHot As Boolean : Public Remote_ReserveFillLevel As Integer
  Public Remote_ReserveFillCold As Boolean
  Public Remote_ReserveDrain As Boolean
  Public Remote_ReserveTransfer As Boolean
  Public Remote_ReserveHeat As Boolean : Public Remote_ReserveHeatTemp As Integer

  Private Sub RunRemoteReserveFunctions()

    If Remote_ReserveReady Then
      Remote_ReserveReady = False : ReserveReady = Not ReserveReady
    End If

    With ReserveControl

    End With
  End Sub

#End Region

  ' TODO
#If 0 Then

#End If



End Class
