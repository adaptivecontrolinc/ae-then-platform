' Version 2015-08-30
'===============================================================================================
'Steam Request Pushbutton for PM Test - Only Visible on mimic while in Override Mode
'===============================================================================================
Public Class ManualSteamTest : Inherits MarshalByRefObject

  ' For convenience
  Private Property parent As ACParent
  Private Property controlCode As ControlCode

  Private _stateTimer As New Timer
  Private _alarmTimer As New Timer
  ' Public Timer Value Display
  Public StateTimerSecs As Integer = _stateTimer.Seconds
  Public AlarmTimerSecs As Integer = _alarmTimer.Seconds

  Private slowFlasher As New Flasher(1500)

  Public Enum EStateVessel
    Off
    Interlock
    FillStart 'TODO - mimic FI command to open topwash for p_ delay timer
    Fill
    Heat
    Hold
  End Enum
  Public Enum EStateKitchen
    Off
    Interlock
    Fill
    Heat
    Hold
  End Enum
  Public Enum EStateReserve
    Off
    Interlock
    Fill
    Heat
    Hold
  End Enum
  Public StateVessel As EStateVessel
  Public StateTank1 As EStateKitchen
  Public StateReserve As EStateReserve
  Public IsEnabled As Boolean
  Public IsSteamRequest As Boolean

  Public ReadOnly Timer As New Timer

  Public Sub New(ByVal parent As ACParent, ByVal controlCode As ControlCode)
    Me.parent = parent
    Me.controlCode = controlCode
    Me.Timer.Seconds = 10
  End Sub

  Public Sub Run(KierLevel As Long, ReserveLevel As Long, DrugroomLevel As Long, VesTemp As Long, Tank1Temp As Integer, ReserveTemp As Integer, SteamRequestPB As Boolean, Enabled As Boolean)
    With controlCode

      If Not (Enabled Or SteamRequestPB) Then Cancel()
      IsEnabled = Enabled
      IsSteamRequest = SteamRequestPB

      ' Kier Control
      Select Case StateVessel
        Case EStateVessel.Off
          If (Enabled And SteamRequestPB) Then StateVessel = EStateVessel.Interlock

        Case EStateVessel.Interlock
        ' TODO - Delay with interlock to begin action

        Case EStateVessel.FillStart 'TODO - mimic FI command to open topwash for p_ delay timer

        Case EStateVessel.Fill
          If (.MachineLevel > 500) Then StateVessel = EStateVessel.Heat
          'If Not (Enabled And SteamRequestPB) Then KierState = KierOff

        Case EStateVessel.Heat
          If (.IO.VesselTemp >= 1800) Then StateVessel = EStateVessel.Hold
          'If Not (Enabled And SteamRequestPB) Then KierState = KierOff

        Case EStateVessel.Hold
          If (.IO.VesselTemp < 1750) Then StateVessel = EStateVessel.Heat
          'If Not (Enabled And SteamRequestPB) Then KierState = KierOff

      End Select


      ' Kitchen Control
      Select Case StateTank1
        Case EStateKitchen.Off
          If (Enabled And SteamRequestPB) Then StateTank1 = EStateKitchen.Interlock

        Case EStateKitchen.Interlock
          ' TODO - Delay with interlock to begin action

        Case EStateKitchen.Fill
          If (.Tank1Level > 500) Then StateTank1 = EStateKitchen.Heat
          'If Not (Enabled And SteamRequestPB) Then StateTank1 = EStateKitchen.Off

        Case EStateKitchen.Heat
          If (.IO.Tank1Temp >= 1800) Then StateTank1 = EStateKitchen.Hold
          'If Not (Enabled And SteamRequestPB) Then StateTank1 = EStateKitchen.Off

        Case EStateKitchen.Hold
          If (.IO.Tank1Temp < 1750) Then StateTank1 = EStateKitchen.Heat
          If Not (Enabled And SteamRequestPB) Then StateTank1 = EStateKitchen.Off

      End Select

      ' Reserve Control
      Select Case StateReserve
        Case EStateReserve.Off
          If (Enabled And SteamRequestPB) Then StateReserve = EStateReserve.Interlock

        Case EStateReserve.Interlock
          ' TODO - Delay with interlock to begin action

        Case EStateReserve.Fill
          If (.ReserveLevel > 500) Then StateReserve = EStateReserve.Heat
          If Not (Enabled And SteamRequestPB) Then StateReserve = EStateReserve.Off

        Case EStateReserve.Heat
          If (.IO.ReserveTemp >= 1800) Then StateReserve = EStateReserve.Hold
          If Not (Enabled And SteamRequestPB) Then StateReserve = EStateReserve.Off

        Case EStateReserve.Hold
          '  If Not (Enabled And SteamRequestPB) Then RTState = RTOff
          If .IO.ReserveTemp < 1750 Then StateReserve = EStateReserve.Heat

      End Select

    End With
  End Sub

  Public Sub Cancel()
    StateVessel = EStateVessel.Off
    StateTank1 = EStateKitchen.Off
    StateReserve = EStateReserve.Off
  End Sub



#Region " I/O PROPERTIES "

  ReadOnly Property IoKierHeatOn() As Boolean
    Get
      Return False ' OrElse (StateVessel = EStateVessel.Heat)
    End Get
  End Property

  ReadOnly Property IoKierTopWashOpen As Boolean
    Get
      Return False ' TODO
    End Get
  End Property

  ReadOnly Property IoKierFillOn As Boolean
    Get
      Return False ' (StateVessel = EStateVessel.Fill)
    End Get
  End Property

  ReadOnly Property IoReserveFillCold As Boolean
    Get
      Return (StateReserve = EStateReserve.Fill)
    End Get
  End Property

  ReadOnly Property IoReserveHeat As Boolean
    Get
      Return (StateReserve = EStateReserve.Heat)
    End Get
  End Property

  ReadOnly Property IoTank1FillOn As Boolean
    Get
      Return (StateTank1 = EStateKitchen.Fill)
    End Get
  End Property

  ReadOnly Property IoTank1Heat As Boolean
    Get
      Return (StateTank1 = EStateKitchen.Heat)
    End Get
  End Property

#End Region

End Class