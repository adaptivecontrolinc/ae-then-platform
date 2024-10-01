' Version 2024-06-26

' Version 2022-04-13
'  Includes Translation Function

#If 0 Then
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acSafetyControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
 'Pressure/Safety control

Option Explicit

Private Enum EPressureState
  Pressurised
  Depressurising
  Depressurised
End Enum
Private State As EPressureState
Private PressTimer As New acTimer
  
Private PressurizationTemperature As Long, DePressurizationTemperature As Long, _
        DePressurizationTime As Long
Private Sub Class_Initialize()
'Set defaults
'  PressurizationTemperature = ControlCode.ParamPressurizationTemperature
'  Maximum PressurizationTemperature, 2050
'  DePressurizationTemperature = ControlCode.ParamDepressurizationTemperature
'  Maximum DePressurizationTemperature, 1950
'  DePressurizationTime = ControlCode.ParamDepressurizationTime
  PressurizationTemperature = 2050
  DePressurizationTemperature = 2040
  DePressurizationTime = 60
  PressTimer = 60
End Sub
Public Sub Run(ByVal CurrentTempInTenths As Long, ByVal TempSafe As Boolean, _
               ByVal PressSafe As Boolean, ByVal EarlyPressure As Boolean)
'Usual pressure State control
  If (CurrentTempInTenths > PressurizationTemperature) Or (Not TempSafe) Or EarlyPressure Then
    State = Pressurised
  End If
  
  If (CurrentTempInTenths > DePressurizationTemperature) Or (Not TempSafe) Or EarlyPressure Then
    PressTimer = DePressurizationTime
  Else
    If PressTimer Or (Not PressSafe) Then
      State = Depressurising
    Else
      State = Depressurised
    End If
  End If
End Sub
' Get and Let for the parameters
Public Property Get Parameters_PressurizationTemperature() As Long
Attribute Parameters_PressurizationTemperature.VB_Description = "Category=Temperature Control\r\nHelp=The temperature in F at which the machine will automatically pressurize."
  Parameters_PressurizationTemperature = PressurizationTemperature
End Property
Public Property Let Parameters_PressurizationTemperature(ByVal TemperatureInTenths As Long)
  PressurizationTemperature = TemperatureInTenths
  If PressurizationTemperature > 2050 Then PressurizationTemperature = 2050
  If PressurizationTemperature < 1700 Then PressurizationTemperature = 1700
End Property
Public Property Get Parameters_DePressurizationTemperature() As Long
Attribute Parameters_DePressurizationTemperature.VB_Description = "Category=Temperature Control\r\nHelp=The temperature in F at which the machine will automatically de-pressurize."
  Parameters_DePressurizationTemperature = DePressurizationTemperature
End Property
Public Property Let Parameters_DePressurizationTemperature(ByVal TemperatureInTenths As Long)
  DePressurizationTemperature = TemperatureInTenths
  If DePressurizationTemperature > PressurizationTemperature - 20 Then DePressurizationTemperature = PressurizationTemperature - 20
  If DePressurizationTemperature > 2030 Then DePressurizationTemperature = 2030
  If DePressurizationTemperature < 1600 Then DePressurizationTemperature = 1600
End Property
Public Property Get Parameters_DePressurizationTime() As Long
Attribute Parameters_DePressurizationTime.VB_Description = "Category=Temperature Control\r\nHelp=The time in seconds that it takes the machine to depressurize."
  Parameters_DePressurizationTime = DePressurizationTime
End Property
Public Property Let Parameters_DePressurizationTime(ByVal DePressurizationTimeInSeconds As Long)
  DePressurizationTime = DePressurizationTimeInSeconds
  If DePressurizationTime < 30 Then DePressurizationTime = 30
  If DePressurizationTime > 600 Then DePressurizationTime = 600
End Property
Public Property Get IsPressurised() As Boolean
  IsPressurised = (State = Pressurised)
End Property
Public Property Get IsDepressurised() As Boolean
  IsDepressurised = (State = Depressurised)
End Property
Public Property Get IsDepressurising() As Boolean
  IsDepressurising = (State = Depressurising)
End Property
Public Property Get DePressurizationTimer() As Long
  DePressurizationTimer = PressTimer.TimeRemaining
End Property
Public Property Get StateString() As String
  If IsDepressurised Then
    StateString = "Machine Depressurized"
  ElseIf IsDepressurising Then
    StateString = "Machine Depressurizing " & TimerString(DePressurizationTimer)
  Else
    StateString = "Machine Pressurized"
  End If
End Property




#End If

Imports Utilities.Translations

Public Class SafetyControl : Inherits MarshalByRefObject
  Private Enum EPressureState
    Pressurized
    Depressurizing
    Depressurized
  End Enum

  Public Sub New()
    pressurizeTemp_ = 2050 ' F 960 'Tenths C
    dePressurizeTemp_ = 1700 ' F 950 C
    dePressurizeTime_ = 60
    timer_.Seconds = 60
  End Sub

  Public Sub Run(ByVal currentTempInTenths As Integer, ByVal tempSafe As Boolean,
                 ByVal pressSafe As Boolean, ByVal EarlyPressure As Boolean)

    If (currentTempInTenths > pressurizeTemp_) Or (Not tempSafe) Or EarlyPressure Then
      state_ = EPressureState.Pressurized
      StateString = Translate("Pressurized")
    End If

    If (currentTempInTenths > dePressurizeTemp_) Or (Not tempSafe) Or EarlyPressure Then
      timer_.Seconds = dePressurizeTime_
    Else
      If (timer_.Seconds > 0) Or (Not pressSafe) Then
        state_ = EPressureState.Depressurizing
        StateString = Translate("Depressurizing") & (" ") & timer_.ToString
      Else
        state_ = EPressureState.Depressurized
        StateString = Translate("Depressurized")
      End If
    End If

    pressurizeTemp_ = MinMax(Parameters_PressurizationTemperature, 1700, 2050) ', 770, 960)
    dePressurizeTemp_ = MinMax(Parameters_DePressurizationTemperature, 1600, 2030) ', 710, 940)
    dePressurizeTime_ = MinMax(Parameters_DePressurizationTime, 30, 600)
  End Sub

#Region " PARAMETERS "

  <Parameter(770, 960), Category("Safety Control"),
    Description("Temp, in tenths, to pressurize the machine at."),
    TranslateDescription("es", "Temp en décimas para presurizar la máquina en")>
  Public Parameters_PressurizationTemperature As Integer

  <Parameter(710, 940), Category("Safety Control"),
    Description("Temp, in tenths, to de-pressurize the machine at."),
    TranslateDescription("es", "Temp, en décimas, de someter la máquina a.")>
  Public Parameters_DePressurizationTemperature As Integer

  <Parameter(30, 600), Category("Safety Control"),
    Description("Time, in seconds, to de-pressurize the machine."),
    TranslateDescription("es", "Tiempo en segundos de presionar a la máquina.")>
  Public Parameters_DePressurizationTime As Integer

#End Region

#Region " PROPERTIES "
  Private state_ As EPressureState

  Public Property StateString() As String

  Private timer_ As New Timer
  Public Property Timer() As Timer
    Get
      Return timer_
    End Get
    Private Set(ByVal value As Timer)
      timer_ = value
    End Set
  End Property
  Private pressurizeTemp_ As Integer
  Public Property PressurizeTemp() As Integer
    Get
      Return pressurizeTemp_
    End Get
    Set(ByVal value As Integer)
      pressurizeTemp_ = MinMax(value, 770, 960)
    End Set
  End Property
  Private dePressurizeTemp_ As Integer
  Public Property DepressurizeTemp() As Integer
    Get
      Return dePressurizeTemp_
    End Get
    Set(ByVal value As Integer)
      If dePressurizeTemp_ > (pressurizeTemp_ - 20) Then dePressurizeTemp_ = pressurizeTemp_ - 20
      dePressurizeTemp_ = MinMax(value, 710, 940)
    End Set
  End Property

  Private dePressurizeTime_ As Integer
  Public Property DepressurizeTime() As Integer
    Get
      Return dePressurizeTime_
    End Get
    Set(ByVal value As Integer)
      dePressurizeTime_ = MinMax(value, 30, 60)
    End Set
  End Property

  Public ReadOnly Property IsPressurized() As Boolean
    Get
      If state_ = EPressureState.Pressurized Then IsPressurized = True
    End Get
  End Property
  Public ReadOnly Property IsDePressurized() As Boolean
    Get
      If state_ = EPressureState.Depressurized Then IsDePressurized = True
    End Get
  End Property
  Public ReadOnly Property IsDePressurizing() As Boolean
    Get
      If state_ = EPressureState.Depressurizing Then IsDePressurizing = True
    End Get
  End Property
  Public ReadOnly Property DePressurizeTimeRemaining() As Integer
    Get
      Return timer_.Seconds
    End Get
  End Property
#End Region



#If 0 Then
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acSafetyControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
 'Pressure/Safety control

Option Explicit

Private Enum EPressureState
  Pressurised
  Depressurising
  Depressurised
End Enum
Private State As EPressureState
Private PressTimer As New acTimer
  
Private PressurizationTemperature As Long, DePressurizationTemperature As Long, _
        DePressurizationTime As Long
Private Sub Class_Initialize()
'Set defaults
'  PressurizationTemperature = ControlCode.ParamPressurizationTemperature
'  Maximum PressurizationTemperature, 2050
'  DePressurizationTemperature = ControlCode.ParamDepressurizationTemperature
'  Maximum DePressurizationTemperature, 1950
'  DePressurizationTime = ControlCode.ParamDepressurizationTime
  PressurizationTemperature = 2050
  DePressurizationTemperature = 2040
  DePressurizationTime = 60
  PressTimer = 60
End Sub
Public Sub Run(ByVal CurrentTempInTenths As Long, ByVal TempSafe As Boolean, _
               ByVal PressSafe As Boolean, ByVal EarlyPressure As Boolean)
'Usual pressure State control
  If (CurrentTempInTenths > PressurizationTemperature) Or (Not TempSafe) Or EarlyPressure Then
    State = Pressurised
  End If
  
  If (CurrentTempInTenths > DePressurizationTemperature) Or (Not TempSafe) Or EarlyPressure Then
    PressTimer = DePressurizationTime
  Else
    If PressTimer Or (Not PressSafe) Then
      State = Depressurising
    Else
      State = Depressurised
    End If
  End If
End Sub
' Get and Let for the parameters
Public Property Get Parameters_PressurizationTemperature() As Long
Attribute Parameters_PressurizationTemperature.VB_Description = "Category=Temperature Control\r\nHelp=The temperature in F at which the machine will automatically pressurize."
  Parameters_PressurizationTemperature = PressurizationTemperature
End Property
Public Property Let Parameters_PressurizationTemperature(ByVal TemperatureInTenths As Long)
  PressurizationTemperature = TemperatureInTenths
  If PressurizationTemperature > 2050 Then PressurizationTemperature = 2050
  If PressurizationTemperature < 1700 Then PressurizationTemperature = 1700
End Property
Public Property Get Parameters_DePressurizationTemperature() As Long
Attribute Parameters_DePressurizationTemperature.VB_Description = "Category=Temperature Control\r\nHelp=The temperature in F at which the machine will automatically de-pressurize."
  Parameters_DePressurizationTemperature = DePressurizationTemperature
End Property
Public Property Let Parameters_DePressurizationTemperature(ByVal TemperatureInTenths As Long)
  DePressurizationTemperature = TemperatureInTenths
  If DePressurizationTemperature > PressurizationTemperature - 20 Then DePressurizationTemperature = PressurizationTemperature - 20
  If DePressurizationTemperature > 2030 Then DePressurizationTemperature = 2030
  If DePressurizationTemperature < 1600 Then DePressurizationTemperature = 1600
End Property
Public Property Get Parameters_DePressurizationTime() As Long
Attribute Parameters_DePressurizationTime.VB_Description = "Category=Temperature Control\r\nHelp=The time in seconds that it takes the machine to depressurize."
  Parameters_DePressurizationTime = DePressurizationTime
End Property
Public Property Let Parameters_DePressurizationTime(ByVal DePressurizationTimeInSeconds As Long)
  DePressurizationTime = DePressurizationTimeInSeconds
  If DePressurizationTime < 30 Then DePressurizationTime = 30
  If DePressurizationTime > 600 Then DePressurizationTime = 600
End Property
Public Property Get IsPressurised() As Boolean
  IsPressurised = (State = Pressurised)
End Property
Public Property Get IsDepressurised() As Boolean
  IsDepressurised = (State = Depressurised)
End Property
Public Property Get IsDepressurising() As Boolean
  IsDepressurising = (State = Depressurising)
End Property
Public Property Get DePressurizationTimer() As Long
  DePressurizationTimer = PressTimer.TimeRemaining
End Property
Public Property Get StateString() As String
  If IsDepressurised Then
    StateString = "Machine Depressurized"
  ElseIf IsDepressurising Then
    StateString = "Machine Depressurizing " & TimerString(DePressurizationTimer)
  Else
    StateString = "Machine Pressurized"
  End If
End Property




#End If
End Class
