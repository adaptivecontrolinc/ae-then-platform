Public Class acTemperatureControlContacts
#If 0 Then
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acTempControlContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'Temperature Control class module
'
'TODO: ?
'   Add temperature mode state (HeatOnly, CoolOnly, HeatOrCool) so we can limit
'   temperature states accordingly.

Option Explicit
'use parameters from other tempcontrol

Public Enum ETempModeContacts
  ModeHeatAndCool
  ModeHeatOnly
  ModeCoolOnly
  ModeChangeDisabled
End Enum
Public Mode As ETempModeContacts

Public Enum ETempStateContacts
  TempOff
  TempStart
  TempPause
  TempPreHeatVent
  TempHeat
  TempPostHeatVent
  TempPreCoolVent
  TempCool
  TempPostCoolVent
End Enum

Public State As ETempStateContacts
Private PreviousState As ETempStateContacts
Public PID As New acPIDcontrolContacts
Private IdleTimer As New acTimer
Public StateTimer As New acTimer
Private EnableTimer As New acTimer
Private HeatDelayTimer As New acTimer
Private CoolDelayTimer As New acTimer
Private ModeChangeTimer As New acTimer

'PID setpoints - these are useful for restarting temp control after crash cooling
Private pidStartTemp As Long
Private pidFinalTemp As Long
Private pidGradient As Long
Public TempLoAlarmTimer As New acTimer
Public TempHiAlarmTimer As New acTimer
Public IgnoreErrors As Boolean
Public TemperatureLow As Boolean
Public TemperatureHigh As Boolean

'Temperature control parameters
'Temp enabled must be made for this time before we can heat or cool
'Can be zero (default = 10)
Private TempEnableDelay As Long
'If we we're cooling and now we want to heat - delay heating for this time (and vice versa)
'This can be used to ensure the heat exchanger drain valve has been open long enough to
'fully drain the heat exchanger prior to heating or cooling.
'Can be zero (default = 10)
Private TempHeatCoolDelay As Long
'Usual mode change delay - can be zero (default = 120)
Private HeatCoolModeChangeDelay As Long
'Time to vent prior to heating - can be zero (default = 10)
Private TempPreHeatVentTime As Long
'Time to vent after heating - can be zero (default = 10)
Private TempPostHeatVentTime As Long
'Time to vent prior to cooling - can be zero (default = 10)
Private TempPreCoolVentTime As Long
'Time to vent after cooling - can be zero (default = 10)
Private TempPostCoolVentTime As Long

Private FirstTempStart As Boolean
  
Private CoolPropBand As Long
Private CoolIntegral As Long
Private CoolMaxGradient As Long
Private CoolStepMargin As Long


Private Sub Class_Initialize()

'Default values for temperature control
  Mode = ModeHeatAndCool
  State = TempOff
  StateTimer = 10
  EnableTimer = 10

'Set parameters to defaults
  TempEnableDelay = 10
  TempHeatCoolDelay = 30
  HeatCoolModeChangeDelay = 120
  TempPreHeatVentTime = 10
  TempPostHeatVentTime = 10
  TempPreCoolVentTime = 10
  TempPostCoolVentTime = 10
  
  FirstTempStart = True
  
End Sub

Public Sub Start(ByVal VesTemp As Long, _
                 ByVal FinalTempInTenths As Long, _
                 ByVal GradientInTenthsPerMinute As Long, ByVal ControlObject As Object)
 Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
 With ControlCode
'Set start temperature, target temperature and gradient
  pidStartTemp = VesTemp
  pidFinalTemp = FinalTempInTenths
  pidGradient = GradientInTenthsPerMinute
  
'Set pid parameters to heat by default
  PID.PropBand = .TemperatureControl.Parameters.HeatPropBand
  PID.Integral = .TemperatureControl.Parameters.HeatIntegral
  PID.MaxGradient = .TemperatureControl.Parameters.HeatMaxGradient
  PID.HoldMargin = .TemperatureControl.Parameters.HeatStepMargin
  
  FirstTempStart = True

'Start the PID
  PID.Start pidStartTemp, pidFinalTemp, pidGradient

'Decide wether we should heat, cool or wait and see
'Decision is deliberately biased in favour of heating - cos that's more likely

'If we're already heating (or about to heat) and the current temp is lower than
'the Final Temp (and a bit) then keep going i.e. don't change state
  If IsHeating Or IsPreHeatVent Then
    If VesTemp <= (FinalTempInTenths + .TemperatureControl.Parameters.HeatPropBand) Then Exit Sub
  End If

'If we're already cooling (or about to cool) and the current temp is higher than
'the Final Temp (by a bit) then keep going i.e. don't change state
  If IsCooling Or IsPreCoolVent Then
    If VesTemp > (FinalTempInTenths + .TemperatureControl.Parameters.CoolPropBand) Then Exit Sub
  End If
  
'If we're venting after heat or cool then let vent finish and allow start state to
'decide which way to go
  If IsPostHeatVent Or IsPostCoolVent Then Exit Sub
  
'Set state to start
  State = TempStart
  StateTimer = 5
End With
End Sub

Public Sub Run(ByVal VesTemp As Long, ByVal GradientInTenthsPerMinute As Long, ByVal ControlObject As Object)
 Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
 With ControlCode



  Select Case State
  
    Case TempOff
      'Set previous state variable
      PreviousState = TempStart
      'Set state timer
      StateTimer = 5
  
    Case TempStart
      'Set previous state variable
      PreviousState = TempStart
      'If no start temp or final temp then set state to off
      If pidStartTemp = 0 Or pidFinalTemp = 0 Then State = TempOff
      'Wait for state timer
      If StateTimer Then Exit Sub
      'Wait for temperature enable
      If EnableTimer Then Exit Sub
      'Reset PID to clear out any "funnies"
      'Comment this out to see if it cures cool on gradient bug
      'PID.Reset VesTemp
      'Run PID to decide whether to heat or cool
      PID.Run VesTemp, GradientInTenthsPerMinute
      'Added to allow us to go straight to cooling on a controlled gradient
      'If we need to cool, don't think about heating yet
      If FirstTempStart Then
        FirstTempStart = False
        If VesTemp > (pidFinalTemp + .TemperatureControl.Parameters.CoolPropBand) Then
          State = TempPreCoolVent
          StateTimer = TempPreCoolVentTime
          Exit Sub
        End If
      End If
      'If PID Output is greater than zero then heat
      If PID.Output > 0 Then
        'Are we allowed to heat ?
        If Mode = ModeCoolOnly Then Exit Sub
        'Have we waited long enough before switching to heating
        If HeatDelayTimer Then Exit Sub
        State = TempPreHeatVent
        StateTimer = TempPreHeatVentTime
      End If
      'If PID Output is less than zero then cool
      If (PID.Output < 0) Then
        'Are we allowed to cool ?
        If Mode = ModeHeatOnly Then Exit Sub
        'Have we waited long enough before switching to cooling
        If CoolDelayTimer Then Exit Sub
        State = TempPreCoolVent
        StateTimer = TempPreCoolVentTime
      End If
    
    Case TempPause
      'Pause PID
      PID.Pause
      'Wait for state timer
      If StateTimer Then Exit Sub
      'Wait for temperature enable
      If EnableTimer Then Exit Sub
      'Switch back to previous state
      State = PreviousState
      'Restart PID
      PID.Restart
      'If venting set timer to parameter value
      If State = TempPreHeatVent Then StateTimer = TempPreHeatVentTime
      If State = TempPostHeatVent Then StateTimer = TempPostHeatVentTime
      If State = TempPreCoolVent Then StateTimer = TempPreCoolVentTime
      If State = TempPostCoolVent Then StateTimer = TempPostCoolVentTime
      
'Heating

    Case TempPreHeatVent
      'Set previous state variable
      PreviousState = TempPreHeatVent
      'Reset cool delay timer
      CoolDelayTimer = TempHeatCoolDelay
      'If enable timer set switch to pause state
      If EnableTimer Then State = TempPause
      'Wait for parameter time then switch to heating
      If StateTimer Then Exit Sub
      State = TempHeat
      'Reset mode change timer
      ModeChangeTimer = HeatCoolModeChangeDelay
      'Reset PID to start from current temp
      PID.Reset VesTemp, GradientInTenthsPerMinute
      
    Case TempHeat
      'Set previous state variable
      PreviousState = TempHeat
      'Reset cool delay timer
      CoolDelayTimer = TempHeatCoolDelay
      'If enable timer set switch to pause state
      If EnableTimer Then State = TempPause
      'Set PID parameters here so that changes apply immediately
      PID.PropBand = .TemperatureControl.Parameters.HeatPropBand
      PID.Integral = .TemperatureControl.Parameters.HeatIntegral
      PID.MaxGradient = .TemperatureControl.Parameters.HeatMaxGradient
      PID.HoldMargin = .TemperatureControl.Parameters.HeatStepMargin
      'Run PID Control
      PID.Run VesTemp, GradientInTenthsPerMinute
      'Reset mode change timer if mode change disabled (Note parameter could be zero)
      If Mode = ModeChangeDisabled Then ModeChangeTimer = HeatCoolModeChangeDelay + 1
      'Reset mode change timer while we're still calling for heating
      If PID.Output >= 0 Then ModeChangeTimer = HeatCoolModeChangeDelay
      If ModeChangeTimer.Finished And ((Not IsHolding) Or (VesTemp > pidFinalTemp + .TemperatureControl.Parameters.HeatExitDeadband)) Then
        State = TempPostHeatVent
        StateTimer = TempPostHeatVentTime
      End If
          
    Case TempPostHeatVent
      'Set previous state variable
      PreviousState = TempPostHeatVent
      'Reset cool delay timer
      CoolDelayTimer = TempHeatCoolDelay
      'If enable timer set switch to pause state
      If EnableTimer Then State = TempPause
      'Wait for parameter time then switch to idle state
      If StateTimer Then Exit Sub
      State = TempStart
      StateTimer = 5
    
'Cooling
    
    Case TempPreCoolVent
      'Set previous state variable
      PreviousState = TempPreCoolVent
      'Reset heat delay timer
      HeatDelayTimer = TempHeatCoolDelay
      'If enable timer set switch to pause state
      If EnableTimer Then State = TempPause
      'Wait for parameter time then switch to cooling
      If StateTimer Then Exit Sub
      State = TempCool
      'Reset mode change timer
      ModeChangeTimer = HeatCoolModeChangeDelay
      'Reset PID to start from current temp
      PID.Reset VesTemp, GradientInTenthsPerMinute
      
    Case TempCool
      'Set previous state variable
      PreviousState = TempCool
      'Reset heat delay timer
      HeatDelayTimer = TempHeatCoolDelay
      'If enable timer set switch to pause state
      If EnableTimer Then State = TempPause
      'Set PID parameters here so that changes apply immediately
      PID.PropBand = CoolPropBand
      PID.Integral = CoolIntegral
      PID.MaxGradient = CoolMaxGradient
      PID.HoldMargin = CoolStepMargin
      'Run PID Control
      PID.Run VesTemp, GradientInTenthsPerMinute
      'Reset mode change timer if mode change disabled (Note parameter could be zero)
      If Mode = ModeChangeDisabled Then ModeChangeTimer = HeatCoolModeChangeDelay + 1
      'Reset mode change timer while we're still calling for cooling
      If PID.Output <= 0 Then ModeChangeTimer = HeatCoolModeChangeDelay
      If ModeChangeTimer.Finished Then
        State = TempPostCoolVent
        StateTimer = TempPostCoolVentTime
      End If
    
    Case TempPostCoolVent
      'Set previous state variable
      PreviousState = TempPostCoolVent
      'Reset heat delay timer
      HeatDelayTimer = TempHeatCoolDelay
      'If enable timer set switch to pause state
      If EnableTimer Then State = TempPause
      'Wait for parameter time then switch to idle state
      If StateTimer Then Exit Sub
      'Switch to cooling state
      State = TempStart
      StateTimer = 5
      
  End Select
  End With
End Sub
Public Sub CheckErrorsAndMakeAlarms(ByVal VesTemp As Long, ByVal GradientInTenthsPerMinute As Long, ByVal ControlObject As Object)
 Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
 With ControlCode
'Check pid pause and reset
  Dim TempError As Long
  TempError = 0
  If IsHeating Then TempError = TempSetpoint - VesTemp
  If IsCooling Then TempError = VesTemp - TempSetpoint
  IgnoreErrors = IsMaxGradient And .TC.IsRamping
  If Not IgnoreErrors Then
    If TempError > .TemperatureControl.Parameters.TemperaturePidReset Then pidReset VesTemp, GradientInTenthsPerMinute
    If TempError > .TemperatureControl.Parameters.TemperaturePidPause Then pidPause
    If IsPidPaused Then
      If TempError < .TemperatureControl.Parameters.TemperaturePidRestart Then pidRestart
    End If
  End If
  MakeAlarms VesTemp, ControlCode
  End With
End Sub
' Make alarms for temperature control
Private Sub MakeAlarms(ByVal VesTemp As Long, ByVal ControlObject As Object)
 Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
 With ControlCode
  'Temperature low/high s
  If IgnoreErrors Or (Not (IsHeating Or IsCooling)) Then
    TempLoAlarmTimer = .TemperatureControl.Parameters.TemperatureAlarmDelay
    TempHiAlarmTimer = .TemperatureControl.Parameters.TemperatureAlarmDelay
  End If
  If (VesTemp > (TempSetpoint - .TemperatureControl.Parameters.TemperatureAlarmBand)) Then
    TempLoAlarmTimer = .TemperatureControl.Parameters.TemperatureAlarmDelay
  End If
  If (VesTemp < (TempSetpoint + .TemperatureControl.Parameters.TemperatureAlarmBand)) Then
    TempHiAlarmTimer = .TemperatureControl.Parameters.TemperatureAlarmDelay
  End If
  If TempLoAlarmTimer.Finished Then
    TemperatureLow = True
  Else
     TemperatureLow = False
  End If
  If TempHiAlarmTimer.Finished Then
    TemperatureHigh = True
  Else
    TemperatureHigh = False
  End If
  End With
End Sub
Public Sub Cancel()

  PID.Cancel
  State = TempOff
  pidStartTemp = 0
  pidFinalTemp = 0
  pidGradient = 0

End Sub
Public Sub pidPause()

  PID.Pause

End Sub
' TODO: don't really need this return value
Public Function pidRestart() As Boolean
  PID.Restart
  pidRestart = True
End Function
Public Sub pidReset(VesTemp As Long, GradientInTenthsPerMinute As Long)
  PID.Reset VesTemp, GradientInTenthsPerMinute
End Sub
Public Sub resetEnableTimer()
  EnableTimer = TempEnableDelay
End Sub
Public Property Let EnableDelay(ByVal Value As Long)
  TempEnableDelay = Value
  If TempEnableDelay < 0 Then TempEnableDelay = 0
  If TempEnableDelay > 30 Then TempEnableDelay = 30
End Property
Public Property Let TempMode(ByVal Value As Long)

  Mode = ModeHeatAndCool
  If Value = 2 Then Mode = ModeChangeDisabled
  If Value = 3 Then Mode = ModeHeatOnly
  If Value = 4 Then Mode = ModeCoolOnly

End Property
Public Property Let CoolingPropBand(ByVal Value As Long)
  CoolPropBand = Value
End Property
Public Property Let CoolingIntegral(ByVal Value As Long)
  CoolIntegral = Value
End Property
Public Property Let CoolingMaxGradient(ByVal Value As Long)
  CoolMaxGradient = Value
End Property
Public Property Let CoolingStepMargin(ByVal Value As Long)
  CoolStepMargin = Value
End Property
Public Property Let PreHeatVentTime(ByVal Value As Long)
  TempPreHeatVentTime = Value
  If TempPreHeatVentTime < 0 Then TempPreHeatVentTime = 0
  If TempPreHeatVentTime > 60 Then TempPreHeatVentTime = 60
End Property
Public Property Let PostHeatVentTime(ByVal Value As Long)
  TempPostHeatVentTime = Value
  If TempPostHeatVentTime < 0 Then TempPostHeatVentTime = 0
  If TempPostHeatVentTime > 60 Then TempPostHeatVentTime = 60
End Property
Public Property Let PreCoolVentTime(ByVal Value As Long)
  TempPreCoolVentTime = Value
  If TempPreCoolVentTime < 0 Then TempPreCoolVentTime = 0
  If TempPreCoolVentTime > 60 Then TempPreCoolVentTime = 60
End Property
Public Property Let PostCoolVentTime(ByVal Value As Long)
  TempPostCoolVentTime = Value
  If TempPostCoolVentTime < 0 Then TempPostCoolVentTime = 0
  If TempPostCoolVentTime > 60 Then TempPostCoolVentTime = 60
End Property
'
'Properties - NOTE: properties return false by default

Public Property Get IsEnabled() As Boolean

  IsEnabled = EnableTimer.Finished

End Property

Public Property Get IsIdle() As Boolean

  If (State = TempStart) Or (State = TempOff) Then IsIdle = True

End Property

Public Property Get IsPaused() As Boolean

  If (State = TempPause) Then IsPaused = True

End Property

Public Property Get IsHeating() As Boolean

  If (State = TempHeat) Then IsHeating = True
  
End Property

Public Property Get IsCooling() As Boolean

  If (State = TempCool) Then IsCooling = True
  
End Property

Public Property Get IsMaxGradient() As Boolean

'Set max gradient only if we are ramping up/down (it's used to disable alarms)
  If PID.IsMaxGradient Then IsMaxGradient = True

End Property

Public Property Get IsHolding() As Boolean

'Make sure IsHolding returns false during Crashcooling
  If PID.IsHolding Then IsHolding = True

End Property

Public Property Get IsPreHeatVent() As Boolean

  If (State = TempPreHeatVent) Then IsPreHeatVent = True
  
End Property

Public Property Get IsPostHeatVent() As Boolean

  If (State = TempPostHeatVent) Then IsPostHeatVent = True

End Property

Public Property Get IsPreCoolVent() As Boolean

  If (State = TempPreCoolVent) Then IsPreCoolVent = True
  
End Property

Public Property Get IsPostCoolVent() As Boolean

  IsPostCoolVent = (State = TempPostCoolVent)

End Property

Public Property Get Output() As Long

'Analog Output
  Output = 0
  If (State = TempHeat) Then Output = PID.Output
  If (State = TempCool) Then Output = -PID.Output

'Limit output
  If Output < 0 Then Output = 0
  If Output > 1000 Then Output = 1000

End Property

Public Property Get IsPidPaused() As Boolean

  If PID.IsPaused Then IsPidPaused = True

End Property
'
'Added for display purposes

Public Property Get TempGradient() As Long

  TempGradient = PID.Gradient

End Property

Public Property Get TempFinalTemp() As Long

  TempFinalTemp = PID.FinalTemp

End Property

Public Property Get TempSetpoint() As Long

  TempSetpoint = PID.Setpoint

End Property

Public Property Get pidPropTerm() As Long

  pidPropTerm = PID.PropTerm

End Property

Public Property Get pidIntegralTerm() As Long

  pidIntegralTerm = PID.IntegralTerm

End Property

Public Property Get pidHeatLossTerm() As Long

  pidHeatLossTerm = PID.HeatLossTerm

End Property

Public Property Get pidGradientTerm() As Long

  pidGradientTerm = PID.GradientTerm

End Property

#End If
End Class
