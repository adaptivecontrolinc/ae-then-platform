Public Class PidContactsNeed
  ' TODO - Do we need this at Mt. Holly

#If 0 Then
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acPIDcontrolContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'PID Control - NOTE: Max Gradient = 0 rather than 99

Option Explicit

Public Enum pidstate                               ' Are we heating, cooling or holding temp
  pidRampUp
  pidRampDown
  pidHold
End Enum
Public State As pidstate

Private pidStartTemp As Long                ' Start temperature for PID
Private pidFinalTemp As Long                ' Final temperature for PID
Private pidGradient As Long                 ' Gradient in degrees per minute
Private pidGradientTimer As New acTimerUp   ' Timer for gradient control
Private pidSetpoint As Long                 ' Calculated setpoint
Private pidOutput As Long                   ' pid output in tenths of percent

Private pidPropTerm As Long                 ' Declared as global so we can see from
Private pidIntegralTerm As Long             ' the outside
Private pidHeatLossTerm As Long             '
Private pidGradientTerm As Long             '

Private pidErrorSum As Long                 ' Declared as global so we can easily reset
Private pidErrorSumCounter As Long          '     "
Private pidErrorSumTimer As New acTimer     '     "

Private pidPropBand As Long                 ' PID control parameters
Private pidIntegral As Long                 '     "
Private pidMaxGradient As Long              '     "
Private pidBalanceTemp As Long              '     "
Private pidBalPercent As Long            '     "
Private pidHoldMargin As Long               ' Go to hold if within this margin of setpoint

Private pidPaused As Boolean                ' PID paused ?
Private pidCancelled As Boolean             ' PID cancelled ?
  
  

Private Sub Class_Initialize()

  pidPropBand = 40              ' Default PID parameters (assumes Farenheit)
  pidIntegral = 250             '     "
  pidMaxGradient = 50           '     "
  pidBalanceTemp = 800          '     "
  pidBalPercent = 100        '     "
  pidHoldMargin = 20            '     "
  
  pidCancelled = True           '

End Sub

Public Sub Start(ByVal StartTempInTenths As Long, ByVal FinalTempInTenths As Long, _
                 ByVal GradientInTenthsPerMinute As Long)
                      
'Set parameters
  pidStartTemp = StartTempInTenths
  pidFinalTemp = FinalTempInTenths
  pidGradient = GradientInTenthsPerMinute
  pidGradientTimer.Start
  
'Reset control
  pidPaused = False
  pidCancelled = False
  ResetIntegral

'Are we heating or cooling ?
  State = pidRampUp
  If pidFinalTemp < pidStartTemp Then State = pidRampDown

End Sub

Public Sub Run(ByVal CurrentTempInTenths As Long, ByVal CurrentGradientInTenthsPerMinute As Long)
  
'Check to see if PID cancelled
  If pidCancelled Then Exit Sub
  pidGradient = CurrentGradientInTenthsPerMinute

'Calculate Setpoint
  Dim RampFinished As Boolean
  If (State = pidHold) Then RampFinished = True
  If (State = pidRampUp) And (pidSetpoint >= pidFinalTemp) Then RampFinished = True
  If (State = pidRampDown) And (pidSetpoint <= pidFinalTemp) Then RampFinished = True

  If ((CurrentTempInTenths > (pidFinalTemp - pidHoldMargin)) And _
     (CurrentTempInTenths < (pidFinalTemp + pidHoldMargin))) Or IsMaxGradient Then
     State = pidHold
     pidSetpoint = pidFinalTemp
     RampFinished = True
  End If

  
  If Not (IsMaxGradient Or RampFinished) Then
    If (State = pidRampUp) Then pidSetpoint = pidStartTemp + ((pidGradientTimer * pidGradient) / 60)
    If (State = pidRampDown) Then pidSetpoint = pidStartTemp - ((pidGradientTimer * pidGradient) / 60)
  Else
    pidSetpoint = pidFinalTemp
    If IsRampUp And (CurrentTempInTenths > (pidFinalTemp - pidHoldMargin)) Then
      State = pidHold
    End If
    If IsRampDown And (CurrentTempInTenths < (pidFinalTemp + pidHoldMargin)) Then
      State = pidHold
    End If
  End If

'Calculate error
  Dim TempError As Long
  TempError = pidSetpoint - CurrentTempInTenths

'Calculate proportional Term
  pidPropTerm = (TempError * 1000) / pidPropBand
  
'If PID output is maxxed out stop Integral action.
'This should prevent Integral saturation i.e. Error Sum/Integral term getting huge!
  Dim StopIntegral As Boolean
  If (pidOutput = 1000) And (TempError > 0) Then StopIntegral = True
  If (pidOutput = -1000) And (TempError < 0) Then StopIntegral = True
'Calculate Error Sum for integral term - add once a second if allowed
  If (Not StopIntegral) And pidErrorSumTimer.Finished Then
      pidErrorSumTimer = 1
      pidErrorSum = pidErrorSum + TempError
  End If

'Calculate Integral Term - limit to +/- 100%
  pidIntegralTerm = (pidErrorSum * pidIntegral) / 6000
  If pidIntegralTerm > 1000 Then pidIntegralTerm = 1000
  If pidIntegralTerm < -1000 Then pidIntegralTerm = -1000

'Calculate heat loss term - limit to +/- 100%
  pidHeatLossTerm = ((CurrentTempInTenths - pidBalanceTemp) * pidBalPercent) / 1000
  If pidHeatLossTerm > 1000 Then pidHeatLossTerm = 1000
  If pidHeatLossTerm < -1000 Then pidHeatLossTerm = -1000

'Calculate gradient term - limit to +/- 100%

  pidGradientTerm = (1000 * pidGradient) / pidMaxGradient
  If (State = pidRampDown) Then pidGradientTerm = (0 - pidGradientTerm)
  If pidGradientTerm > 1000 Then pidGradientTerm = 1000
  If pidGradientTerm < -1000 Then pidGradientTerm = -1000
  If IsHolding Then pidGradientTerm = 0

'Calculate Output  - limit to +/- 100%
  pidOutput = pidPropTerm + pidIntegralTerm + pidHeatLossTerm + pidGradientTerm
  If pidOutput > 1000 Then pidOutput = 1000
  If pidOutput < -1000 Then pidOutput = -1000
  
End Sub

Public Sub Pause()

  pidPaused = True
  pidGradientTimer.Pause

End Sub

Public Sub Restart()

  pidPaused = False
  pidGradientTimer.Restart
  
End Sub

Public Sub Reset(ByVal CurrentTempInTenths As Long, ByVal GradientInTenthsPerMinute As Long)

  pidPaused = False
  pidStartTemp = CurrentTempInTenths
  pidGradientTimer.Start
  ResetIntegral
  pidGradient = GradientInTenthsPerMinute


End Sub

Public Sub Cancel()

  pidCancelled = True
  ResetIntegral
  pidStartTemp = 0
  pidFinalTemp = 0
  pidGradient = 0
  pidSetpoint = 0
  pidOutput = 0

End Sub

Private Sub ResetIntegral()

'Reset integral
  pidErrorSum = 0
  pidErrorSumTimer = 1
  pidErrorSumCounter = 0

End Sub

'
'Properties

Public Property Let PropBand(Value As Long)

  pidPropBand = Value
  If pidPropBand < 0 Then pidPropBand = 0
  If pidPropBand > 1000 Then pidPropBand = 1000

End Property

Public Property Let Integral(Value As Long)

  pidIntegral = Value
  If pidIntegral < 0 Then pidIntegral = 0
  If pidIntegral > 1000 Then pidIntegral = 1000
  
End Property

Public Property Let MaxGradient(Value As Long)

  pidMaxGradient = Value
  If pidMaxGradient < 0 Then pidMaxGradient = 0
  If pidMaxGradient > 1000 Then pidMaxGradient = 1000

End Property
Public Property Get BalanceTemp() As Long
  BalanceTemp = pidBalanceTemp
End Property
Public Property Let BalanceTemp(Value As Long)
  pidBalanceTemp = Value
  Minimum pidBalanceTemp, 0:  Maximum pidBalanceTemp, 1000
End Property
Public Property Get BalPercent() As Long
  BalPercent = pidBalPercent
End Property
Public Property Let BalPercent(Value As Long)
  pidBalPercent = Value
  Minimum pidBalPercent, 0:  Maximum pidBalPercent, 1000
End Property

Public Property Let HoldMargin(Value As Long)

  pidHoldMargin = Value
  If pidHoldMargin < 10 Then pidHoldMargin = 10
  If pidHoldMargin > 100 Then pidHoldMargin = 100

End Property

'
'Properties - NOTE: all properties return false by default

Public Property Get Output() As Long

  Output = pidOutput
  If pidCancelled Then Output = 0

End Property

Public Property Get Gradient() As Long

  Gradient = pidGradient

End Property

Public Property Get FinalTemp() As Long

  FinalTemp = pidFinalTemp

End Property

Public Property Get Setpoint() As Long

  Setpoint = pidSetpoint

End Property

Public Property Get IsRampUp() As Boolean

  If State = pidRampUp Then IsRampUp = True

End Property

Public Property Get IsRampDown() As Boolean

  IsRampDown = (State = pidRampDown)
  
End Property

Public Property Get IsHolding() As Boolean

  If State = pidHold Then IsHolding = True
  
End Property

Public Property Get IsRamping() As Boolean

  If (Not (IsMaxGradient Or IsHolding)) Then IsRamping = True

End Property

Public Property Get IsMaxGradient() As Boolean

  If pidGradient = 0 Then IsMaxGradient = True

End Property

Public Property Get IsPaused() As Boolean

  If pidPaused Then IsPaused = True

End Property

'
'Added for debug purposes - delete (eventually)

Public Property Get PropTerm() As Long

  PropTerm = pidPropTerm

End Property

Public Property Get IntegralTerm() As Long

  IntegralTerm = pidIntegralTerm

End Property

Public Property Get HeatLossTerm() As Long

  HeatLossTerm = pidHeatLossTerm

End Property

Public Property Get GradientTerm() As Long

  GradientTerm = pidGradientTerm

End Property

Public Property Get Propband2() As Long
  Propband2 = pidPropBand
End Property




#End If
End Class
