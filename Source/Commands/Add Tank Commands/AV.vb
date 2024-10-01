Public Class AV










  ' AV.vb
#If 0 Then
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Add dosing command
'
' - Stop timer if pump stops

Option Explicit
Implements ACCommand
Public Enum AVState
  AVOff
  AVWaitReady
  AVPause
  AVDose
  AVDosePause
  AVTransfer
  AVTransferEmpty1
  AVRinse
  AVTransferEmpty2
  AVRinseToDrain
  AVTransferToDrain
End Enum
Public State As AVState
Public Timer As New acTimer
Public AddTime As Long
Public AddCurve As Long
Public StartLevel As Long
Public DesiredLevel As Long
Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Parameters=|0-99|m Curve:|0-9|\r\nName=Add Dose\r\nMinutes=(1+'1)\r\nHelp=Doses the contents of the side tank over the time specified, using one of ten curves. Curve 0 is linear, odd numbers are progressive adds and even numbers are regressive adds "
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  
  AddTime = Param(1) * 60
  AddCurve = Param(2)
  With ControlCode
    .DR.ACCommand_Cancel: .FI.ACCommand_Cancel: .HD.ACCommand_Cancel
    .HC.ACCommand_Cancel: .LD.ACCommand_Cancel: .PH.ACCommand_Cancel
    .RC.ACCommand_Cancel: .RH.ACCommand_Cancel: .RI.ACCommand_Cancel
    .RT.ACCommand_Cancel: .SA.ACCommand_Cancel: .TM.ACCommand_Cancel
    .UL.ACCommand_Cancel: .WK.ACCommand_Cancel: .WT.ACCommand_Cancel
  End With
  
  
  'For adds under pressure use the next line
  State = AVWaitReady

End Sub
Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  
  With ControlCode
    Select Case State
      
      Case AVWaitReady
        If .AP.IsOn Then
          If (Not .AddReady) Or (Not .AP.IsReady) Then
             Timer = 10
          End If
        Else
          If Not .AddReady Then Timer = 10
        End If
        If Timer.Finished Then
           If AddTime = 0 Then
              .AP.AddMixing = False
              State = AVTransferEmpty1
              Timer = .Parameters.AddTransferTimeBeforeRinse
           Else
              .AP.AddMixing = False
              State = AVDosePause
              Timer = AddTime
              StartLevel = .AdditionLevel
           End If
        End If
        
      Case AVPause
        Timer.Pause
        If (Not .Parent.IsPaused) And (Not .IO_EStop_PB) And .IO_MainPumpRunning Then
          State = AVDosePause
          Timer.Restart
        End If
        
      Case AVDose
        'Check level - if level low switch to circulate
        If .AdditionLevel < Setpoint() Then State = AVDosePause
        'Check level - if level high then switch to transfer
        If .AdditionLevel > (Setpoint() + 30) Then State = AVTransfer
        'If pump not running go to pause state
        If .Parent.IsPaused Or .IO_EStop_PB Or Not .IO_MainPumpRunning Then State = AVPause
        'If we're finished transfer empty
        If Timer.Finished Then
          State = AVTransferEmpty1
          Timer = .Parameters.AddTransferTimeBeforeRinse
        End If
      
      Case AVTransfer
        'Check level - if level low switch to circulate
        If .AdditionLevel < Setpoint() Then State = AVDosePause
        'If pump not running go to pause state
        If Not .IO_MainPumpRunning Then State = AVPause
        'If we're finished transfer empty
        If Timer.Finished Then
          State = AVTransferEmpty1
          Timer = .Parameters.AddTransferTimeBeforeRinse
        End If


      Case AVDosePause
        'Check level - if high switch to dose
        If .AdditionLevel > Setpoint() Then State = AVDose
        'If pump not running go to pause state
        If Not .IO_MainPumpRunning Then State = AVPause
        'If we're finished transfer empty
        If Timer.Finished Then
          State = AVTransferEmpty1
          Timer = .Parameters.AddTransferTimeBeforeRinse
        End If
      
      Case AVTransferEmpty1
        If .AdditionLevel > 0 Then Timer = .Parameters.AddTransferTimeBeforeRinse
        If Timer.Finished Then
          State = AVRinse
          Timer = .Parameters.AddTransferRinseTime
        End If
        
      Case AVRinse
        If Timer.Finished Then
          State = AVTransferEmpty2
          Timer = .Parameters.AddTransferTimeAfterRinse
        End If
         
      Case AVTransferEmpty2
        If .AdditionLevel > 0 Then Timer = .Parameters.AddTransferTimeAfterRinse
        If Timer.Finished Then
          .AddReady = False
          .AP.ACCommand_Cancel
          State = AVRinseToDrain
          Timer = .Parameters.AddTransferRinseToDrainTime
        End If
      
      Case AVRinseToDrain
        If Timer.Finished Then
          State = AVTransferToDrain
          Timer = .Parameters.AddTransferToDrainTime
        End If
        
      Case AVTransferToDrain
        If .AdditionLevel > 0 Then Timer = .Parameters.AddTransferToDrainTime
        If Timer.Finished Then
          .AddReady = False
          .AP.ACCommand_Cancel
          State = AVOff
        End If
    End Select
  End With

End Sub
Friend Sub ACCommand_Cancel()
  State = AVOff
End Sub
Friend Property Get IsOn() As Boolean
  IsOn = (State <> AVOff)
End Property
Friend Property Get IsActive() As Boolean
  If IsTransfer Or (State = AVDosePause) Or IsRinse Then IsActive = True
End Property
Friend Property Get IsWaitReady() As Boolean
  If (State = AVWaitReady) Then IsWaitReady = True
End Property
Friend Property Get IsTransfer() As Boolean
  If (State = AVDose) Or (State = AVTransfer) Or _
     (State = AVTransferEmpty1) Or (State = AVTransferEmpty2) Then IsTransfer = True
End Property
Friend Property Get IsDosing() As Boolean
  If (State = AVDose) Then IsDosing = True
End Property
Friend Property Get IsDosePaused() As Boolean
  If (State = AVDosePause) Then IsDosePaused = True
End Property
Friend Property Get IsRinse() As Boolean
  If (State = AVRinse) Then IsRinse = True
End Property
Friend Property Get IsRinseToDrain() As Boolean
  If (State = AVRinseToDrain) Then IsRinseToDrain = True
End Property
Friend Property Get IsTransferToDrain() As Boolean
  If (State = AVTransferToDrain) Then IsTransferToDrain = True
End Property
Friend Property Get IsDelayed() As Boolean
  If (State = AVWaitReady) Then IsDelayed = True
End Property
'Function uses the expressions (y = x * x) and (y = sqr(x)) to generate curves for
'dosing control.
'x = A fraction between 0 and 1 (I know) representing elapsed time
'y = A fraction between 0 and 1 representing the amount we should have transferred so far
'
'The curves are scaled by adding a percent of the difference between the curve value and
'the equivalent linear value to the original curve value (??).
Private Function Setpoint() As Long
'If timer has finished exit function
  If Timer.Finished Then
    Setpoint = 0
    Exit Function
  End If
  
'Amount we should have transferred so far
  Dim ElapsedTime As Double, LinearTerm As Double, TransferAmount As Double
  ElapsedTime = (AddTime - Timer) / AddTime
  LinearTerm = ElapsedTime
  TransferAmount = StartLevel * LinearTerm
  
'Calculate scaling factor (0-1) for progressive and digressive curves
  If AddCurve > 0 Then
    Dim ScalingFactor As Double
    ScalingFactor = (10 - AddCurve) / 10
'Calculate term for progressive transfer (0-1) if odd curve
    If (AddCurve Mod 2) = 1 Then
      Dim MaxOddCurve As Double, OddTerm As Double
      MaxOddCurve = (ElapsedTime * ElapsedTime)
      OddTerm = MaxOddCurve + ((LinearTerm - MaxOddCurve) * ScalingFactor)
      TransferAmount = StartLevel * OddTerm
    Else
'Calculate term for digressive transfer (0-1) if even curve
      Dim MaxEvenCurve As Double, EvenTerm As Double
      MaxEvenCurve = Sqr(ElapsedTime)
      EvenTerm = MaxEvenCurve - ((MaxEvenCurve - LinearTerm) * ScalingFactor)
      TransferAmount = StartLevel * EvenTerm
    End If
  End If
   
'Calculate setpoint and limit to 0-1000
  Setpoint = StartLevel - CLng(TransferAmount)
  If Setpoint < 0 Then Setpoint = 0
  If Setpoint > 1000 Then Setpoint = 1000
  
'Global variable for display purposes - yuk!!
  DesiredLevel = Setpoint
End Function
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property

#End If


  ' TODO 
#If 0 Then
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Add dosing command
'
' - Stop timer if pump stops

Option Explicit
Implements ACCommand
Public Enum AVState
  AVOff
  AVWaitReady
  AVPause
  AVDose
  AVDosePause
  AVTransfer
  AVTransferEmpty1
  AVRinse
  AVTransferEmpty2
  AVRinseToDrain
  AVTransferToDrain
End Enum
Public State As AVState
Public Timer As New acTimer
Public AddTime As Long
Public AddCurve As Long
Public StartLevel As Long
Public DesiredLevel As Long
Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Parameters=|0-99|m Curve:|0-9|\r\nName=Add Dose\r\nMinutes=(1+'1)\r\nHelp=Doses the contents of the side tank over the time specified, using one of ten curves. Curve 0 is linear, odd numbers are progressive adds and even numbers are regressive adds "
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  
  AddTime = Param(1) * 60
  AddCurve = Param(2)
  With ControlCode
    .DR.ACCommand_Cancel: .FI.ACCommand_Cancel: .HD.ACCommand_Cancel
    .HC.ACCommand_Cancel: .LD.ACCommand_Cancel: .PH.ACCommand_Cancel
    .RC.ACCommand_Cancel: .RH.ACCommand_Cancel: .RI.ACCommand_Cancel
    .RT.ACCommand_Cancel: .SA.ACCommand_Cancel: .TM.ACCommand_Cancel
    .UL.ACCommand_Cancel: .WK.ACCommand_Cancel: .WT.ACCommand_Cancel
  End With
  
  
  'For adds under pressure use the next line
  State = AVWaitReady

End Sub
Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  
  With ControlCode
    Select Case State
      
      Case AVWaitReady
        If .AP.IsOn Then
          If (Not .AddReady) Or (Not .AP.IsReady) Then
             Timer = 10
          End If
        Else
          If Not .AddReady Then Timer = 10
        End If
        If Timer.Finished Then
           If AddTime = 0 Then
              .AP.AddMixing = False
              State = AVTransferEmpty1
              Timer = .Parameters.AddTransferTimeBeforeRinse
           Else
              .AP.AddMixing = False
              State = AVDosePause
              Timer = AddTime
              StartLevel = .AdditionLevel
           End If
        End If
        
      Case AVPause
        Timer.Pause
        If (Not .Parent.IsPaused) And (Not .IO_EStop_PB) And .IO_MainPumpRunning Then
          State = AVDosePause
          Timer.Restart
        End If
        
      Case AVDose
        'Check level - if level low switch to circulate
        If .AdditionLevel < Setpoint() Then State = AVDosePause
        'Check level - if level high then switch to transfer
        If .AdditionLevel > (Setpoint() + 30) Then State = AVTransfer
        'If pump not running go to pause state
        If .Parent.IsPaused Or .IO_EStop_PB Or Not .IO_MainPumpRunning Then State = AVPause
        'If we're finished transfer empty
        If Timer.Finished Then
          State = AVTransferEmpty1
          Timer = .Parameters.AddTransferTimeBeforeRinse
        End If
      
      Case AVTransfer
        'Check level - if level low switch to circulate
        If .AdditionLevel < Setpoint() Then State = AVDosePause
        'If pump not running go to pause state
        If Not .IO_MainPumpRunning Then State = AVPause
        'If we're finished transfer empty
        If Timer.Finished Then
          State = AVTransferEmpty1
          Timer = .Parameters.AddTransferTimeBeforeRinse
        End If


      Case AVDosePause
        'Check level - if high switch to dose
        If .AdditionLevel > Setpoint() Then State = AVDose
        'If pump not running go to pause state
        If Not .IO_MainPumpRunning Then State = AVPause
        'If we're finished transfer empty
        If Timer.Finished Then
          State = AVTransferEmpty1
          Timer = .Parameters.AddTransferTimeBeforeRinse
        End If
      
      Case AVTransferEmpty1
        If .AdditionLevel > 0 Then Timer = .Parameters.AddTransferTimeBeforeRinse
        If Timer.Finished Then
          State = AVRinse
          Timer = .Parameters.AddTransferRinseTime
        End If
        
      Case AVRinse
        If Timer.Finished Then
          State = AVTransferEmpty2
          Timer = .Parameters.AddTransferTimeAfterRinse
        End If
         
      Case AVTransferEmpty2
        If .AdditionLevel > 0 Then Timer = .Parameters.AddTransferTimeAfterRinse
        If Timer.Finished Then
          .AddReady = False
          .AP.ACCommand_Cancel
          State = AVRinseToDrain
          Timer = .Parameters.AddTransferRinseToDrainTime
        End If
      
      Case AVRinseToDrain
        If Timer.Finished Then
          State = AVTransferToDrain
          Timer = .Parameters.AddTransferToDrainTime
        End If
        
      Case AVTransferToDrain
        If .AdditionLevel > 0 Then Timer = .Parameters.AddTransferToDrainTime
        If Timer.Finished Then
          .AddReady = False
          .AP.ACCommand_Cancel
          State = AVOff
        End If
    End Select
  End With

End Sub
Friend Sub ACCommand_Cancel()
  State = AVOff
End Sub
Friend Property Get IsOn() As Boolean
  IsOn = (State <> AVOff)
End Property
Friend Property Get IsActive() As Boolean
  If IsTransfer Or (State = AVDosePause) Or IsRinse Then IsActive = True
End Property
Friend Property Get IsWaitReady() As Boolean
  If (State = AVWaitReady) Then IsWaitReady = True
End Property
Friend Property Get IsTransfer() As Boolean
  If (State = AVDose) Or (State = AVTransfer) Or _
     (State = AVTransferEmpty1) Or (State = AVTransferEmpty2) Then IsTransfer = True
End Property
Friend Property Get IsDosing() As Boolean
  If (State = AVDose) Then IsDosing = True
End Property
Friend Property Get IsDosePaused() As Boolean
  If (State = AVDosePause) Then IsDosePaused = True
End Property
Friend Property Get IsRinse() As Boolean
  If (State = AVRinse) Then IsRinse = True
End Property
Friend Property Get IsRinseToDrain() As Boolean
  If (State = AVRinseToDrain) Then IsRinseToDrain = True
End Property
Friend Property Get IsTransferToDrain() As Boolean
  If (State = AVTransferToDrain) Then IsTransferToDrain = True
End Property
Friend Property Get IsDelayed() As Boolean
  If (State = AVWaitReady) Then IsDelayed = True
End Property
'Function uses the expressions (y = x * x) and (y = sqr(x)) to generate curves for
'dosing control.
'x = A fraction between 0 and 1 (I know) representing elapsed time
'y = A fraction between 0 and 1 representing the amount we should have transferred so far
'
'The curves are scaled by adding a percent of the difference between the curve value and
'the equivalent linear value to the original curve value (??).
Private Function Setpoint() As Long
'If timer has finished exit function
  If Timer.Finished Then
    Setpoint = 0
    Exit Function
  End If
  
'Amount we should have transferred so far
  Dim ElapsedTime As Double, LinearTerm As Double, TransferAmount As Double
  ElapsedTime = (AddTime - Timer) / AddTime
  LinearTerm = ElapsedTime
  TransferAmount = StartLevel * LinearTerm
  
'Calculate scaling factor (0-1) for progressive and digressive curves
  If AddCurve > 0 Then
    Dim ScalingFactor As Double
    ScalingFactor = (10 - AddCurve) / 10
'Calculate term for progressive transfer (0-1) if odd curve
    If (AddCurve Mod 2) = 1 Then
      Dim MaxOddCurve As Double, OddTerm As Double
      MaxOddCurve = (ElapsedTime * ElapsedTime)
      OddTerm = MaxOddCurve + ((LinearTerm - MaxOddCurve) * ScalingFactor)
      TransferAmount = StartLevel * OddTerm
    Else
'Calculate term for digressive transfer (0-1) if even curve
      Dim MaxEvenCurve As Double, EvenTerm As Double
      MaxEvenCurve = Sqr(ElapsedTime)
      EvenTerm = MaxEvenCurve - ((MaxEvenCurve - LinearTerm) * ScalingFactor)
      TransferAmount = StartLevel * EvenTerm
    End If
  End If
   
'Calculate setpoint and limit to 0-1000
  Setpoint = StartLevel - CLng(TransferAmount)
  If Setpoint < 0 Then Setpoint = 0
  If Setpoint > 1000 Then Setpoint = 1000
  
'Global variable for display purposes - yuk!!
  DesiredLevel = Setpoint
End Function
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property


#End If
End Class
