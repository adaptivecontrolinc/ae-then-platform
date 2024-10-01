Public Class TC
  ' TODO - Used?
#If 0 Then


  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Temperature Command
Option Explicit
Implements ACCommand
Public Enum TCState
  TCOff
  TCStart
  TCRamp
  TCHold
End Enum
Public State As TCState
Attribute State.VB_VarUserMemId = 0
Public StateString As String
Public DegreesPerContact As Long
Public Gradient As Long
Public FinalTemp As Long
Public RateofRise As Long
Public OverrunTimer As New acTimer

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Parameters=|0-99|.|0-9|  TO |0-280|F\r\nName=Heat\r\nGradient=10\r\nTarget='3\r\nHelp=Heat/Cool at the given gradient per contact to the given target temperature.  A gradient and target of 0 disables any previous control."
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject

  With ControlCode
    .AC.ACCommand_Cancel
    .AT.ACCommand_Cancel: .CO.ACCommand_Cancel: .DR.ACCommand_Cancel
    .FI.ACCommand_Cancel: .HE.ACCommand_Cancel: .HD.ACCommand_Cancel
    .HC.ACCommand_Cancel: .LD.ACCommand_Cancel: .PH.ACCommand_Cancel
    .RC.ACCommand_Cancel: .RH.ACCommand_Cancel: .RI.ACCommand_Cancel
    If .RF.FillType = 86 Then .RF.ACCommand_Cancel:
    .RT.ACCommand_Cancel: .RW.ACCommand_Cancel: .SA.ACCommand_Cancel
    .TM.ACCommand_Cancel: .TP.ACCommand_Cancel: .UL.ACCommand_Cancel
    .WK.ACCommand_Cancel: .WT.ACCommand_Cancel
    .TemperatureControl.Cancel
   
     State = TCStart
     
    DegreesPerContact = (Param(1) * 10) + Param(2) 'in tenths
    Gradient = .NumberofContactsPerMin * DegreesPerContact
    FinalTemp = (Param(3) * 10)
    
    'No gradient or final temp - cancel command
    If (Gradient = 0 And FinalTemp = 0) Or (FinalTemp > 2800) Then ACCommand_Cancel
    
    'For compatibility - max gradient used to be 99 rather than 0
    If DegreesPerContact = 999 Then Gradient = 0
    RateofRise = Gradient
    If RateofRise = 0 Then RateofRise = 50
    If FinalTemp > .VesTemp Then
      OverrunTimer = (((FinalTemp - .VesTemp) / RateofRise) * 60) + 60
    Else
      OverrunTimer = (((.VesTemp - FinalTemp) / RateofRise) * 60) + 60
    End If
    
 End With
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode
  
    Static pidPaused As Boolean
    If .TemperatureControlContacts.IsPidPaused Then
      pidPaused = True
      OverrunTimer.Pause
    Else
      If pidPaused = True Then
         pidPaused = False
         OverrunTimer.Restart
      End If
    End If
    
    Select Case State
      Case TCOff
        StateString = ""
        
      Case TCStart
        StateString = "TC: Starting "
        Gradient = .NumberofContactsPerMin * DegreesPerContact
        State = TCRamp
        .TemperatureControlContacts.CoolingIntegral = .TemperatureControl.Parameters.CoolIntegral
        .TemperatureControlContacts.CoolingMaxGradient = .TemperatureControl.Parameters.CoolMaxGradient
        .TemperatureControlContacts.CoolingPropBand = .TemperatureControl.Parameters.CoolPropBand
        .TemperatureControlContacts.CoolingStepMargin = .TemperatureControl.Parameters.CoolStepMargin
        .TemperatureControlContacts.Start .VesTemp, FinalTemp, Gradient, ControlCode
        'Check Temperature mode - Change during TCHold if necessary
        .TemperatureControlContacts.TempMode = 0
        If .TemperatureControl.Parameters.HeatCoolModeChange = 1 Then .TemperatureControlContacts.TempMode = 2
        If .TemperatureControl.Parameters.HeatCoolModeChange = 2 Then .TemperatureControlContacts.TempMode = 2
      
      Case TCRamp
        If Not .PumpRequest Then
          StateString = "TC: Pump Request Off "
        ElseIf Not .IO_MainPumpRunning Then
          StateString = "TC: Pump Running Signal Lost "
        Else
          If .TemperatureControlContacts.IsHeating Or .TemperatureControlContacts.IsPreHeatVent Then
             StateString = "TC: Heat to " & Pad(.TemperatureControlContacts.TempFinalTemp / 10, "0", 3) & "F"
          ElseIf .TemperatureControlContacts.IsCooling Or .TemperatureControlContacts.IsPreCoolVent Then
             StateString = "TC: Cool to " & Pad(.TemperatureControlContacts.TempFinalTemp / 10, "0", 3) & "F"
          Else
             StateString = "TC: Waiting to heat/cool"
          End If
        End If
        Gradient = .NumberofContactsPerMin * DegreesPerContact
        If .TemperatureControlContacts.IsHolding Then
           If ((.VesTemp > (FinalTemp - .TemperatureControl.Parameters.CoolStepMargin)) And _
              (.VesTemp < (FinalTemp + .TemperatureControl.Parameters.HeatStepMargin))) Then
              StepOn = True
              State = TCHold
           End If
        End If
      
      Case TCHold
        StateString = "TC: Holding " & Pad(.TemperatureControl.TempFinalTemp / 10, "0", 3) & "F"
        Gradient = .NumberofContactsPerMin * DegreesPerContact
        If .TemperatureControl.Parameters.HeatCoolModeChange = 1 Then .TemperatureControlContacts.TempMode = 0
    
    End Select
  End With
End Sub
Friend Sub ACCommand_Cancel()
  State = TCOff
End Sub
Friend Property Get IsOn() As Boolean
  IsOn = (State <> TCOff)
End Property
Friend Property Get IsActive() As Boolean
  If IsRamping Or IsHolding Then IsActive = True
End Property
Friend Property Get IsRamping() As Boolean
  If (State = TCRamp) Then IsRamping = True
End Property
Friend Property Get IsHolding() As Boolean
  If (State = TCHold) Then IsHolding = True
End Property
Friend Property Get IsOverrun() As Boolean
  IsOverrun = IsRamping And OverrunTimer.Finished
End Property
Private Property Get ACCommand_IsOn() As Boolean
ACCommand_IsOn = IsOn
End Property



#End If
End Class
