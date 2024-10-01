'American & Efird - MX Package
' Version 2024-07-22

Imports Utilities.Translations

<Command("Temperature", "|0-9|.|0-9| TO |0-280|F", "('1*10) + '2", "'3", ""),
TranslateCommand("es", "Temperatura", "|0-9|.|0-9| A |0-280|F"),
Description("Heat at the given gradient to the given target temperature.  A gradient and target of 0 disables any previous control."),
TranslateDescription("es", "El calor/se refresca en el gradiente dado a la temperatura dada de la blanco. Un gradiente y una blanco de 0 inhabilita cualquier control"),
 Category("Temperature Functions"), TranslateCategory("es", "Temperature Functions")>
Public Class TP
  Inherits MarshalByRefObject
  Implements ACCommand
  Private Const commandName_ As String = ("TP: ")

  'Command States
  Public Enum EState
    Off
    Start
    Ramp
    Hold
  End Enum
  Public State As EState
  Public Status As String
  Public Gradient As Integer
  Public FinalTemp As Integer
  Friend Timer As New Timer

  'Keep a local reference to the control code object for convenience
  Private ReadOnly controlCode As ControlCode

  Public Sub New(ByVal controlCode As ControlCode)
    Me.controlCode = controlCode
    Me.controlCode.Commands.Add(Me)
  End Sub

  Public Sub ParametersChanged(ByVal ParamArray param() As Integer) Implements ACCommand.ParametersChanged
    'If the command is not running then do nothing the changes will be picked up when the command starts
    If Not IsOn Then Exit Sub

    'Enable command parameter adjustments, if parameter set
    If controlCode.Parameters.EnableCommandChg = 1 Then
      Start(param)
    End If
  End Sub

  Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With controlCode

      'Cancel Commands
      If .AC.IsForeground Then .AC.Cancel()
      If .AT.IsForeground Then .AT.Cancel()
      .KA.Cancel() : .WK.Cancel()
      .DR.Cancel() : .FI.Cancel() : .HC.Cancel() : .HD.Cancel()
      .RC.Cancel() : .RH.Cancel() : .RI.Cancel() : .TM.Cancel() ' .TC.Cancel() ' TODO
      .LD.Cancel() : .PH.Cancel() : .SA.Cancel() : .UL.Cancel()
      If .RF.FillType = EFillType.Vessel Then .RF.Cancel()
      If .RT.IsForeground Then .RT.Cancel()
      .RW.Cancel()
      .CO.Cancel() : .HE.Cancel() ': .TP.Cancel() : 
      .WT.Cancel()

      '  .TemperatureControlContacts.Cancel ' TODO

      ' Clear Parent.Signal 
      If .Parent.Signal <> "" Then .Parent.Signal = ""

      'Check array bounds just to be on the safe side
      If param.GetUpperBound(0) >= 1 Then Gradient = param(1) * 10
      If param.GetUpperBound(0) >= 2 Then Gradient += param(2)
      If param.GetUpperBound(0) >= 3 Then FinalTemp = param(3) * 10

      'No Gradient or Final temp - Cancel command
      If (Gradient = 0 AndAlso FinalTemp = 0) OrElse (FinalTemp > 2800) Then Cancel()

      'For compatibility - max gradient used to be 99 rather than 0
      If Gradient = 99 Then Gradient = 0

      'Set default start state
      State = EState.Start

      'This is a background command so step on
      Return True
    End With
  End Function

  Function Run() As Boolean Implements ACCommand.Run
    With controlCode

      Select Case State

        Case EState.Off
          Status = (" ")

        Case EState.Start
          Status = Translate("Starting")
          .TemperatureControl.Start(.Temp, FinalTemp, Gradient)
          .AirpadOn = True
          State = EState.Ramp

        Case EState.Ramp
          Status = .TemperatureControl.Status
          If .TemperatureControl.Pid.IsHold Then
            'We've finished the ramp and we're within the step margin
            Dim tempError As Integer = Math.Abs(.TemperatureControl.FinalTemp - .Temp)
            If .TemperatureControl.IsCooling AndAlso tempError < .Parameters.CoolStepMargin Then State = EState.Hold
            If .TemperatureControl.IsHeating AndAlso tempError < .Parameters.HeatStepMargin Then State = EState.Hold
          End If

        Case EState.Hold
          Status = Translate("Holding") & (" ") & (FinalTemp / 10).ToString("#0") & "F"

      End Select
    End With
  End Function

  Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
  End Sub

  ReadOnly Property IsOn() As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property

  ReadOnly Property IsActive() As Boolean
    Get
      Return (State >= EState.Start) AndAlso (State <= EState.Ramp)
    End Get
  End Property

#If 0 Then


Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode
    Select Case State
      Case TPOff
      
      Case TPStart
        State = TPRamp
        StepOn = True
        .TemperatureControl.CoolingIntegral = .TemperatureControl.Parameters.CoolIntegral
        .TemperatureControl.CoolingMaxGradient = .TemperatureControl.Parameters.CoolMaxGradient
        .TemperatureControl.CoolingPropBand = .TemperatureControl.Parameters.CoolPropBand
        .TemperatureControl.CoolingStepMargin = .TemperatureControl.Parameters.CoolStepMargin
        .TemperatureControl.Start .VesTemp, FinalTemp, Gradient
        'Check Temperature mode - Change during TPHold if necessary
        .TemperatureControl.TempMode = 0
        If .TemperatureControl.Parameters.HeatCoolModeChange = 1 Then .TemperatureControl.TempMode = 2
        If .TemperatureControl.Parameters.HeatCoolModeChange = 2 Then .TemperatureControl.TempMode = 2
      
      Case TPRamp
        If ((.VesTemp > (FinalTemp - .TemperatureControl.Parameters.CoolStepMargin)) And _
           (.VesTemp < (FinalTemp + .TemperatureControl.Parameters.HeatStepMargin))) Then
           State = TPHold
        End If
      
      Case TPHold
        'Change mode to Heat and Cool if necessary
        If .TemperatureControl.Parameters.HeatCoolModeChange = 1 Then .TemperatureControl.TempMode = 0
    End Select
  End With
End Sub

Friend Property Get IsActive() As Boolean
  If IsRamping Or IsHolding Then IsActive = True
End Property
Friend Property Get IsRamping() As Boolean
  If (State = TPRamp) Then IsRamping = True
End Property
Friend Property Get IsHolding() As Boolean
  If (State = TPHold) Then IsHolding = True
End Property

#End If

End Class
