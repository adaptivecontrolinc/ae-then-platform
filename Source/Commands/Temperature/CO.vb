'American & Efird - Mt. Holly Then Platform to GMX
' Version 2022-05-23

Imports Utilities.Translations

'//NOTE: 
'   Use Custom ProgramGroups database field 'MaximumGradient' to customize the standard time calculated based on a machine specific MaximumGradient.  
'   Refer to ControlCode.InitializeLocalDatabase

<Command("Cool", "|0-9|.|0-9| TO |0-280|F", " ('1*10) + '2", "'3", ""),
  TranslateCommand("es", "Enfriamiento", "|0-9|.|0-9| A |0-280|F"),
  Description("Cool at the desired gradient to the desired target temperature. A gradient and target of 0 disables previous control."),
  TranslateDescription("es", "Refresqúese en el gradiente dado a la temperatura dada de 'la blanco. Un gradiente y una blanco de 0 inhabilita cualquier 'control anterior."),
  Category("Temperature Functions"), TranslateCategory("es", "Temperature Functions")>
Public Class CO
  Inherits MarshalByRefObject
  Implements ACCommand
  Private Const commandName_ As String = ("CO: ")
  Public Enum EState
    Off
    Interlock
    Start
    Ramp
    Hold
  End Enum
  Property State As EState
  Property StateString As String
  Property Gradient As Integer
  Property FinalTemp As Integer
  Property RateOfRise As Integer

  Friend ReadOnly Timer As New Timer
  Friend ReadOnly TimerAdvance As New Timer
  Public ReadOnly Property TimeAdvance As Integer
    Get
      Return TimerAdvance.Seconds
    End Get
  End Property
  Friend TimerOverrun As New Timer
  Public ReadOnly Property TimeOverrun As Integer
    Get
      Return TimerOverrun.Seconds
    End Get
  End Property

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
      .RC.Cancel() : .RH.Cancel() : .RI.Cancel() : .TM.Cancel()
      .LD.Cancel() : .PH.Cancel() : .SA.Cancel() : .UL.Cancel()
      If .RT.IsForeground Then .RT.Cancel()
      ' .CO.Cancel() :
      .HE.Cancel() : .TP.Cancel() : .WT.Cancel()

      ' Clear Parent.Signal 
      If .Parent.Signal <> "" Then .Parent.Signal = ""

      'Check array bounds just to be on the safe side
      If param.GetUpperBound(0) >= 1 Then Gradient = param(1) * 10
      If param.GetUpperBound(0) >= 2 Then Gradient += param(2)
      If param.GetUpperBound(0) >= 3 Then FinalTemp = param(3) * 10

      'No Gradient or Final temp - Cancel command
      If (Gradient = 0 AndAlso FinalTemp = 0) OrElse (FinalTemp > 1400) Then Cancel()

      'For compatibility - max gradient used to be 99 rather than 0
      If Gradient = 99 Then Gradient = 0

      'Calculate time it should take to ramp to final temp
      Dim rateOfRise As Integer = Gradient
      If FinalTemp > .Temp Then
        If (rateOfRise = 0) Then rateOfRise = MinMax(.Parameters.HeatMaxGradient, 5, 250)
        'Avoid potential divide by zero
        If rateOfRise > 0 Then
          TimerOverrun.Minutes = Convert.ToInt32((((FinalTemp - .Temp) / rateOfRise)) + 1)
        End If
      Else
        If (rateOfRise = 0) Then rateOfRise = MinMax(.Parameters.CoolMaxGradient, 5, 250)
        'Avoid potential divide by zero
        If rateOfRise > 0 Then
          TimerOverrun.Minutes = Convert.ToInt32((((.Temp - FinalTemp) / rateOfRise)) + 1)
        End If
      End If

      ' Set Default Start state
      If (FinalTemp > .SafetyControl.Parameters_PressurizationTemperature) Then
        If .MachineClosed Then
          ' All closed - continue
          State = EState.Start
        Else
          ' Signal to check lids, if not already set
          If Not .MachineClosed Then
            .Parent.Signal = Translate("Close Machine Lid")
          End If
          ' Something isn't closed (or flagged closed with the Expansion tank)
          State = EState.Interlock
        End If
      Else
        ' Low Temp Setpoint
        State = EState.Start
      End If

      'This is a foreground command so do not step on
      Return False
    End With
  End Function

  Function Run() As Boolean Implements ACCommand.Run
    With controlCode

      'Pause the overrun timer if temperature control is paused
      If .TemperatureControl.IsPaused Then
        TimerOverrun.Pause()
      Else
        If TimerOverrun.Paused Then TimerOverrun.Restart()
      End If

      Select Case State

        Case EState.Off
          StateString = (" ")

        Case EState.Interlock
          StateString = Translate("Interlock")
          If Not .MachineClosed Then StateString = Translate("Close Machine Lid")
          ' Reset Timer
          If Not .AdvancePb Then TimerAdvance.Seconds = 2
          If .MachineClosed Then
            StateString = Translate("Hold Advance") & (" ") & TimerAdvance.ToString
            If TimerAdvance.Finished Then
              If .Parent.Signal <> "" Then .Parent.Signal = ""
              State = EState.Start
            End If
          End If

        Case EState.Start
          If Not .HeatEnabled Then
            StateString = Translate("Cooling not enabled")
            If Not .TempValid OrElse (.Temp > 2800) Then StateString = Translate("Temp > 280.0F Or Not Valid")
            If .IO.ContactTemp Then StateString = Translate("Contact Temp Input On")
            If Not (.PumpControl.IsRunning) Then StateString = Translate("Pump Not Running")
            If Not .MachineClosed Then StateString = Translate("Close Machine Lid")
          Else
            StateString = Translate("Starting")
            .TemperatureControl.Start(.Temp, FinalTemp, Gradient)
            .AirpadOn = True
            State = EState.Ramp
          End If

        Case EState.Ramp
          If (.TemperatureControl.State = Temperature.EState.Heat) OrElse (.TemperatureControl.State = Temperature.EState.Cool) Then
            If .TemperatureControl.Pid.IsMaxGradient OrElse (.TemperatureControl.Pid.PidSetpoint = .TemperatureControl.Pid.FinalTemp) Then
              StateString = Translate("Ramping to") & (" ") & (.TemperatureControl.Pid.FinalTemp / 10).ToString("#0") & "F"
            Else
              StateString = Translate("Ramping to") & (" ") & (.TemperatureControl.Pid.PidSetpoint / 10).ToString("#0") & " / " & (.TemperatureControl.Pid.FinalTemp / 10).ToString("#0") & "F"
            End If
          Else : StateString = .TemperatureControl.Status
          End If
          If .TemperatureControl.Pid.IsHold Then
            'We've finished the ramp and we're within the step margin
            If (.Temp - FinalTemp) < .Parameters.CoolStepMargin Then
              State = EState.Hold
              Return True  'Step on
            End If
          End If

        Case EState.Hold
          StateString = Translate("Holding") & (" ") & (FinalTemp / 10).ToString("#0") & "F"

      End Select
    End With
  End Function

  Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
  End Sub

  ReadOnly Property IsOn As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property

  ReadOnly Property IsActive() As Boolean
    Get
      Return (State >= EState.Interlock) AndAlso (State <= EState.Ramp)
    End Get
  End Property

  Public ReadOnly Property IsOverrun() As Boolean
    Get
      IsOverrun = IsActive AndAlso TimerOverrun.Finished
    End Get
  End Property


  ' TODO
#If 0 Then
VERSION 1.0 CLASS

Implements ACCommand
Public Enum COState
  COOff
  COStart
  CORamp
  COHold
End Enum
Public State As COState
Public StateString As String
Public Gradient As Long
Public FinalTemp As Long
Public RateofRise As Long
Public OverrunTimer As New acTimer


  With ControlCode
   .DR.ACCommand_Cancel: .FI.ACCommand_Cancel
    .HD.ACCommand_Cancel: .HE.ACCommand_Cancel: .HC.ACCommand_Cancel
    .LD.ACCommand_Cancel: .PH.ACCommand_Cancel: .RC.ACCommand_Cancel
    If .RF.FillType = 86 Then .RF.ACCommand_Cancel:
    .RH.ACCommand_Cancel: .RI.ACCommand_Cancel: .RT.ACCommand_Cancel
    .RW.ACCommand_Cancel: .SA.ACCommand_Cancel: .TC.ACCommand_Cancel
    .TM.ACCommand_Cancel: .TP.ACCommand_Cancel: .UL.ACCommand_Cancel
    .WK.ACCommand_Cancel: .WT.ACCommand_Cancel
    .TemperatureControlContacts.Cancel
    
    State = COStart
    
    Gradient = (Param(1) * 10) + Param(2)
    FinalTemp = (Param(3) * 10)
    
    'No gradient or final temp - cancel command
      If (Gradient = 0 And FinalTemp = 0) Or (FinalTemp > 2800) Then ACCommand_Cancel
    
    'For compatibility - max gradient used to be 99 rather than 0
      If Gradient = 999 Then Gradient = 0
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
    If .TemperatureControl.IsPidPaused Then
      pidPaused = True
      OverrunTimer.Pause
    Else
      If pidPaused = True Then
        pidPaused = False
        OverrunTimer.Restart
      End If
    End If
  
    Select Case State
    
      Case COOff
        StateString = ""
        
      Case COStart
        StateString = "CO: Starting "
        State = CORamp
        .TemperatureControl.CoolingIntegral = .TemperatureControl.Parameters.CoolIntegral
        .TemperatureControl.CoolingMaxGradient = .TemperatureControl.Parameters.CoolMaxGradient
        .TemperatureControl.CoolingPropBand = .TemperatureControl.Parameters.CoolPropBand
        .TemperatureControl.CoolingStepMargin = .TemperatureControl.Parameters.CoolStepMargin
        .TemperatureControl.Start .VesTemp, FinalTemp, Gradient
        'Check Temperature mode - Change during COHold if necessary
        .TemperatureControl.TempMode = 0
        If .TemperatureControl.Parameters.HeatCoolModeChange = 1 Then .TemperatureControl.TempMode = 2
        If .TemperatureControl.Parameters.HeatCoolModeChange = 2 Then .TemperatureControl.TempMode = 2
      
      Case CORamp
        If Not .IO_MainPumpRunning Then
          StateString = "CO: Pump Running Signal Lost "
        ElseIf Not .PumpRequest Then
          StateString = "CO: Pump Request Off "
        Else
          StateString = "CO: Ramp to " & Pad(.VesTemp / 10, "0", 3) & " / " & Pad(.TemperatureControl.TempFinalTemp / 10, "0", 3) & "F"
        End If
        If .TemperatureControl.IsHolding Then
           If ((.VesTemp > (FinalTemp - .TemperatureControl.Parameters.CoolStepMargin)) And _
              (.VesTemp < (FinalTemp + .TemperatureControl.Parameters.HeatStepMargin))) Then
              StepOn = True
              State = COHold
           End If
        End If
        
      Case COHold
        StateString = "CO: Holding " & Pad(.TemperatureControl.TempFinalTemp / 10, "0", 3) & "F"
        If .TemperatureControl.Parameters.HeatCoolModeChange = 1 Then .TemperatureControl.TempMode = 0
        
    End Select
  End With
End Sub
Friend Sub ACCommand_Cancel()
  State = COOff
End Sub
Friend Property Get IsOn() As Boolean
  IsOn = (State <> COOff)
End Property
Friend Property Get IsOverrun() As Boolean
  IsOverrun = IsRamping And OverrunTimer.Finished
End Property
Friend Property Get IsActive() As Boolean
  If IsRamping Or IsHolding Then IsActive = True
End Property
Friend Property Get IsRamping() As Boolean
  If (State = CORamp) Then IsRamping = True
End Property
Friend Property Get IsHolding() As Boolean
  If (State = COHold) Then IsHolding = True
End Property
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property

#End If
End Class
