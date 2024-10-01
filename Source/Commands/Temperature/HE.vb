'American & Efird - Mt. Holly Then Platform to GMX
' Version 2022-02-03

Imports Utilities.Translations

'//NOTE: 
'   Use Custom ProgramGroups database field 'MaximumGradient' to customize the standard time calculated based on a machine specific MaximumGradient.  
'   Refer to ControlCode.InitializeLocalDatabase

<Command("Heat", "|0-9|.|0-9| TO |0-280|F", "('1*10) + '2", "'3", ""),
  TranslateCommand("es", "Calentamiento", "|0-9|.|0-9| A |0-280|F"),
  Description("Heat at the desired gradient to the desired target temperature. A gradient and target of 0 disables previous control."),
  TranslateDescription("es", "Calor en el gradiente dado a la temperatura dada de 'la blanco. Un gradiente y una blanco de 0 inhabilita cualquier 'control anterior."),
  Category("Temperature Functions"), TranslateCategory("es", "Temperature Functions")>
Public Class HE : Inherits MarshalByRefObject : Implements ACCommand
  Private Const commandName_ As String = ("HE: ")
  Private ReadOnly controlCode As ControlCode

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
      If .RF.IsForeground AndAlso .RF.FillType = EFillType.Vessel Then .RF.Cancel()
      If .RT.IsForeground Then .RT.Cancel()
      .RW.Cancel()
      .CO.Cancel() '.HE.Cancel() : ' .TC.Cancel
      .TP.Cancel() : .WT.Cancel()

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

      'Calculate time it should take to ramp to final temp
      Dim rateOfRise As Integer = Gradient
      If FinalTemp > .Temp Then
        If (rateOfRise = 0) Then rateOfRise = MinMax(.Parameters.HeatMaxGradient, 5, 250)
        'Avoid potential divide by zero
        If rateOfRise > 0 Then
          TimerOverrun.Minutes = Convert.ToInt32((((FinalTemp - .Temp) / rateOfRise)) + 10)
        End If
      Else
        If (rateOfRise = 0) Then rateOfRise = MinMax(.Parameters.CoolMaxGradient, 5, 250)
        'Avoid potential divide by zero
        If rateOfRise > 0 Then
          TimerOverrun.Minutes = Convert.ToInt32((((.Temp - FinalTemp) / rateOfRise)) + 5)
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
            StateString = Translate("Heating not enabled")
            If Not .TempValid OrElse (.Temp > 2800) Then StateString = Translate("Temp > 280.0 F Or Not Valid")
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
          If .TemperatureControl.Pid.IsHold Then
            'We've finished the ramp and we're within the step margin
            If (FinalTemp - .Temp) < .Parameters.HeatStepMargin Then
              State = EState.Hold
              Return True  'Step on
            End If
          End If
          If (.TemperatureControl.State = Temperature.EState.Heat) OrElse (.TemperatureControl.State = Temperature.EState.Cool) Then
            If .TemperatureControl.Pid.IsMaxGradient OrElse (.TemperatureControl.Pid.PidSetpoint = .TemperatureControl.Pid.FinalTemp) Then
              StateString = Translate("Ramping to") & (" ") & (.TemperatureControl.Pid.FinalTemp / 10).ToString("#0") & "F"
            Else
              StateString = Translate("Ramping to") & (" ") & (.TemperatureControl.Pid.PidSetpoint / 10).ToString("#0") & " / " & (.TemperatureControl.Pid.FinalTemp / 10).ToString("#0") & "F"
            End If
          Else : StateString = .TemperatureControl.Status
          End If

        Case EState.Hold
          StateString = Translate("Holding") & (" ") & (FinalTemp / 10).ToString("#0") & "F"

      End Select
    End With
  End Function

  Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
    StateString = ""
    Timer.Cancel()
    TimerOverrun.Cancel()
  End Sub

  ReadOnly Property IsOn() As Boolean Implements ACCommand.IsOn
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

End Class
