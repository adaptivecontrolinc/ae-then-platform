'American & Efird - Then Platform
' Version 2024-07-22

Imports Utilities.Translations

<Command("Wait Temp", "", "", "", ""), _
Description("Wait at this step until desired temp is reached."), _
TranslateDescription("es", "Esperar a este paso hasta que la temperatura deseada."), _
Category("Temperature Functions")> _
Public Class WT
  Inherits MarshalByRefObject
  Implements ACCommand
  Private Const commandName_ As String = ("WT: ")

  'Command States
  Public Enum EState
    Off
    Active
  End Enum
  Public State As EState
  Public Status As String

  Private ReadOnly ControlCode As ControlCode
  Sub New(ByVal controlCode As ControlCode)
    Me.ControlCode = controlCode
  End Sub

  Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
  End Sub

  Public Sub ParametersChanged(ByVal ParamArray param() As Integer) Implements ACCommand.ParametersChanged
  End Sub

  Public Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With ControlCode

      'Cancel Commands
      .AC.Cancel() : .AT.Cancel()
      .WK.Cancel() ' .KA.Cancel() : 
      .DR.Cancel() : .FI.Cancel() : .HC.Cancel() : .HD.Cancel()
      .RI.Cancel() : .RC.Cancel() : .RH.Cancel() : .TM.Cancel()
      .LD.Cancel() : .PH.Cancel() : .SA.Cancel() : .UL.Cancel()
      If .RF.FillType = EFillType.Vessel Then .RF.Cancel()
      If .RT.IsForeground Then .RT.Cancel()
      .RW.Cancel()

      State = EState.Active
      If Not (.TP.IsOn OrElse .HE.IsOn OrElse .CO.IsOn) Then Cancel()
    End With
  End Function

  Public Function Run() As Boolean Implements ACCommand.Run
    With ControlCode
      Select Case State

        Case EState.Off
          Status = (" ")

        Case EState.Active
          Status = Translate("Waiting") & (" ") & (.TemperatureControl.FinalTemp \ 10).ToString("#0") & "F"
          If .TemperatureControl.Pid.IsHold Then
            'We've finished the ramp and we're within the step margin
            Dim tempError As Integer = Math.Abs(.TemperatureControl.FinalTemp - .Temp)
            If .TemperatureControl.IsCooling AndAlso tempError < .Parameters.CoolStepMargin Then Cancel()
            If .TemperatureControl.IsHeating AndAlso tempError < .Parameters.HeatStepMargin Then Cancel()
          End If

      End Select
    End With
  End Function

  ReadOnly Property IsOn As Boolean Implements ACCommand.IsOn
    Get
      IsOn = (State <> EState.Off)
    End Get
  End Property

End Class
