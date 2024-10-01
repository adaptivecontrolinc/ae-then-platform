'American & Efird - Then Platform
' Version 2024-09-26

Imports Utilities.Translations

<Command("Flow Rate", "I:|0-99|% O:|0-99|%", "", " ", ""), _
Description("Sets the desired percentage of gallons per minute flowrate, for In-to-out and Out-to-in flow."), _
Category("Flow Control")> _
Public Class FL
  Inherits MarshalByRefObject
  Implements ACCommand

  'Command States
  Public Enum EState
    Off
    Active
  End Enum
  Property State As EState
  Property Status As String
  Property InToOutFlow As Integer
  Property OutToInFlow As Integer

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
      'Check array bounds just to be on the safe side
      If param.GetUpperBound(0) >= 1 Then InToOutFlow = CInt((param(1) / 100) * controlCode.Parameters.FlowRateMachineRange)
      If param.GetUpperBound(0) >= 2 Then OutToInFlow = CInt((param(2) / 100) * controlCode.Parameters.FlowRateMachineRange)
    End If

  End Sub

  Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With controlCode
      .FP.Cancel()

      'Check array bounds just to be on the safe side
      If param.GetUpperBound(0) >= 1 Then InToOutFlow = CInt((param(1) / 100) * controlCode.Parameters.FlowRateMachineRange)
      If param.GetUpperBound(0) >= 2 Then OutToInFlow = CInt((param(2) / 100) * controlCode.Parameters.FlowRateMachineRange)

      'Set default state
      State = EState.Active
    End With

    'This is a background command so step on
    Return True
  End Function

  Function Run() As Boolean Implements ACCommand.Run
    With controlCode

      Select Case State
        Case EState.Off
          Status = (" ")

        Case EState.Active
          If .PumpControl.IsOutToIn Then
            Status = Translate("Desired Flowrate") & (" ") & (OutToInFlow / 10).ToString("#0.0") & " gpm"
          Else
            Status = Translate("Desired Flowrate") & (" ") & (InToOutFlow / 10).ToString("#0.0") & " gpm"
          End If

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


#If 0 Then
  ' TODO

  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Controls pump speed based off of the flow rate.
Option Explicit
Implements ACCommand
Public Enum FLState
  FLOff
  FLOn
End Enum
Public State As FLState
Public InToOutFlow As Long
Public OutToInFlow As Long
Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Name=Flow Rate\r\nParameters=I:|0-99|% O:|0-99|%\r\nHelp=Sets the desired percentage of gallons per minute flowrate. percentage is based off of the maximum flowrate."
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode
    .FP.ACCommand_Cancel
    If .Parameters.FlowRateRange > 0 Then
      InToOutFlow = ((Param(1) / 100) * .Parameters.FlowRateRange)
      OutToInFlow = ((Param(2) / 100) * .Parameters.FlowRateRange)
    End If
  End With
  State = FLOn
  StepOn = True
End Sub
Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
End Sub
Friend Sub ACCommand_Cancel()
  State = FLOff
End Sub
Friend Property Get InToOutFlowRate() As Long 'in gallons per min
  InToOutFlowRate = InToOutFlow
End Property
Friend Property Get OutToInFlowRate() As Long 'in gallons per min
  OutToInFlowRate = OutToInFlow
End Property
Friend Property Get IsOn() As Boolean
  IsOn = (State <> FLOff)
End Property
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property



#End If
End Class
