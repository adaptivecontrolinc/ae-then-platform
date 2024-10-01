'American & Efird
' Version 2024-09-16

<Command("Flow Percent", "I:|0-100|% O:|0-100|%", "", " ", ""),
  Description("Sets the flow percent, pump speed output, for In-to-out and Out-to-in flow."),
  Category("Flow Control")>
Public Class FP : Inherits MarshalByRefObject : Implements ACCommand
  Private ReadOnly controlCode As ControlCode

  Public Enum EState
    Off
    Active
  End Enum
  Property State As EState
  Property InToOutPercent As Integer
  Property OutToInPercent As Integer

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
      If param.GetUpperBound(0) >= 1 Then InToOutPercent = param(1) * 10
      If param.GetUpperBound(0) >= 2 Then OutToInPercent = param(2) * 10
    End If

  End Sub

  Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With controlCode
      .DP.Cancel()
      .FL.Cancel()

      'Check array bounds just to be on the safe side
      If param.GetUpperBound(0) >= 1 Then InToOutPercent = param(1) * 10
      If param.GetUpperBound(0) >= 2 Then OutToInPercent = param(2) * 10

      ' Use default percent output when set to 0?
      If InToOutPercent = 0 Then InToOutPercent = .PumpControl.Parameters_PumpSpeedDefault
      If OutToInPercent = 0 Then OutToInPercent = .PumpControl.Parameters_PumpSpeedDefault

      'Set limits on pump speed
      InToOutPercent = MinMax(InToOutPercent, 0, 1000)
      OutToInPercent = MinMax(OutToInPercent, 0, 1000)

      'Set default state
      State = EState.Active

      'This is a background command so do not step on
      Return True
    End With
  End Function

  Function Run() As Boolean Implements ACCommand.Run
  End Function

  Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
  End Sub

  ReadOnly Property IsOn As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property


  ' TODO CHeck
#If 0 Then


  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Flow Percent command - pump speed
Option Explicit
Implements ACCommand
Public Enum FPState
  FPoff
  FPOn
End Enum
Public State As FPState
Public FPPercentInToOut As Long
Public FPPercentOutToIn As Long

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Parameters= IO |0-100|% OI |0-100|%\r\nName=Flow Percentage\r\nHelp=Sets pump speed for in to out and out to in"
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode
    .DP.ACCommand_Cancel
    .FL.ACCommand_Cancel
    If Param(1) > 0 Then
      FPPercentInToOut = Param(1) * 10
    Else
      FPPercentInToOut = .Parameters.PumpSpeedDefault
    End If
    If Param(2) > 0 Then
      FPPercentOutToIn = Param(2) * 10
    Else
      FPPercentOutToIn = .Parameters.PumpSpeedDefault
    End If
  End With
  State = FPOn
'Set limits on pump speed
  MinMax FPPercentInToOut, 0, 1000
  MinMax FPPercentOutToIn, 0, 1000
  StepOn = True
End Sub
Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
End Sub
Friend Sub ACCommand_Cancel()
State = FPoff
End Sub
Friend Property Get OutputInToOut() As Long
  OutputInToOut = FPPercentInToOut
End Property
Friend Property Get OutputOutToIn() As Long
  OutputOutToIn = FPPercentOutToIn
End Property
Public Property Get IsOn() As Boolean
  IsOn = (State <> FPoff)
End Property
Private Property Get ACCommand_IsOn() As Boolean
ACCommand_IsOn = IsOn
End Property



#End If
End Class
