'American & Efird - MX 
' Version 2016-05-26

Imports Utilities.Translations

<Command("Flow Contacts", "I:|0-99| O:|0-99|", "", "", ""), _
TranslateCommand("es", "Flow Reverse Time", "I:|0-99| O:|0-99|"), _
Description("Sets the time, in minutes, between flow reversals. I is minutes on In to Out flow, O is minutes on Out to In flow."), _
TranslateDescription("es", "Set flow reversal time."), _
Category("Flow Control"), TranslateCategory("es", "Flow Control")> _
Public Class FC
  Inherits MarshalByRefObject
  Implements ACCommand

  'Command States
  Public Enum EState
    Off
    Active
  End Enum
  Property State As EState
  Property Status As String
  Property InToOutTime As Integer
  Property OutToInTime As Integer

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
      ' Cancel Flow Reversal Commands
      .FR.Cancel()

      'Check Parameters
      If param.GetUpperBound(0) >= 1 Then InToOutTime = param(1)
      If param.GetUpperBound(0) >= 2 Then OutToInTime = param(2)

      ' Update Pump Control FlowReversal Parameters
      .PumpControl.UpdateFlowReversalTimes(InToOutTime, OutToInTime)

      ' Set default state
      State = EState.Active

      'This is a background command so step on
      Return True
    End With
  End Function

  Function Run() As Boolean Implements ACCommand.Run
    With controlCode

      Select Case State
        Case EState.Off
          Status = (" ")

        Case EState.Active
          Status = Translate("Active")

      End Select
    End With
  End Function

  Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
  End Sub

  Sub Clear()
    InToOutTime = 0
    OutToInTime = 0
  End Sub

  ReadOnly Property IsOn As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property

  ' TODO

#If 0 Then
  ' TODO REmove

  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Flow reverse command
Option Explicit
Implements ACCommand
Public Enum FCState
  FCOff
  FCOutToIn
  FCInToOut
  FCDecelInToOut
  FCDecelOutToIn
  FCSwitchToInToOut
  FCSwitchToOutToIn
End Enum
Public State As FCState
Public StateString As String
Public InToOutTurns As Long, OutToInTurns As Long
Public Timer As New acTimer
Public InToOut As Long, OutToIn As Long

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Name=Flow Contacts\r\nParameters=I:|0-99| O:|0-99|\r\nHelp=Sets the number of turns between flow reversals. I is turns on In to Out flow, O is turns on Out to In flow. "
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode
    .FR.ACCommand_Cancel
    InToOutTurns = Param(1)
    OutToInTurns = Param(2)
  End With
  If IsInToOut Then
    OutToIn = OutToIn + 1
  Else: InToOut = InToOut + 1
  End If
  
  State = FCInToOut
  StepOn = True
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode

    Select Case State
      
      Case FCOff
        StateString = ""
        
      Case FCOutToIn
        StateString = "FC: Out to In " & (OutToInTurns - .NumberOfContacts) & " / " & .NumberOfContacts
        If (.NumberOfContacts >= OutToInTurns) And (InToOutTurns > 0) Then
          State = FCDecelOutToIn
          .NumberOfContacts = 0
          .SystemVolume = 0
          Timer = MiniMaxi(.Parameters.PumpSpeedDecelTime, 15, 40)
        End If
      
      Case FCDecelOutToIn
        StateString = "FC: Decel Out to In " & TimerString(Timer.TimeRemaining)
         If Timer.Finished Then
          State = FCSwitchToInToOut
          .NumberOfContacts = 0
          .SystemVolume = 0
          Timer = .Parameters.FlowReverseTime
        End If
     
      Case FCSwitchToInToOut
        StateString = "FC: Switch In to Out " & TimerString(Timer.TimeRemaining)
        If Timer.Finished Then
          InToOut = InToOut + 1
          State = FCInToOut
          .NumberOfContacts = 0
          .SystemVolume = 0
        End If
     
      Case FCInToOut
        StateString = "FC: In to Out " & (InToOutTurns - .NumberOfContacts) & " / " & .NumberOfContacts
        If (.NumberOfContacts >= InToOutTurns) And (OutToInTurns > 0) Then
          State = FCDecelInToOut
          Timer = MiniMaxi(.Parameters.PumpSpeedDecelTime, 15, 40)
          .NumberOfContacts = 0
          .SystemVolume = 0
        End If
      
      Case FCDecelInToOut
        StateString = "FC: Decel In to Out " & TimerString(Timer.TimeRemaining)
        If Timer.Finished Then
          State = FCSwitchToOutToIn
          Timer = .Parameters.FlowReverseTime
          .NumberOfContacts = 0
          .SystemVolume = 0
        End If
      
      Case FCSwitchToOutToIn
        StateString = "FC: Switch Out to In " & TimerString(Timer.TimeRemaining)
        If Timer.Finished Then
          OutToIn = OutToIn + 1
          State = FCOutToIn
          .NumberOfContacts = 0
          .SystemVolume = 0
        End If

    End Select
  End With
End Sub

Friend Sub ACCommand_Cancel()
  State = FCOff
End Sub
Friend Property Get IsInToOut() As Long
  IsInToOut = (State = FCInToOut) Or (State = FCSwitchToInToOut) Or (State = FCDecelInToOut)
End Property
Friend Property Get IsOutToIn() As Long
  IsOutToIn = (State = FCOutToIn) Or (State = FCSwitchToOutToIn) Or (State = FCDecelOutToIn)
End Property
Friend Property Get IsDeceling() As Boolean
  IsDeceling = ((State = FCDecelInToOut) Or (State = FCDecelOutToIn))
End Property
Friend Property Get IsSwitching() As Boolean
  IsSwitching = (State = FCSwitchToInToOut) Or (State = FCSwitchToOutToIn)
End Property
Friend Property Get IsReversing() As Boolean
  IsReversing = IsDeceling Or IsSwitching
End Property
Friend Property Get IsOn() As Boolean
  IsOn = (State <> FCOff)
End Property
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property

#End If

End Class
