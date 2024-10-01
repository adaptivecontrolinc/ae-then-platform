'American & Efird - MX
' Version 2016-05-26

Imports Utilities.Translations

<Command("Flow Reverse Time", "I:|0-99| O:|0-99|", "", "", ""), _
TranslateCommand("es", "Flow Reverse Time", "I:|0-99| O:|0-99|"), _
Description("Sets the time, in minutes, between flow reversals. I is minutes on In to Out flow, O is minutes on Out to In flow."), _
TranslateDescription("es", "Set flow reversal time."), _
Category("Flow Control"), TranslateCategory("es", "Flow Control")> _
Public Class FR
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
      .FC.Cancel()

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

#If 0 Then
' TODO

Attribute VB_Name = "FR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Flow reverse command
Option Explicit
Implements ACCommand
Public Enum FRState
  FROff
  FRPauseOutToIn
  FROutToIn
  FRPauseInToOut
  FRInToOut
  FRDecelInToOut
  FRDecelOutToIn
  FRSwitchToInToOut
  FRSwitchToOutToIn
End Enum
Public State As FRState
Public StateString As String
Public InToOutTime As Long, OutToInTime As Long
Public Timer As New acTimer
Public InToOut As Long, OutToIn As Long

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Name=Flow Reverse\r\nParameters=I:|0-99| O:|0-99|\r\nHelp=Sets the time, in minutes, between flow reversals. I is minutes on In to Out flow, O is minutes on Out to In flow. "
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode
    .FC.ACCommand_Cancel
    InToOutTime = Param(1) * 60
    OutToInTime = Param(2) * 60
  End With
 
  If InToOutTime = 0 Then
    If IsInToOut Then OutToIn = OutToIn + 1
     State = FROutToIn
  ElseIf OutToInTime = 0 Then
    If IsOutToIn Then InToOut = InToOut + 1
     State = FRInToOut
  Else
    If IsInToOut Then
      OutToIn = OutToIn + 1
    Else: InToOut = InToOut + 1
    End If
    State = FRInToOut
  End If
  Timer = InToOutTime
  StepOn = True
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode
  
    Static PumpWasRunning As Boolean
    If .IO_MainPumpRunning Then
      If Not PumpWasRunning Then Timer.Restart
    Else
      Timer.Pause
    End If
    PumpWasRunning = .IO_MainPumpRunning

    Select Case State
      Case FROff
        StateString = ""
        
      Case FRPauseOutToIn
        StateString = "FR: Paused Out to In "
        If .IO_MainPumpRunning Then
          Timer.Restart
          State = FROutToIn
        End If
        
      Case FRPauseInToOut
        StateString = "FR: Paused In to Out "
        If .IO_MainPumpRunning Then
          Timer.Restart
          State = FRInToOut
        End If
        
      Case FROutToIn
        StateString = "FR: Out to In " & TimerString(Timer.TimeRemaining)
        If Not .IO_MainPumpRunning Then
          State = FRPauseOutToIn
          Timer.Pause
        End If
        If Timer.Finished And (InToOutTime > 0) Then
          State = FRDecelOutToIn
          Timer = MiniMaxi(.Parameters.PumpSpeedDecelTime, 15, 40)
        End If
      
      Case FRDecelOutToIn
        StateString = "FR: Decel Out to In " & TimerString(Timer.TimeRemaining)
        If Timer.Finished Then
          State = FRSwitchToInToOut
          Timer = .Parameters.FlowReverseTime
        End If
     
      Case FRSwitchToInToOut
        StateString = "FR: Switch In to Out " & TimerString(Timer.TimeRemaining)
        If Timer.Finished Then
          InToOut = InToOut + 1
          State = FRInToOut
          Timer = InToOutTime
        End If
     
      Case FRInToOut
        StateString = "FR: In to Out " & TimerString(Timer.TimeRemaining)
        If Not .IO_MainPumpRunning Then
          State = FRPauseInToOut
          Timer.Pause
        End If
        If Timer.Finished And (OutToInTime > 0) Then
          State = FRDecelInToOut
          Timer = MiniMaxi(.Parameters.PumpSpeedDecelTime, 15, 40)
        End If
      
      Case FRDecelInToOut
        StateString = "FR: Decel In to Out " & TimerString(Timer.TimeRemaining)
        If Timer.Finished Then
          State = FRSwitchToOutToIn
          Timer = .Parameters.FlowReverseTime
        End If
      
      Case FRSwitchToOutToIn
        StateString = "FR: Switch Out to In " & TimerString(Timer.TimeRemaining)
        If Timer.Finished Then
          OutToIn = OutToIn + 1
          State = FROutToIn
          Timer = OutToInTime
        End If

    End Select
  End With
End Sub

Friend Sub ACCommand_Cancel()
  State = FROff
End Sub
Friend Property Get IsInToOut() As Long
  IsInToOut = (State = FRInToOut) Or (State = FRPauseInToOut) Or (State = FRSwitchToInToOut) Or (State = FRDecelInToOut)
End Property
Friend Property Get IsOutToIn() As Long
  IsOutToIn = (State = FROutToIn) Or (State = FRPauseOutToIn) Or (State = FRSwitchToOutToIn) Or (State = FRDecelOutToIn)
End Property
Friend Property Get IsDeceling() As Boolean
  IsDeceling = ((State = FRDecelInToOut) Or (State = FRDecelOutToIn))
End Property
Friend Property Get IsSwitching() As Boolean
  IsSwitching = (State = FRSwitchToInToOut) Or (State = FRSwitchToOutToIn)
End Property
Friend Property Get IsReversing() As Boolean
'Property used for Add-By-Contacts to prevent transferring while switching;
' also prevents error in desired level based on number of contacts
  IsReversing = (State = FRDecelInToOut) Or (State = FRSwitchToInToOut) Or (State = FRDecelOutToIn) Or (State = FRSwitchToOutToIn)
End Property
Friend Property Get IsOn() As Boolean
  IsOn = (State <> FROff)
End Property
Friend Sub RestartInToOut()
  If IsOn Then
    State = FRInToOut
    Timer = InToOutTime
  End If
End Sub
Friend Sub RestartOutToIn()
  If IsOn Then
    State = FROutToIn
    Timer = OutToInTime
  End If
End Sub
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property


#End If

End Class
