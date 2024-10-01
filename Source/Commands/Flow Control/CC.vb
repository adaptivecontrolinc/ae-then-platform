'American & Efird - MX Then Multiflex Package
' Version 2024-07-30

<Command("Circulation", "Start: |0-1|", " ", "", ""), _
  Description("Start (1) or stop (0) circulation."), _
  Category("Flow Control")> _
Public Class CC : Inherits MarshalByRefObject : Implements ACCommand

  'Command States
  Public Enum EState
    Off
    Active
  End Enum
  Public State As EState

  'Keep a local reference to the control code object for convenience
  Private ReadOnly controlCode As ControlCode

  Public Sub New(ByVal controlCode As ControlCode)
    Me.controlCode = controlCode
    Me.controlCode.Commands.Add(Me)
  End Sub
  Public Sub ParametersChanged(ByVal ParamArray param() As Integer) Implements ACCommand.ParametersChanged
    If Not IsOn Then Exit Sub
    Start(param)
  End Sub
  Public Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With controlCode
      'Command parameters
      Dim circulate As Integer

      'Check array bounds just to be on the safe side
      If param.GetUpperBound(0) >= 1 Then circulate = param(1)

      Select Case circulate
        Case 1
          'Start Circulation and Cancel this command
          .PumpControl.StartAuto()
          Cancel()
        Case Else
          'Stop Circulation, but remain active - used with TM command to soak carrier
          .PumpControl.Cancel()
          State = EState.Active
      End Select

      'This is a background command, step on
      Return True
    End With
  End Function

  Public Function Run() As Boolean Implements ACCommand.Run
    With controlCode
      Select Case State
        Case EState.Off
        Case EState.Active
          'Circulation stopped, if it starts back remotely, then cancel this command
          If .PumpControl.IsStarting OrElse .PumpControl.IsRunning Then Cancel()
      End Select
    End With
  End Function

  Public Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
  End Sub

  Public ReadOnly Property IsOn() As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property

End Class
