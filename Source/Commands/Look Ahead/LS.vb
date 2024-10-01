'American & Efird - MX Package Machines
' Version 2024-07-29

Imports Utilities.Translations

<Command("Look Ahead Stop", "", "", "", "", CommandType.ParallelCommand), _
Description("The look ahead wont look further than this in the procedure."),
Category("Look Ahead Functions")>
Public Class LS : Inherits MarshalByRefObject : Implements ACCommand
  Private ReadOnly ControlCode As ControlCode

  Public Enum EState
    Off
    Active
  End Enum
  Property State As EState
  Property StateString As String

  Sub New(ByVal controlCode As ControlCode)
    Me.ControlCode = controlCode
    Me.ControlCode.Commands.Add(Me)
  End Sub

  Public Sub ParametersChanged(ParamArray param() As Integer) Implements ACCommand.ParametersChanged
    'If the command is not running then do nothing the changes will be picked up when the command starts
    If Not IsOn Then Exit Sub
  End Sub

  Public Function Start(ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With ControlCode
      With ControlCode
        ' Just so we log the command state transitions
        If State = EState.Off Then
          State = EState.Active
        Else
          State = EState.Off
        End If
        Return True
      End With
    End With
  End Function

  Public Function Run() As Boolean Implements ACCommand.Run
    With ControlCode
      Select Case State
        Case EState.Off
          StateString = (" ")

        Case EState.Active
          StateString = Translate("Look Ahead Stopped")
          If Not .Parent.IsProgramRunning Then Cancel()

      End Select
    End With
  End Function

  Public Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
  End Sub

  Public ReadOnly Property IsOn As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property

End Class
