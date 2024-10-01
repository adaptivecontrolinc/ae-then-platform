
' Version 2024-07-30
'American & Efird 

<Command("Batch Weight", "|0-9999| Lbs", " ", "", "", CommandType.BatchParameter),
TranslateCommand("es", "Peso del Lote", "|0-9999| kgs"),
Description("Sets batch weight. The batch weight is used in conjunction with the Liquor Ratio (LR command) to calculate the working volume."),
TranslateDescription("es", "Fija el peso de la hornada."),
Category("Batch Commands"), TranslateCategory("es", "Batch Commands")>
Public Class BW
  Inherits MarshalByRefObject
  Implements ACCommand

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
    'If the command is not running then do nothing the changes will be picked up when the command starts
    If Not IsOn Then Exit Sub
  End Sub

  Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With controlCode

      'Check array bounds just to be on the safe side
      If param.GetUpperBound(0) >= 1 Then .BatchWeight = param(1)

      Cancel()
    End With
  End Function

  Function Run() As Boolean Implements ACCommand.Run
  End Function

  Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
  End Sub

  Public ReadOnly Property IsOn() As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property

End Class
