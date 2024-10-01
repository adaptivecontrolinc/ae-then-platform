
'American & Efird
' Version 2024-07-30

<Command("Begin Reprocess", "", "", "", ""), _
  TranslateCommand("es", "Comience Tratan de nuevo", ""), _
  Description("Logs all program time to Corrective Add delay."), _
  TranslateDescription("es", "El medir el tiempo del comienzo trata de nuevo (correctivo agregue). Se registra el tiempo como correctivo agrega retrasa."), _
  Category("Production Reports"), TranslateCategory("es", "Production Reports")> _
Public Class BR
  Inherits MarshalByRefObject
  Implements ACCommand

  'Command States
  Public Enum EState
    Off
    BeginReprocess
  End Enum
  Public State As EState

  Private ReadOnly ControlCode As ControlCode
  Sub New(ByVal controlCode As ControlCode)
    Me.ControlCode = controlCode
  End Sub

  Sub ParametersChanged(ByVal ParamArray param() As Integer) Implements ACCommand.ParametersChanged
  End Sub

  Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With ControlCode

      ' TODO - VB6 only cancelled?: LD,SA,UL
      .LD.Cancel()
      .PH.Cancel()
      .SA.Cancel()
      .UL.Cancel()
      .BO.Cancel()

      State = EState.BeginReprocess
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
      IsOn = State <> EState.Off
    End Get
  End Property

End Class
