'American & Efird - MX Package
' Version 2024-07-30

<Command("BoilOut", "", "", "", ""), _
  TranslateCommand("es", "BoilOut", ""), _
  Description("Logs all program time to Boilout Delay."), _
  TranslateDescription("es", "Registra toda la hora del  programa a Boilout retrasa."),
  Category("Production Reports"), TranslateCategory("es", "Production Reports")>
Public Class BO : Inherits MarshalByRefObject : Implements ACCommand

  Public Enum EState
    Off
    BoilOut
  End Enum
  Public State As EState

  Private ReadOnly ControlCode As ControlCode
  Sub New(ByVal controlCode As ControlCode)
    Me.ControlCode = controlCode
  End Sub
  Sub ParametersChanged(ByVal ParamArray param() As Integer) Implements ACCommand.ParametersChanged
    ' Do nothing
  End Sub
  Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With ControlCode
      .LD.Cancel() : .UL.Cancel()
      State = EState.BoilOut
      Return True
    End With
  End Function

  Function Run() As Boolean Implements ACCommand.Run
  End Function
  Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
  End Sub

  ReadOnly Property IsOn() As Boolean Implements ACCommand.IsOn
    Get
      IsOn = (State <> EState.Off)
    End Get
  End Property

End Class
