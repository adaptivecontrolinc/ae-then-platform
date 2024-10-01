'American & Efird
' Version 2024-09-16

Imports Utilities.Translations

<Command("Differential Pressure", "I:|0-99|.|0-9| Psi O:|0-99|.|0-9|Psi", "", "", "", CommandType.Standard),
  TranslateCommand("es", "Presión Diferenciada", "I:|0-99|.|0-9| Psi O:|0-99|.|0-9|Psi"),
  Description("Sets the desired differential pressure for InsideOut (I) and for OutsideIn (O)."),
  TranslateDescription("es", "Fija la presión diferenciada deseada."),
  Category("Flow Control"), TranslateCategory("es", "Flow Control")>
Public Class DP : Inherits MarshalByRefObject : Implements ACCommand
  Private ReadOnly controlCode As ControlCode

  Public Enum EState
    Off
    Active
  End Enum
  Property State As EState
  Property Status As String
  Property InsideOutPress As Integer
  Property OutsideInPress As Integer

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
      If param.GetUpperBound(0) >= 2 Then InsideOutPress = (param(1) * 10) + param(2)
      If param.GetUpperBound(0) >= 4 Then OutsideInPress = (param(3) * 10) + param(4)

      ' Limit Setpoints
      If InsideOutPress > controlCode.Parameters.PackageDiffPressRange Then InsideOutPress = controlCode.Parameters.PackageDiffPressRange
      If OutsideInPress > controlCode.Parameters.PackageDiffPressRange Then OutsideInPress = controlCode.Parameters.PackageDiffPressRange

    End If
  End Sub

  Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With controlCode
      '.FL.Cancel()
      .FP.Cancel()

      'Check array bounds just to be on the safe side
      If param.GetUpperBound(0) >= 2 Then InsideOutPress = (param(1) * 10) + param(2) ' in tenths psi
      If param.GetUpperBound(0) >= 4 Then OutsideInPress = (param(3) * 10) + param(4) ' in tenths psi

      ' Limit Setpoints
      If InsideOutPress > .Parameters.PackageDiffPressRange Then InsideOutPress = .Parameters.PackageDiffPressRange
      If OutsideInPress > .Parameters.PackageDiffPressRange Then OutsideInPress = .Parameters.PackageDiffPressRange

      'Set default state
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
          If .PumpControl.IsOutToIn Then
            Status = Translate("Desired Pressure") & (" ") & (OutsideInPress / 10).ToString("#0.0") & "PSI"
          Else
            Status = Translate("Desired Pressure") & (" ") & (InsideOutPress / 10).ToString("#0.0") & "PSI"
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

End Class
