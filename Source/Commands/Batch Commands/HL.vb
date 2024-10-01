'American & Efird - Mt. Holly Then Platform to GMX
' Version 2022-02-03

<Command("Half Load", "|0-1|", " ", "", "", CommandType.BatchParameter), _
TranslateCommand("es", "Media carga", "|0-1|"), _
Description("Set to '1' to run only one Kier, with second kier isolated. Default value '0' runs both kiers."), _
TranslateDescription("es", "Establezca en '1' para ejecutar solo Kier, con kier segundo aislado. El valor por defecto '0' ejecuta tanto kiers."), _
Category("Batch Commands"), TranslateCategory("es", "Batch Commands")> _
Public Class HL
  Inherits MarshalByRefObject
  Implements ACCommand

  'Command States
  Public Enum EState
    Off
    Active
  End Enum
  Property State As EState
  Property KierBLoaded As Boolean
  Property NumberSpindles As Integer
  Public PackageHeightCalculated As Integer

  Private Property timer_ As New Timer

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
      If param.GetUpperBound(0) >= 1 Then KierBLoaded = (param(1) = 0)

      If Not KierBLoaded Then
        State = EState.Active
      Else
        State = EState.Off
      End If

      ' TODO - check for PT & NP and calculate working level

      ' Calculate Package Height
      If (.Parameters.NumberOfSpindales > 0) Then
        ' Capacity to Isolate one kier
        If .HL.KierBLoaded Then
          NumberSpindles = .Parameters.NumberOfSpindales * 2
        Else
          NumberSpindles = .Parameters.NumberOfSpindales
        End If

        ' Calculate package height
        PackageHeightCalculated = CInt(.NP.NumberPackages / NumberSpindles)
        If PackageHeightCalculated > 4 Then
          .PackageHeight = 4
        ElseIf PackageHeightCalculated > 3 Then
          .PackageHeight = 3
        ElseIf PackageHeightCalculated > 2 Then
          .PackageHeight = 2
        ElseIf PackageHeightCalculated > 1 Then
          .PackageHeight = 1
        Else
          .PackageHeight = 0
        End If
      End If

      ' Determine working level based on NP, PT, & Corresponding Level Parameters
      If (.PackageHeight > 0) AndAlso (.PackageHeight <= 4) Then
        Select Case .PackageType
          Case 0
            .WorkingLevel = .Parameters.PackageType0Level(.PackageHeight)
          Case 1
            .WorkingLevel = .Parameters.PackageType1Level(.PackageHeight)
          Case 2
            .WorkingLevel = .Parameters.PackageType2Level(.PackageHeight)
          Case 3
            .WorkingLevel = .Parameters.PackageType3Level(.PackageHeight)
          Case 4
            .WorkingLevel = .Parameters.PackageType4Level(.PackageHeight)
            '  Case 5
            '    .WorkingLevel = .Parameters.PackageType5Level(.PackageHeight)
            '  Case 6
            '    .WorkingLevel = .Parameters.PackageType6Level(.PackageHeight)
            '  Case 7
            '    .WorkingLevel = .Parameters.PackageType7Level(.PackageHeight)
            '  Case 8
            '    .WorkingLevel = .Parameters.PackageType8Level(.PackageHeight)
        End Select
      End If

      'Foreground Step for moment
      Return False
    End With
  End Function

  Function Run() As Boolean Implements ACCommand.Run
    With controlCode

      Select Case State
        Case EState.Off

        Case EState.Active
          Return True
          If Not .Parent.IsProgramRunning Then

            State = EState.Off
          End If

      End Select
    End With
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
