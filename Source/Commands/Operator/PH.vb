'American & Efird - Mt. Holly Then Platform to GMX
' Version 2022-02-03

Imports Utilities.Translations

<Command("pH Check", "Low:|0-14|.|0-9| High:|0-14|.|0-9|", "", "", "('StandardTimeCheckPH)=5"), _
TranslateCommand("es", "Chequeo  pH", "Bajo:|0-14|.|0-9| Alto:|0-14|.|0-9|"), _
Description("Signals the operator to check the ph."), _
TranslateDescription("es", "Señala a operador para comprobar el pH."), _
Category("Operator Functions"), TranslateCategory("es", "Operator Functions")> _
Public Class PH
  Inherits MarshalByRefObject
  Implements ACCommand
  Private Const commandName_ As String = ("PH: ")

  'Command States
  Public Enum EState
    Off
    Interlock
    ManualCheck
    Done
  End Enum
  Property State As EState
  Property StateString As String
  Property LowLimit As Integer
  Property HighLimit As Integer
  Private flashSlow_ As Boolean
  Private flashFast_ As Boolean

  Property Timer As New Timer
  Property TimerOverrun As New Timer

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

      ' Cancel Commands
      .AC.Cancel()
      If .AT.IsForeground Then .AT.Cancel()
      .KA.Cancel() : .WK.Cancel()
      .DR.Cancel() : .FI.Cancel() : .HC.Cancel() : .HD.Cancel()
      .RI.Cancel() : .RC.Cancel() : .RH.Cancel() : .TM.Cancel()
      .LD.Cancel() ': .PH.Cancel() 
      .SA.Cancel() : .UL.Cancel()
      If .RF.IsForeground AndAlso .RF.FillType = EFillType.Vessel Then .RF.Cancel() ' TODO Check
      If .RT.IsForeground Then .RT.Cancel()
      .RW.Cancel()
      .WT.Cancel()

      ' Cancel PressureOn command if active during Atmospheric Commands
      .PR.Cancel()

      ' Clear Parent.Signal 
      If .Parent.Signal <> "" Then .Parent.Signal = ""

      'Check Parameters
      If param.GetUpperBound(0) >= 1 Then LowLimit = (param(1) * 10)
      If param.GetUpperBound(0) >= 2 Then LowLimit += param(2)

      If param.GetUpperBound(0) >= 3 Then HighLimit = (param(3) * 10)
      If param.GetUpperBound(0) >= 4 Then HighLimit += param(4)

      'Set Default State
      TimerOverrun.Minutes = .Parameters.StandardTimeCheckPH
      State = EState.Interlock
      Timer.Seconds = 2
    End With
  End Function

  Function Run() As Boolean Implements ACCommand.Run
    With controlCode

      'Reset Advance Timer
      If (Not .AdvancePb) OrElse (.Parent.Signal <> "") Then Timer.Seconds = 2

      ' Flashers in sync
      flashSlow_ = .FlashSlow
      flashFast_ = .FlashFast

      Select Case State

        Case EState.Off
          StateString = (" ")

        Case EState.Interlock
          StateString = Translate("Wait Safe")
          If .TempSafe Then
            State = EState.ManualCheck
            .Parent.Signal = Translate("Check pH") & (" ") & (LowLimit / 10).ToString("#0.0") & " to " & (HighLimit / 10).ToString("#0.0")
          End If

        Case EState.ManualCheck
          StateString = Translate("Check pH") & (" ") & (LowLimit / 10).ToString("#0.0") & " to " & (HighLimit / 10).ToString("#0.0") & (", ") & Translate("Hold Advance")
          If .AdvancePb Then
            If .Parent.Signal <> "" Then .Parent.Signal = ""
            ' Run Held
            StateString = Translate("Hold Advance") & (" ") & Timer.ToString
            If Timer.Finished Then
              If .Parent.Signal <> "" Then .Parent.Signal = ""
              State = EState.Done
              .IO.AdvancePb = False
              Timer.Seconds = 2
            End If
          End If
          If Not .TempSafe Then State = EState.Interlock

        Case EState.Done
          StateString = Translate("Completing") & Timer.ToString(1)
          If Timer.Finished Then Cancel()

      End Select
    End With
  End Function

  Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
    StateString = ""
    Timer.Cancel()
    TimerOverrun.Cancel()
  End Sub

  ReadOnly Property IsOn As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property

  ReadOnly Property IsOverrun As Boolean
    Get
      Return IsOn AndAlso TimerOverrun.Finished
    End Get
  End Property

  ReadOnly Property IoLampSignal As Boolean
    Get
      Return (State = EState.ManualCheck) AndAlso flashSlow_
    End Get
  End Property

End Class
