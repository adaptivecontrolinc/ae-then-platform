'American & Efird
' Version 2022-05-23

Imports Utilities.Translations

<Command("Kitchen Clean", "", "", "", "10"), 
TranslateCommand("es", "Esperar para la Cocina", ""), 
Description("Fill and transfers the kitchen tank to the drain."), 
TranslateDescription("es", "Relleno y transferencias la cocina tanque al dren."), 
Category("Kitchen Functions"), TranslateCategory("es", "Kitchen Functions")> 
Public Class CK : Inherits MarshalByRefObject : Implements ACCommand
  Private ReadOnly controlCode As ControlCode
  Private Const commandName_ As String = ("CK: ")

  ' Command States
  Public Enum EState
    Off
    WaitIdle
    FillToClean
    TransferDrain1
    RinseDrain
    TransferDrain2
  End Enum
  Public State As EState
  Public StateWas As EState
  Public Status As String
  Public NumberOfRinses As Integer

  Public Timer As New Timer
  Public TimerOverrun As New Timer

  Sub New(ByVal controlCode As ControlCode)
    Me.controlCode = controlCode
    Me.controlCode.Commands.Add(Me)
  End Sub

  Sub ParametersChanged(ByVal ParamArray param() As Integer) Implements ACCommand.ParametersChanged
    ' If the command is not running then do nothing the changes will be picked up when the command starts
    If Not IsOn Then Exit Sub

    ' Enable command parameter adjustments, if parameter set
    If controlCode.Parameters.EnableCommandChg = 1 Then Start(param)
  End Sub

  Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With controlCode

      ' Cancel Commands (including kitchen commands if we've jumped here before completing the kitchen prepare/transfer)
      If .KA.IsOn AndAlso Not .KA.IsBackgroundRinsing Then .KA.Cancel()
      If .KP.KP1.IsOn Then .KP.KP1.Cancel()
      If .LA.KP1.IsOn Then .LA.KP1.Cancel()
      .Tank1Ready = False
      .WK.Cancel()
      .DR.Cancel() : .FI.Cancel() : .HC.Cancel() : .HD.Cancel() : .RH.Cancel() : .RI.Cancel() : .TM.Cancel() ': .TU.Cancel()
      .LD.Cancel() : .PH.Cancel() : .SA.Cancel() : .UL.Cancel()
      .WT.Cancel()

      ' Command parameters

      ' Initial State
      State = EState.WaitIdle
      Timer.Seconds = 5
      TimerOverrun.Minutes = 10

    End With
  End Function

  Function Run() As Boolean Implements ACCommand.Run
    With controlCode

      Select Case State
        Case EState.Off
          Status = (" ")

        Case EState.WaitIdle
          Status = commandName_ & Translate("Wait Idle") & Timer.ToString(1)
          If .KP.KP1.IsForeground AndAlso .KP.KP1.Param_Destination = EKitchenDestination.Add Then
            Timer.Seconds = 2
            Status = commandName_ & Translate("Wait") & .KP.KP1.Status
          End If
          If .LA.KP1.IsForeground AndAlso .LA.KP1.Param_Destination = EKitchenDestination.Add Then
            Timer.Seconds = 2
            Status = commandName_ & Translate("Wait") & .LA.KP1.Status
          End If
          If .KA.IsActive AndAlso .KA.Destination = EKitchenDestination.Add Then
            Timer.Seconds = 2
            Status = commandName_ & Translate("Wait") & .KA.Status
          End If
          If Timer.Finished Then
            State = EState.FillToClean
            NumberOfRinses = 2
            'This is a background command, now we can step on
            Return True
          End If

        Case EState.FillToClean
          If (.Tank1Level >= (.Parameters.CKFillLevel - .Parameters.Tank1FillLevelDeadBand)) Then
            State = EState.TransferDrain1
          End If
          Status = commandName_ & Translate("Filling") & (" ") & (.Tank1Level / 10).ToString("#0.0") & "%"


        Case EState.TransferDrain1
          If (.Tank1Level > 10) Then
            Status = commandName_ & Translate("Draining") & (" ") & (.Tank1Level / 10).ToString("#0.0") & "%"
            Timer.Seconds = MinMax(.Parameters.CKTransferToDrainTime1, 10, 600)
          Else
            Status = commandName_ & Translate("Draining") & Timer.ToString(1)
          End If
          If Timer.Finished Then
            State = EState.RinseDrain
            Timer.Seconds = MinMax(.Parameters.CKRinseTime, 1, 60)
          End If


        Case EState.RinseDrain
          If Timer.Finished OrElse (.Tank1Level >= (.Parameters.CKFillLevel - .Parameters.Tank1FillLevelDeadBand)) Then
            NumberOfRinses = -1
            State = EState.TransferDrain2
            Timer.Seconds = MinMax(.Parameters.CKTransferToDrainTime2, 10, 600)
          End If
          Status = commandName_ & Translate("Filling") & (" ") & (.Tank1Level / 10).ToString("#0.0") & "%" & (" (") & NumberOfRinses.ToString("#0") & (")")


        Case EState.TransferDrain2
          Status = commandName_ & Translate("Draining")
          If (.Tank1Level > 10) Then
            Status = Status & (" ") & (.Tank1Level / 10).ToString("#0.0") & "%"
            Timer.Seconds = MinMax(.Parameters.CKTransferToDrainTime2, 10, 600)
          Else
            Status = commandName_ & Timer.ToString(1)
          End If
          If Timer.Finished Then
            If NumberOfRinses > 0 Then
              State = EState.RinseDrain
              Timer.Seconds = MinMax(.Parameters.CKRinseTime, 1, 60)
            Else
              State = EState.Off
            End If
          End If

      End Select
    End With
  End Function

  Public Sub Cancel() Implements ACCommand.Cancel
    Status = (" ")
    State = EState.Off
    Timer.Cancel()
    TimerOverrun.Cancel()
  End Sub

  ReadOnly Property IsOn As Boolean Implements ACCommand.IsOn
    Get
      IsOn = State <> EState.Off
    End Get
  End Property

  ReadOnly Property IsOverrun As Boolean
    Get
      Return IsOn AndAlso TimerOverrun.Finished
    End Get
  End Property

  ReadOnly Property IoTank1Fill As Boolean
    Get
      Return (State = EState.FillToClean) OrElse (State = EState.RinseDrain)
    End Get
  End Property

  ReadOnly Property IoTank1Transfer As Boolean
    Get
      Return (State = EState.TransferDrain1) OrElse (State = EState.TransferDrain2)
    End Get
  End Property

End Class
