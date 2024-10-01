' A&E Mt Holly Then Platform moved to GMX
' Version 2022-04-19

Imports Utilities.Translations

Public Class LidControl : Inherits MarshalByRefObject
  Private ReadOnly ControlCode As ControlCode

  Public Enum EState
    Off
    ReleaseHoldDown
    ReleasePin
    OpenBand
    RaiseLid
    LidOpen
    LowerLid
    CloseBand
    EngagePin
  End Enum
  Property State As EState
  Property StateString As String
  Property LidClosed As Boolean
  Property LidOpen As Boolean

  ' Timers
  Property Timer As New Timer

  Sub New(ByVal controlCode As ControlCode)
    Me.ControlCode = controlCode
    Me.Timer.Seconds = 10
  End Sub

  Public Sub Run()
    With ControlCode
      LidClosed = .IO.KierLidClosed AndAlso .IO.LockingPinRaised

      Select Case State
        Case EState.Off
          StateString = ""
          ' Start opening lid
          If .IO.OpenLidPb AndAlso .MachineSafe AndAlso .LevelOkToOpenLid AndAlso (Not .IO.MainPump) AndAlso (Not .EStop) Then
            State = EState.ReleaseHoldDown
            Timer.Seconds = MinMax(.Parameters.CarrierHoldDownTime, 5, 30)
          End If
          ' Start lowering the lid
          If .IO.CloseLidPb AndAlso (Not .EStop) Then
            If .IO.KierLidClosed Then
              If Not .IO.LockingPinRaised Then
                State = EState.CloseBand
                Timer.Seconds = .Parameters.LockingBandClosingTime
              Else
                State = EState.EngagePin
              End If
            Else
              State = EState.LowerLid
            End If
          End If

        Case EState.ReleaseHoldDown
          If Timer.Finished Then
            State = EState.ReleasePin
            Timer.Seconds = 10
          End If
          StateString = Translate("Release Hold Down") & Timer.ToString(1)


        Case EState.ReleasePin
          If .IO.LockingPinRaised Then Timer.Seconds = 5
          If Timer.Finished Then
            State = EState.OpenBand
            Timer.Seconds = .Parameters.LockingBandOpeningTime
          End If
          StateString = Translate("Release Pin") & Timer.ToString(1)


        Case EState.OpenBand
          If Timer.Finished Then
            State = EState.RaiseLid
            Timer.Seconds = MinMax(.Parameters.LidRaisingTime, 5, 120)
          End If
          StateString = Translate("Open Band") & Timer.ToString(1)


        Case EState.RaiseLid
          If (Not .IO.OpenLidPb) Then
            State = EState.LidOpen
          Else
            Timer.Seconds = MinMax(.Parameters.LidRaisingTime, 5, 120)
          End If
          If Timer.Finished OrElse .IO.LidRaisedSw Then
            State = EState.LidOpen
          End If
          StateString = Translate("Raise Lid") & Timer.ToString(1)


        Case EState.LidOpen
          If .IO.OpenLidPb Then
            State = EState.RaiseLid
            Timer.Seconds = .Parameters.LidRaisingTime
          End If
          If .IO.CloseLidPb Then
            State = EState.LowerLid
          End If
          StateString = Translate("Lid Open")


        Case EState.LowerLid
          If (Not .IO.KierLidClosed) Then Timer.Seconds = 3
          If Not .IO.CloseLidPb Then
            State = EState.LidOpen
          End If
          If Timer.Finished Then
            State = EState.CloseBand
            Timer.Seconds = .Parameters.LockingBandClosingTime
          End If
          StateString = Translate("Lower Lid") & Timer.ToString(1)


        Case EState.CloseBand
          If Timer.Finished Then
            State = EState.EngagePin
          End If
          StateString = Translate("Close Band") & Timer.ToString(1)


        Case EState.EngagePin
          If (Not .IO.LockingPinRaised) Then Timer.Seconds = 3
          If Timer.Finished Then
            State = EState.Off
          End If
          StateString = Translate("Raise Pin") & Timer.ToString(1)

      End Select

    End With
  End Sub

  Public ReadOnly Property IsActive As Boolean
    Get
      Return (State <> EState.Off)
    End Get
  End Property

  Public ReadOnly Property IoReleasePin As Boolean
    Get
      Return (State >= EState.ReleasePin) AndAlso (State <= EState.CloseBand)
    End Get
  End Property

  Public ReadOnly Property IoLockingBandOpen As Boolean
    Get
      Return (State >= EState.OpenBand) AndAlso (State <= EState.LowerLid)
    End Get
  End Property

  Public ReadOnly Property IoRaiseLid As Boolean
    Get
      Return (State = EState.RaiseLid)
    End Get
  End Property

  Public ReadOnly Property IoLowerLid As Boolean
    Get
      Return (State = EState.LowerLid)
    End Get
  End Property

  Public ReadOnly Property IoLockingBandClose As Boolean
    Get
      Return (State = EState.Off) OrElse (State = EState.CloseBand) OrElse (State = EState.EngagePin)
    End Get
  End Property

  Public ReadOnly Property IoReleaseHoldDown As Boolean
    Get
      Return (State >= EState.ReleaseHoldDown) AndAlso (State <= EState.EngagePin)
    End Get
  End Property

End Class
