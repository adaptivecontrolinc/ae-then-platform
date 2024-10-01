'Version 2021-07-01
' Includes Translation Function
' Added functionality for mimic call to flush DP level transmitter and settle after.  
' - Only run if idle or load, unload procedures are active.

Imports Utilities.Translations
Public Class LevelFlush : Inherits MarshalByRefObject

  Public Enum EState
    Off
    Interlock
    LevelFlush
    LevelSettle
  End Enum
  Property State As EState
  Property StateString As String
  Property Level As Integer
  Property FlushCount As Integer

  Public Timer As New Timer

  Property DatePumpStartAuto() As Date
  Property DateFlushStartManual() As Date
  Private ReadOnly controlCode As ControlCode
  Public Sub New(ByVal controlCode As ControlCode)
    Me.controlCode = controlCode
  End Sub

  Public Sub Run()
    With controlCode

      ' Safe-To-Fill conditions
      Dim safe As Boolean = (.TempSafe OrElse .HD.HDCompleted) AndAlso .PressSafe

      ' Issue with inconsistent state value in histories - 
      Static StartState As EState
      Do
        ' Remember state and loop until state does not change
        ' This makes sure we go through all the state changes before proceeding and setting IO
        StartState = State

        ' Force state machine to interlock state if the machine is not safe
        If IsActive AndAlso Not safe Then State = EState.Interlock

        Select Case State
          Case EState.Off
            StateString = (" ")

          Case EState.Interlock
            StateString = Translate("Flush Level") & (" ") & Translate("Starting") & Timer.ToString(1)
            If .Parent.IsProgramRunning Then
              ' Only allow manual level flush during these commands if a program is running
              If Not (.LD.IsOn OrElse .UL.IsOn OrElse .DR.IsOn OrElse .FI.IsOn) Then
                StateString = Translate("Command Active") : Timer.Seconds = 1
              End If
            End If
            If Not .TempSafe Then StateString = Translate("Level FLush") & Translate("Interlocked") & Timer.ToString(1) : Timer.Seconds = 1
            If .EStop Then StateString = Translate("EStop") : Timer.Seconds = 1
            If Timer.Finished Then
              Timer.Seconds = MinMax(.Parameters.LevelGaugeFlushTime, 10, 60)
              State = EState.LevelFlush
            End If

          Case EState.LevelFlush
            StateString = Translate("Flush Level") & Timer.ToString(1)
            If Timer.Finished Then
              FlushCount += 1
              State = EState.LevelSettle
              Timer.Seconds = MinMax(.Parameters.LevelGaugeSettleTime, 10, 120)
            End If
            ' Lost Machine safe
            If Not .TempSafe Then State = EState.Interlock

          Case EState.LevelSettle
            StateString = Translate("Settling") & Timer.ToString(1)
            If Timer.Finished Then
              State = EState.Off
            End If

        End Select

      Loop Until (StartState = State)

    End With
  End Sub

  Public ReadOnly Property IsOff As Boolean
    Get
      Return (State = EState.Off)
    End Get
  End Property
  Public ReadOnly Property IsActive As Boolean
    Get
      Return (State >= EState.Interlock) AndAlso (State = EState.LevelSettle)
    End Get
  End Property
  Public ReadOnly Property IsFlushing As Boolean
    Get
      Return (State = EState.LevelFlush)
    End Get
  End Property

  Public Sub Cancel()
    State = EState.Off
    Timer.Cancel()
  End Sub
  Public Sub Reset()
    FlushCount = 0
  End Sub
  Public Sub LevelFlushStartManual()
    With controlCode
      'Disable manual start if no program is running and also during drain
      If (.Parent.IsProgramRunning) AndAlso Not (.LD.IsOn OrElse .UL.IsOn OrElse .DR.IsOn OrElse .FI.IsOn) Then Exit Sub

      ' Safe
      If Not (.TempSafe OrElse .HD.HDCompleted) AndAlso .PressSafe Then Exit Sub

      FlushCount += 1
      DateFlushStartManual = Date.UtcNow.ToLocalTime
      State = EState.Interlock

    End With
  End Sub
  Public Sub LevelFlushStartAuto()
    With controlCode
      ' Safe
      If Not (.TempSafe OrElse .HD.HDCompleted) AndAlso .PressSafe Then Exit Sub

      FlushCount += 1
      DatePumpStartAuto = Date.UtcNow.ToLocalTime
    End With
  End Sub

End Class