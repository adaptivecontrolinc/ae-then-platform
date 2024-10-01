Imports Utilities.Translations

Partial Class ControlCode

  ReadOnly Property Status() As String
    Get
      'If Parent.Signal <> "" Then Return Parent.Signal

      If IO.EmergencyStop Then Return "EStop Pressed"
      If Alarms.LidNotLocked Then Return "Lid Not Locked: Lock and Hold Run" & LidControl.StateString
      'Status = "Lid Not Locked: Lock and Hold Run " & TimerString(LidLockTimer.TimeRemaining)

      If Not String.IsNullOrEmpty(Parent.Signal) Then
        If IsSignalEnabled Then
          If Parent.AckState = AckStateValue.AckMessage Then Return Parent.Signal
          If Parent.AckState = AckStateValue.UnackMessage Then Return Parent.Signal
        Else
          Parent.Signal = ""
        End If
      End If

      If Not Parent.IsProgramRunning Then Return "Machine Idle: " & ProgramStoppedTimer.ToString


      'ElseIf (AP.IsSlow Or AP.IsFast) And SlowFlash Then return "Local Add Addition Tank: Prepare"
      If AC.IsForeground AndAlso AC.Status <> "" Then Return AC.Status
      If AD.IsForeground AndAlso AD.StateString <> "" Then Return AD.StateString
      If AF.IsActive AndAlso (AF.FillType = EFillType.Vessel) AndAlso AF.Status <> "" Then Return AF.Status
      If AT.IsForeground AndAlso AT.Status <> "" Then Return AT.Status

      If KA.IsOn AndAlso (Not KA.IsBackgroundRinsing) AndAlso KA.Status <> "" Then Return KA.Status
      If WK.IsOn AndAlso WK.Status <> "" Then Return WK.Status

      If DR.IsOn AndAlso DR.Status <> "" Then Return DR.Status
      If FI.IsOn AndAlso FI.Status <> "" Then Return FI.Status
      If HC.IsOn AndAlso HC.Status <> "" Then Return HC.Status
      If HD.IsOn AndAlso HD.Status <> "" Then Return HD.Status
      If PR.IsInterlock AndAlso PR.StateString <> "" Then Return PR.StateString
      If RC.IsOn AndAlso RC.Status <> "" Then Return RC.Status
      If RH.IsOn AndAlso RH.Status <> "" Then Return RH.Status
      If RI.IsOn AndAlso RI.Status <> "" Then Return RI.Status
      ' If TC.IsOn andalso TC.StateString <> "" Then Return TC.StateString 'TODO
      If TM.IsOn AndAlso TM.Status <> "" Then Return TM.Status

      If LD.IsOn AndAlso LD.Status <> "" Then Return LD.Status
      If PH.IsOn AndAlso PH.StateString <> "" Then Return PH.StateString
      If SA.IsOn AndAlso SA.StateString <> "" Then Return SA.StateString
      If UL.IsOn AndAlso UL.StateString <> "" Then Return UL.StateString


      If RD.IsWaitIdle AndAlso RD.Status <> "" Then Return RT.Status
      If RF.IsForeground AndAlso RF.Status <> "" Then Return RF.Status
      If RH.IsActive AndAlso RH.Status <> "" Then Return RH.Status
      If RT.IsForeground AndAlso RT.Status <> "" Then Return RT.Status
      If RW.IsActive AndAlso RW.Status <> "" Then Return RW.Status
      If RP.IsWaitReady AndAlso RP.StateString <> "" AndAlso (Not KP.KP1.IoToReserveTank) AndAlso FlashSlow Then Return RP.StateString
      '  If (RP.IsSlow Or RP.IsFast) And SlowFlash Then Status = "Local Add Reserve Tank: Prepare"

      If CO.IsActive AndAlso CO.StateString <> "" Then Return CO.StateString
      If HC.IsOn AndAlso HE.StateString <> "" Then Return HE.StateString
      If HE.IsActive AndAlso HE.StateString <> "" Then Return HE.StateString
      If WT.IsOn AndAlso WT.Status <> "" Then Return WT.Status


      ' Default Return Value
      Return ""
    End Get
  End Property

  ReadOnly Property IsSignalEnabled As Boolean
    Get
      ' Operator Calls
      If LD.IsOn OrElse PH.IsOn OrElse SA.IsOn OrElse UL.IsOn OrElse TM.IsOn OrElse CO.IsOn OrElse HE.IsOn OrElse TP.IsOn Then Return True
      Return False
    End Get
  End Property

End Class
