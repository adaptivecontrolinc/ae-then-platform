'American & Efird - 
' Version 2024-09-17

Imports Utilities.Translations

<Command("Add Fill", "|HCMV| |0-99|% Mix?|0-1|", "", "", ""), _
TranslateCommand("es", "Agregar Relleno", "|HCMV| |0-99|% Mix?|0-1|"), _
Description("Fills the add tank to the desired level (in percent) from the desired fill type: H=Hot, C=Cold, M=Mix,V=Vessel.  Also enables mix circulation once level is sufficient."), _
TranslateDescription("es", "Llene el tanque de adición a nivel desried."),
Category("Add Tank Commands"), TranslateCategory("es", "Add Tank Commands")>
Public Class AF : Inherits MarshalByRefObject : Implements ACCommand
  Private Const commandName_ As String = ("AF: ")

  Public Enum EState
    Off
    WaitIdle
    FillStart
    FillActive
    Done
  End Enum
  Public State As EState
  Public Status As String
  Public FillLevel As Integer
  Public FillType As EFillType
  Public FillMixOn As EMixState

  Property Timer As New Timer
  Friend Property TimerOverrun As New Timer
  Public ReadOnly Property TimeOverrun As Integer
    Get
      Return TimerOverrun.Seconds
    End Get
  End Property

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

  Public Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With controlCode

      'Command parameters - check array bounds just to be on the safe side

      ' Fill Type [H=72, C=67, M=77, V=86]
      Dim intFillType As Integer = 0
      If param.GetUpperBound(0) >= 1 Then intFillType = param(1)
      If param.GetUpperBound(0) >= 2 Then FillLevel = param(2) * 10

      Dim intFillMixerOn As Integer = 0
      If param.GetUpperBound(0) >= 3 Then intFillMixerOn = param(3)
      If intFillMixerOn > 0 Then FillMixOn = EMixState.Active

      ' Set the default fill level 
      FillLevel = MinMax(FillLevel, 0, 1000)

      ' Set the Fill Type (H=72, C=67, M=77, V=86)
      Select Case intFillType
        Case 86
          FillType = EFillType.Vessel
        Case 77
          FillType = EFillType.Mix
        Case 72
          FillType = EFillType.Hot
        Case Else
          FillType = EFillType.Cold
      End Select
      ' Restrict the Fill Type [2016-08-31]
      'If (.Parameters.FillEnableBlend = 0) AndAlso FillType = EFillType.Mix Then FillType = EFillType.Hot
      'If (.Parameters.FillEnableHot = 0) AndAlso FillType = EFillType.Hot Then FillType = EFillType.Cold

      ' Set the default State
      State = EState.WaitIdle
      Timer.Seconds = 1

      ' TODO - maybe use AddControl.IsOverrun
      TimerOverrun.Minutes = 5

      ' This is a background command so tell the control system to step on
      Return True
    End With
  End Function

  Public Function Run() As Boolean Implements ACCommand.Run
    With controlCode

      Select Case State

        Case EState.Off
          Status = (" ")


        Case EState.WaitIdle
          If .KP.KP1.Param_Destination = EKitchenDestination.Add Then
            If .KP.KP1.IsForeground Then Timer.Seconds = 2
            TimerOverrun.Seconds = 300
            Status = commandName_ & .KP.KP1.DrugroomDisplay
          ElseIf .KA.Destination = EKitchenDestination.Add Then
            If .KA.IsForeground Then Timer.Seconds = 2
            TimerOverrun.Seconds = 300
            Status = commandName_ & "Tank 1 " & .KA.Status
          Else
            If FillType = EFillType.Vessel Then
              If Not .TempSafe Then
                Timer.Seconds = 2
                Status = commandName_ & "Runback Wait Temp Safe"
              Else
                Status = commandName_ & "Runback Starting" & Timer.ToString(1)
              End If
            Else
              Status = commandName_ & "Starting" & Timer.ToString(1)
            End If
          End If
          If Timer.Finished Then
            ' Call the Add Control subroutine
            .AddControl.FillAuto(FillType, FillLevel, FillMixOn)

            State = EState.FillStart
          End If
          'StateString = Translate("Wait Idle") & Timer.ToString(1)
          'If .KP.KP1.Param_Destination = EKitchenDestination.Add Then StateString = commandName_ & "Wait: " & .KP.KP1.Status
          'If .KA.Destination = EKitchenDestination.Add Then StateString = commandName_ & "Wait: " & .KA.StateString
          'If FillType = EFillType.Vessel AndAlso Not .TempSafe Then StateString = commandName_ & "Runback Wait Temp Safe"


        Case EState.FillStart
          If .AddControl.FillIsActive Then
            State = EState.FillActive
            Timer.Seconds = 2
          End If
          Status = .AddControl.Status


        Case EState.FillActive
          If .AddControl.FillIsActive AndAlso .AddControl.FillIsForeground Then
            Timer.Seconds = 1
          End If
          If Timer.Finished Then
            State = EState.Done
            Timer.Seconds = 2
          End If
          Status = commandName_ & .AddControl.Status


        Case EState.Done
          If Timer.Finished Then
            ' Record Level once airpad has equalized after complete delay
            .GetRecordedLevel = True
            State = EState.Off
          End If
          Status = Translate("Completing") & Timer.ToString(1)

      End Select
    End With
  End Function

  Public Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
    Status = ""
    Timer.Cancel()
    controlCode.AddControl.FillCancel()
  End Sub

  Public ReadOnly Property IsOn() As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property

  Public ReadOnly Property IsActive As Boolean
    Get
      Return (State >= EState.WaitIdle) AndAlso (State < EState.Done)
    End Get
  End Property

#If 0 Then
  VERSION 1.0 CLASS
'===============================================================================================
'AF - Addition tank fill command
'===============================================================================================
Option Explicit
Implements ACCommand
Public Enum AFState
  Off
  Interlock
  Fill
  FillPause
End Enum
Public State As AFState
Public StateString As String
Public FillLevel As Long
Public FillType As Long
Public AddMixing As Boolean
Public Timer As New acTimer
Public BackFillTimer As New acTimer
Public OverrunTimer As New acTimer
  
Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Parameters=|HCMV| |0-99|% Mix?|0-1|\r\nName=Add Fill\r\nHelp=Fills the addition tank to the specified level with the specified water type. "
Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  
  ' Command Parameters:
  '   Parameters=|HCMV| |0-99|% Mix?|0-1|
  '   Name=Add Fill
  '   Help=Fills the addition tank to the specified level with the specified water type.

 
  ' Notes for fill type H=72, C=67, M=77, V=86
  '   Letter Parameters would normally be a variant type, but Plant explorer doesn't display variants so convert to long
  If Param(1) = 86 Then
    FillType = 86 'Runback from vessel
  ElseIf Param(1) = 77 Then
    FillType = 77
  ElseIf Param(1) = 72 Then
    FillType = 72
  Else: FillType = 67
  End If

  FillLevel = Param(2) * 10
  If Param(3) = 1 Then
    AddMixing = True
  Else
    AddMixing = False
  End If
    
'Carry on my wayward son...
  StepOn = True
  State = Interlock
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)

'Early bind ControlCode for speed and convenience
  Dim ControlCode As ControlCode
  Set ControlCode = ControlObject

  With ControlCode
    Select Case State
    
      Case Off
        StateString = ""
        
      Case Interlock
        If (.KP.KP1.DispenseTank = 1) Then
          OverrunTimer.TimeRemaining = 300
          Timer = 2
          StateString = "AF Wait: " & .KP.KP1.StateString
        ElseIf (.KA.Destination = 65) Then
          OverrunTimer.TimeRemaining = 300
          Timer = 2
          StateString = "AF Wait: " & .KA.StateString
        Else
          If FillType = 86 Then
            If Not .TempSafe Then
              Timer = 2
              StateString = "AF: Runback Wait for TempSafe "
            End If
          Else
            StateString = "AF Starting: " & TimerString(Timer.TimeRemaining)
          End If
        End If
        If Timer.Finished Then
          BackFillTimer.TimeRemaining = .Parameters.AddRunbackPulseTime
          State = Fill
        End If
         
      Case Fill
        If (FillType = 72) Then
          StateString = "AF: Fill Hot " & Pad(.AdditionLevel / 10, "0", 3) & "% / " & Pad(FillLevel / 10, "0", 3) & "%"
        ElseIf (FillType = 67) Then
          StateString = "AF: Fill Cold " & Pad(.AdditionLevel / 10, "0", 3) & "% / " & Pad(FillLevel / 10, "0", 3) & "%"
        ElseIf (FillType = 77) Then
          StateString = "AF: Fill Mixed " & Pad(.AdditionLevel / 10, "0", 3) & "% / " & Pad(FillLevel / 10, "0", 3) & "%"
        ElseIf (FillType = 86) Then
          StateString = "AF: Fill From Vessel " & Pad(.AdditionLevel / 10, "0", 3) & "% / " & Pad(FillLevel / 10, "0", 3) & "%"
          If BackFillTimer.Finished And .Parameters.AddRunbackPulseTime > 0 Then
            BackFillTimer.TimeRemaining = .Parameters.AddRunbackPulseTime
            State = FillPause
          End If
        End If
        'Issue when backfilling add tank on 72 package machines results in overfilling due to too fast flow
        If (.AdditionLevel >= (FillLevel - .Parameters.AdditionFillDeadband)) Then State = Off
         
      Case FillPause
        StateString = "AF: Paused " & TimerString(BackFillTimer.TimeRemaining)
        If (.AdditionLevel >= (FillLevel - .Parameters.AdditionFillDeadband)) Then State = Off
        If BackFillTimer.Finished Or (.Parameters.AddRunbackPulseTime = 0) Then
          BackFillTimer.TimeRemaining = .Parameters.AddRunbackPulseTime
          State = Fill
        End If

     End Select
  End With
  
End Sub
Public Sub ACCommand_Cancel()
  State = Off
End Sub

Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property
Friend Property Get IsOn() As Boolean
  If (State > Off) Then IsOn = True
End Property
Friend Property Get IsActive() As Boolean
  If IsOn And ((State = Fill) Or (State = FillPause)) Then IsActive = True
End Property
Friend Property Get IsInterlocked() As Boolean
If (State = Interlock) Then IsInterlocked = True
End Property
Friend Property Get IsFillWithHot() As Boolean
  If (State = Fill) And (FillType = 72) Then IsFillWithHot = True
End Property
Friend Property Get IsFillWithCold() As Boolean
  If (State = Fill) And (FillType = 67) Then IsFillWithCold = True
End Property
Friend Property Get IsFillWithMixed() As Boolean
  If (State = Fill) And (FillType = 77) Then IsFillWithMixed = True
End Property
Friend Property Get IsFillWithVessel() As Boolean
  If (State = Fill) And (FillType = 86) Then IsFillWithVessel = True
End Property
Friend Property Get IsFillPause() As Boolean
  If (State = FillPause) Then IsFillPause = True
End Property
Friend Property Get IsOverrun() As Boolean
  If IsOn And OverrunTimer.Finished Then IsOverrun = True
End Property

#End If



#If 0 Then
  ' TODO
VERSION 1.0 CLASS
Attribute VB_Name = "AF"

'===============================================================================================
'AF - Addition tank fill command
'===============================================================================================
Option Explicit
Implements ACCommand
Public Enum AFState
  Off
  Interlock
  Fill
  FillPause
End Enum
Public State As AFState
Public StateString As String
Public FillLevel As Long
Public FillType As Long
Public AddMixing As Boolean
Public Timer As New acTimer
Public BackFillTimer As New acTimer
Public OverrunTimer As New acTimer
  



Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)

'Early bind ControlCode for speed and convenience
  Dim ControlCode As ControlCode
  Set ControlCode = ControlObject

  With ControlCode
    Select Case State
    
      Case Off
        StateString = ""
        
      Case Interlock
        If (.KP.KP1.DispenseTank = 1) Then
          OverrunTimer.TimeRemaining = 300
          Timer = 2
          StateString = "AF Wait: " & .KP.KP1.StateString
        ElseIf (.KA.Destination = 65) Then
          OverrunTimer.TimeRemaining = 300
          Timer = 2
          StateString = "AF Wait: " & .KA.StateString
        Else
          If FillType = 86 Then
            If Not .TempSafe Then
              Timer = 2
              StateString = "AF: Runback Wait for TempSafe "
            End If
          Else
            StateString = "AF Starting: " & TimerString(Timer.TimeRemaining)
          End If
        End If
        If Timer.Finished Then
          BackFillTimer.TimeRemaining = .Parameters.AddRunbackPulseTime
          State = Fill
        End If
         
      Case Fill
        If (FillType = 72) Then
          StateString = "AF: Fill Hot " & Pad(.AdditionLevel / 10, "0", 3) & "% / " & Pad(FillLevel / 10, "0", 3) & "%"
        ElseIf (FillType = 67) Then
          StateString = "AF: Fill Cold " & Pad(.AdditionLevel / 10, "0", 3) & "% / " & Pad(FillLevel / 10, "0", 3) & "%"
        ElseIf (FillType = 77) Then
          StateString = "AF: Fill Mixed " & Pad(.AdditionLevel / 10, "0", 3) & "% / " & Pad(FillLevel / 10, "0", 3) & "%"
        ElseIf (FillType = 86) Then
          StateString = "AF: Fill From Vessel " & Pad(.AdditionLevel / 10, "0", 3) & "% / " & Pad(FillLevel / 10, "0", 3) & "%"
          If BackFillTimer.Finished And .Parameters.AddRunbackPulseTime > 0 Then
            BackFillTimer.TimeRemaining = .Parameters.AddRunbackPulseTime
            State = FillPause
          End If
        End If
        'Issue when backfilling add tank on 72 package machines results in overfilling due to too fast flow
        If (.AdditionLevel >= (FillLevel - .Parameters.AdditionFillDeadband)) Then State = Off
         
      Case FillPause
        StateString = "AF: Paused " & TimerString(BackFillTimer.TimeRemaining)
        If (.AdditionLevel >= (FillLevel - .Parameters.AdditionFillDeadband)) Then State = Off
        If BackFillTimer.Finished Or (.Parameters.AddRunbackPulseTime = 0) Then
          BackFillTimer.TimeRemaining = .Parameters.AddRunbackPulseTime
          State = Fill
        End If

     End Select
  End With
  
End Sub
Public Sub ACCommand_Cancel()
  State = Off
End Sub

Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property
Friend Property Get IsOn() As Boolean
  If (State > Off) Then IsOn = True
End Property
Friend Property Get IsActive() As Boolean
  If IsOn And ((State = Fill) Or (State = FillPause)) Then IsActive = True
End Property
Friend Property Get IsInterlocked() As Boolean
If (State = Interlock) Then IsInterlocked = True
End Property
Friend Property Get IsFillWithHot() As Boolean
  If (State = Fill) And (FillType = 72) Then IsFillWithHot = True
End Property
Friend Property Get IsFillWithCold() As Boolean
  If (State = Fill) And (FillType = 67) Then IsFillWithCold = True
End Property
Friend Property Get IsFillWithMixed() As Boolean
  If (State = Fill) And (FillType = 77) Then IsFillWithMixed = True
End Property
Friend Property Get IsFillWithVessel() As Boolean
  If (State = Fill) And (FillType = 86) Then IsFillWithVessel = True
End Property
Friend Property Get IsFillPause() As Boolean
  If (State = FillPause) Then IsFillPause = True
End Property
Friend Property Get IsOverrun() As Boolean
  If IsOn And OverrunTimer.Finished Then IsOverrun = True
End Property

#End If
End Class