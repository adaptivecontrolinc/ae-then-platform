'American & Efird - Mt. Holly Then Platform
' Version 2024-09-04

Imports Utilities.Translations

<Command("Wait for Kitchen", "", " ", "", ""), 
TranslateCommand("es", "Esperar para la cocina", ""), 
Description("Wait for kichen commands to complete before continuing program."), 
TranslateDescription("es", "Esperar órdenes de cocina completar antes de continuar el programa."), 
Category("Kitchen Functions")> 
Public Class WK : Inherits MarshalByRefObject : Implements ACCommand
  Private Const commandName_ As String = ("WK: ")
  Private ReadOnly controlCode As ControlCode

  Public Enum Estate
    Off
    [On] = 1
    Delay
    Done
  End Enum
  Property State As Estate
  Property Status As String
  Property Timer As New Timer
  Friend Property TimerOverrun As New Timer
  Public ReadOnly Property TimeOverrun As Integer
    Get
      Return TimerOverrun.Seconds
    End Get
  End Property

  Public Sub New(ByVal controlCode As ControlCode)
    Me.controlCode = controlCode
    Me.controlCode.Commands.Add(Me)
  End Sub

  Sub ParametersChanged(ByVal ParamArray param() As Integer) Implements ACCommand.ParametersChanged
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
      .AT.Cancel() : .DR.Cancel() : .FI.Cancel()
      .HD.Cancel() : .HC.Cancel() : .LD.Cancel()
      .PH.Cancel() : .RC.Cancel() : .RH.Cancel()
      .RI.Cancel() : .RT.Cancel() : .RW.Cancel()
      .SA.Cancel() : .TM.Cancel() : .UL.Cancel()
      .WT.Cancel()

      ' Command parameters

      ' Initial State
      State = Estate.On
      Timer.Seconds = 5

    End With
  End Function

  Function Run() As Boolean Implements ACCommand.Run
    With controlCode

      Select Case State

        Case Estate.Off
          Status = (" ")

        Case Estate.On
          'Changed acAddPrepare to look for WK.IsON, and if so, to not set the ready flags.
          'Using the DispenseTank value wouldn't work because its' value was being cleared to 0 before cancelling the KP
          'Need to make sure that (KP.ison = false) as well as (LAActive = false) and (LA.IsTank1Active = false)
          Status = Translate("Waiting for Kitchen")
          If (.KP.IsOn OrElse .LA.IsActive) Then
            If .KP.IsOn Then Status = .KP.KP1.DrugroomDisplay
            If .LA.IsActive Then Status = .LA.KP1.DrugroomDisplay
          Else
            State = Estate.Done
            Timer.Seconds = 5
          End If
          If TimerOverrun.Finished Then State = Estate.Delay

        Case Estate.Delay
          Status = Translate("Delay Off") & Timer.ToString(1)
          If (.KP.KP1.IsOn AndAlso .KP.KP1.IsWaitReady) OrElse (.LA.KP1.IsOn AndAlso .LA.KP1.IsWaitReady) Then
            If .KP.IsOn Then Status = .KP.KP1.DrugroomDisplay
            If .LA.IsActive Then Status = .LA.KP1.DrugroomDisplay
          Else
            State = Estate.Done
            Timer.Seconds = 5
          End If

        Case Estate.Done
          Status = Translate("Completing") & Timer.ToString(1)
          If Timer.Finished Then
            .AddReady = False
            .ReserveReady = False
            Cancel()
          End If

      End Select

    End With
  End Function

  Sub Cancel() Implements ACCommand.Cancel
    State = Estate.Off
    Status = ""
    Timer.Cancel()
    TimerOverrun.Cancel()
  End Sub
  ReadOnly Property IsOn As Boolean Implements ACCommand.IsOn
    Get
      Return State <> Estate.Off
    End Get
  End Property
  ReadOnly Property IsOverrun As Boolean
    Get
      Return IsOn AndAlso TimerOverrun.Finished
    End Get
  End Property

  ' TODO check and remove:
#If 0 Then
Public Enum WKState
  WKOff
  WKSetup
  WKOn
  WKDelayOff
End Enum
Public State As WKState
Attribute State.VB_VarUserMemId = 0
Public Timer As New acTimer

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Name=Wait for Kitchen\r\nHelp=Waits for the kitchen transfer to be complete."
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode
    .AC.ACCommand_Cancel
    .AT.ACCommand_Cancel: .DR.ACCommand_Cancel: .FI.ACCommand_Cancel
    .HD.ACCommand_Cancel: .HC.ACCommand_Cancel: .LD.ACCommand_Cancel
    .PH.ACCommand_Cancel: .RC.ACCommand_Cancel: .RH.ACCommand_Cancel
    .RI.ACCommand_Cancel: .RT.ACCommand_Cancel: .RW.ACCommand_Cancel
    .SA.ACCommand_Cancel: .TM.ACCommand_Cancel: .UL.ACCommand_Cancel
    .WT.ACCommand_Cancel
    
    Timer.TimeRemaining = 10
    State = WKOn
  End With
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
   With ControlCode
    Select Case State
    
      Case WKOff
      
      Case WKSetup
      
      Case WKOn
      'Changed acAddPrepare to look for WK.IsON, and if so, to not set the ready flags.
      'Using the DispenseTank value wouldn't work because it's value was being cleared to 0 before cancelling the KP
      'Need to make sure that (KP.ison = false) as well as (LAActive = false) and (LA.IsTank1Active = false)
      If Not (.KP.IsOn And .LAActive And .LA.KP1.IsOn) Then
        Timer.TimeRemaining = 10
        State = WKDelayOff
      End If
      
      Case WKDelayOff
      If (.KP.IsOn Or .LAActive Or .LA.KP1.IsOn) Then Timer.TimeRemaining = 10
      If Timer.Finished Then
        .AddReady = False
        .ReserveReady = False
        ACCommand_Cancel
      End If
    
    End Select
   End With
End Sub

Friend Sub ACCommand_Cancel()
  State = WKOff
End Sub
Friend Property Get IsOn() As Boolean
  IsOn = (State <> WKOff)
End Property
Friend Property Get IsDelayOff() As Boolean
  IsDelayOff = (State = WKDelayOff)
End Property
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property

#End If
End Class

