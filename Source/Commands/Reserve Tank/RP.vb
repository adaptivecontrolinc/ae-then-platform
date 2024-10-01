'American & Efird - MX Then Multiflex Package
' Version 2016-08-18

Imports Utilities.Translations

<Command("Reserve Prepare", "Msg:|0-99| Mix:|0-99|mins", "", "", ""), _
TranslateCommand("es", "Reserve preparar", "Msg:|0-99| Mix:|0-99|mins"), _
Description("Signal the operator to prepare the reserve tank"), _
TranslateDescription("es", "Una señal al operador para preparar el depósito de reserva."),
Category("Reserve Tank Commands"), TranslateCategory("es", "Reserve Tank Commands")>
Public Class RP : Inherits MarshalByRefObject : Implements ACCommand

  'Command states
  Public Enum EState
    Off
    WaitIdle
    Slow
    Fast
    FillToMixLevel
    MixingDelay
    Ready
  End Enum
  Public Property State As EState
  Public Property StateString As String
  Public Property Message As Integer
  Public Property MixLevel As Integer
  Public Property MixRequested As Boolean
  Public Property MixTime As Integer
  Public Property Calloff As Integer

  Public Property Timer() As New Timer
  Friend TimerOverrun As New Timer
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

  'Uses these for flashy lamp stuff
  Private ReadOnly slowFlasher As New Flasher(800)
  Private ReadOnly fastFlasher As New Flasher(400)

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

      ' Cancel reserve tank commands
      .RD.Cancel()

      'Command parameters
      Message = 0
      MixLevel = 0
      MixTime = 0
      MixRequested = False

      'Check array bounds just to be on the safe side
      If param.GetUpperBound(0) >= 1 Then Message = param(1)
      If param.GetUpperBound(0) >= 2 Then MixTime = param(2) * 60

      ' Calculate Mix Level
      MixLevel = MinMax(.Parameters.ReserveHeatEnableLevel, .Parameters.ReserveMixerOnLevel, 750)

      ' NOTE: MX Then Machines do not have Reserve Mixers or Circulation Piping (Like Fong)
      '       Reset Mix Time, but left state logic for reference
      MixTime = 0

      'Get CallOff Value
      Calloff = Utilities.Programs.GetCallOffNumber(.Parent.CurrentStep, .Parent.Programs)

      'Set default state
      State = EState.WaitIdle

      'Set overrun timer
      Me.TimerOverrun.Minutes = .Parameters.StandardTimeAddPrepare

      'Tell the control system to step on
      Return True
    End With
  End Function

  Public Function Run() As Boolean Implements ACCommand.Run
    With controlCode

      Select Case State
        Case EState.Off
          StateString = (" ")

        Case EState.WaitIdle
          StateString = Translate("Wait Idle") & Timer.ToString
          If .RF.IsForeground Then StateString = .RF.Status : Timer.Seconds = 1
          If .RD.IsOn Then StateString = .RD.Status : Timer.Seconds = 1
          ' Kitchen Tank
          ' TODO
          If (.KA.IsOn AndAlso .KA.Destination = EKitchenDestination.Reserve) Then StateString = .KA.Status : Timer.Seconds = 1
          If (.KP.KP1.IsOn AndAlso (.KP.KP1.DispenseTank = 2)) Then StateString = .KP.KP1.DrugroomDisplay : Timer.Seconds = 1
          'If (.KP.KP1.IsOn AndAlso (Not .KP.KP1.IsBackground) AndAlso (.KP.KP1.Param_Destination = EKitchenDestination.Reserve)) Then StateString = .KP.KP1.DrugroomDisplay : Timer.Seconds = 1
          If (.LA.KP1.IsOn AndAlso (Not .LA.KP1.IsBackground) AndAlso (.LA.KP1.Param_Destination = EKitchenDestination.Reserve)) Then StateString = .LA.KP1.DrugroomDisplay : Timer.Seconds = 1
          If (.Parent.IsPaused OrElse .EStop) Then Timer.Seconds = 1
          ' Reserve tank idle
          If Timer.Finished Then State = EState.Slow

        Case EState.Slow
          StateString = Translate("Wait for Reserve Ready ")
          If .ReserveReady Then
            ' MixTime requested/remaining
            If MixTime > 0 Then
              State = EState.FillToMixLevel
            Else : State = EState.Ready
            End If
          End If

        Case EState.Fast
          StateString = Translate("Wait for Reserve Ready")
          If .ReserveReady Then
            ' MixTime requested/remaining
            If MixTime > 0 Then
              State = EState.FillToMixLevel
            Else : State = EState.Ready
            End If
          End If

        Case EState.FillToMixLevel
          ' Calculate Mix Level
          MixLevel = MinMax(.Parameters.ReserveHeatEnableLevel, .Parameters.ReserveMixerOnLevel, 750)
          StateString = Translate("Filling To Mix") & (" ") & (.ReserveLevel / 10).ToString("#0.0") & " / " & (MixLevel / 10).ToString("#0.0") & "% "
          If .ReserveLevel >= MixLevel Then
            If MixTime > 0 Then MixRequested = True
            State = EState.MixingDelay
            Timer.Seconds = MixTime
          End If

        Case EState.MixingDelay
          StateString = Translate("Mixing") & Timer.ToString(1)
          If Timer.Finished Then State = EState.Ready
          ' Lost Reserve Ready
          If Not .ReserveReady Then
            State = EState.Slow
            ' Update MixTime with time remaining
            MixTime = Timer.Seconds
          End If

        Case EState.Ready
          StateString = Translate("Reserve Ready")
          ' Lost Reserve Ready
          If Not .ReserveReady Then
            State = EState.Slow
          End If

      End Select
    End With
  End Function

  Public Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
    StateString = Nothing
    Calloff = 0
    MixRequested = False
  End Sub

  Public ReadOnly Property IsOn() As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property

  Public ReadOnly Property IsOverrun() As Boolean
    Get
      Return IsOn AndAlso IsWaitReady AndAlso TimerOverrun.Finished
    End Get
  End Property

  Public ReadOnly Property IsWaitReady() As Boolean
    Get
      Return (State = EState.Slow) OrElse (State = EState.Fast) OrElse _
             (State = EState.FillToMixLevel) OrElse (State = EState.MixingDelay)
    End Get
  End Property

  Public ReadOnly Property IoReserveRinse As Boolean
    Get
      Return (State = EState.FillToMixLevel)
    End Get
  End Property

  Public ReadOnly Property IoReserveLamp() As Boolean
    Get
      Return (State = EState.FillToMixLevel) OrElse (State = EState.MixingDelay) OrElse
              (State = EState.Ready) OrElse ((State = EState.Slow) And controlCode.FlashSlow) OrElse ((State = EState.Fast) And controlCode.FlashFast)
    End Get
  End Property




  ' TODO
#If 0 Then


  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===============================================================================================
'RP - Reserve Tank Prepare Command
'===============================================================================================
Option Explicit
Implements ACCommand
Public Enum RPState
  Off
  WaitIdle
  Slow
  Fast
  Ready
End Enum
Public State As RPState
Public StateString As String
Public ReservePreparePrompt As Long
Public AdvanceTimer As New acTimer

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Parameters=Prompt |0-99|\r\nName=Reserve Prepare\r\nHelp=Signals the operator to prepare the reserve tank.  If AutoReady set to '1' then tank will automatically be set to ""Ready"" once command starts."
Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  ReservePreparePrompt = Param(1)
  
  ControlCode.ReserveReadyFromPB = False
  
'Carry on my wayward son...
  StepOn = True
  State = WaitIdle
  AdvanceTimer.TimeRemaining = 2
  
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)

'Early bind ControlCode for speed and convenience
  Dim ControlCode As ControlCode
  Set ControlCode = ControlObject

  With ControlCode
    Select Case State
      Case Off
        StateString = ""
        
      Case WaitIdle
        StateString = "RP: Wait to lose Run Pushbutton "
        If Not .IO_RemoteRun Then
          AdvanceTimer.TimeRemaining = 2
          State = Slow
        End If
      
      Case Slow
        StateString = "RP: Hold Run for Ready " & TimerString(AdvanceTimer.TimeRemaining)
        If Not .IO_RemoteRun Then AdvanceTimer.TimeRemaining = 2
        If .RT.IsWaitReady Then State = Fast
        If .ReserveReadyFromPB Or AdvanceTimer.Finished Then
          .ReserveReady = True
          State = Ready
        End If
         
      Case Fast
        StateString = "RP: Hold Run for Ready " & TimerString(AdvanceTimer.TimeRemaining)
        If Not .IO_RemoteRun Then AdvanceTimer.TimeRemaining = 2
        If .ReserveReadyFromPB Or AdvanceTimer.Finished Then
          .ReserveReady = True
          State = Ready
        End If
      
      Case Ready
        StateString = "RP: Ready "
        
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
  IsActive = IsOn
End Property
Friend Property Get IsSlow() As Boolean
  If (State = Slow) Then IsSlow = True
End Property
Friend Property Get IsFast() As Boolean
  If (State = Fast) Then IsFast = True
End Property
Friend Property Get IsWaitReady() As Boolean
  If IsSlow Or IsFast Then IsWaitReady = False
End Property
Friend Property Get IsReady() As Boolean
  If (State = Ready) Then IsReady = True
End Property

#End If
End Class

