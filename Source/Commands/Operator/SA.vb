'American & Efird - Mt. Holly Then Platform to GMX
' Version 2022-03-23

Imports Utilities.Translations

<Command("Sample", "", "", "", "'StandardTimeSample=15"),
Description("Signals the operator to take a sample and starts recording the sample time. If the sample time exceeds the parameter Standard Sample Time then the overrun time is logged as Sample Delay."),
Category("Operator Functions")>
Public Class SA : Inherits MarshalByRefObject : Implements ACCommand
  Private ReadOnly controlCode As ControlCode
  Private Const commandName_ As String = ("SA: ")

  Public Enum EState
    Off
    Interlock
    '    SampleKierIsolate
    '    SampleKierDrain
    '    SampleKierRinse
    '    SampleKierDrain2
    Sample
    '   SampleLidCheck
    '   SampleKierDrainClose
    '   SampleKierUnIsolate
    Done
  End Enum
  Property State As EState
  Property StateString As String
  Private flashSlow_ As Boolean
  Private flashFast_ As Boolean

  Property Timer As New Timer
  Property TimerOverrun As New Timer

  Public Sub New(ByVal controlCode As ControlCode)
    Me.controlCode = controlCode
    Me.controlCode.Commands.Add(Me)
  End Sub

  Public Sub ParametersChanged(ByVal ParamArray param() As Integer) Implements ACCommand.ParametersChanged
    'If the command is not running then do nothing the changes will be picked up when the command starts
    If Not IsOn Then Exit Sub
    ' Do nothing
  End Sub

  Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With controlCode

      ' Cancel Commands
      .AC.Cancel()
      If .AT.IsForeground Then .AT.Cancel()
      .KA.Cancel() : .WK.Cancel()
      .DR.Cancel() : .FI.Cancel() : .HC.Cancel() : .HD.Cancel()
      .RI.Cancel() : .RC.Cancel() : .RH.Cancel() : .TM.Cancel()
      .LD.Cancel() : .PH.Cancel() ': .SA.Cancel() :
      .UL.Cancel()
      If .RF.IsForeground AndAlso .RF.FillType = EFillType.Vessel Then .RF.Cancel() ' TODO Check
      If .RT.IsForeground Then .RT.Cancel()
      .WT.Cancel()

      ' Clear Parent.Signal 
      If .Parent.Signal <> "" Then .Parent.Signal = ""

      TimerOverrun.Minutes = .Parameters.StandardTimeSample
      State = EState.Interlock
    End With
  End Function

  Function Run() As Boolean Implements ACCommand.Run
    With controlCode

      ' Flashers in sync
      flashSlow_ = .FlashSlow
      flashFast_ = .FlashFast

      Select Case State

        Case EState.Off
          StateString = (" ")

        Case EState.Interlock
          StateString = Translate("Wait Safe")
          If .TempSafe Then
            ' Reset Recorded Level status - this command drains the Sample Tank, losing some value
            .GetRecordedLevel = False
            '            State = EState.SampleKierIsolate
            '            Timer.Seconds = MinMax(.Parameters.SampleIsolateTime, 10, 120)
            .Parent.Signal = Translate("Sample")
            .AirpadOn = False ' TODO Check
            State = EState.Sample
            Timer.Seconds = 2
          End If

#If 0 Then

        Case EState.SampleKierIsolate
          StateString = Translate("Sample Isolating") & Timer.ToString(1)
          If Timer.Finished Then
            State = EState.SampleKierDrain
            Timer.Seconds = MinMax(.Parameters.SampleDrainTime, 10, 120)
          End If

        Case EState.SampleKierDrain
          StateString = Translate("Sample Draining") & Timer.ToString(1)
          If Timer.Finished Then
            State = EState.SampleKierRinse
            Timer.Seconds = MinMax(.Parameters.SampleRinseTime, 20, 300)
          End If

        Case EState.SampleKierRinse
          StateString = Translate("Sample Rinsing") &  & Timer.ToString(1)
          If Timer.Finished Then
            State = EState.SampleKierDrain2
            Timer.Seconds = MinMax(.Parameters.SampleDrainTime, 10, 120)
          End If

        Case EState.SampleKierDrain2
          StateString = Translate("Sample Draining") &  & Timer.ToString(1)
          If Timer.Finished Then
            .Parent.Signal = "Sample"
            State = EState.Sample
          End If

#End If

        Case EState.Sample
          StateString = Translate("Sampling") & (" ") & Translate("Hold Advance") & Timer.ToString(1)
          If Not .AdvancePb Then
            Timer.Seconds = 2
          Else
            If .Parent.Signal <> "" Then .Parent.Signal = ""
            ' Run Held, check Sample Lid closed for 2 seconds
            If Not (.IO.KierLidClosed) Then
              Timer.Seconds = 2
              If Not .MachineClosed Then StateString = Translate("Machine Not Closed")
            Else
              StateString = Translate("Advancing") & Timer.ToString(1)
              If Timer.Finished Then
                '   State = EState.SampleLidCheck
                '   Timer.Seconds = 2
                .IO.AdvancePb = False
                State = EState.Done
                Timer.Seconds = 5
                If .PumpControl.PumpEnabled AndAlso Not .PumpControl.IsActive Then
                  .PumpControl.StartAuto() : Timer.Seconds = 10
                End If
              End If
            End If
          End If

#If 0 Then

        Case EState.SampleLidCheck
          StateString = Translate("Checking Lids") &  & Timer.ToString(1)
          If Not (.IO.KierLidClosed) Then
            If Not .MachineClosed Then StateString = Translate("Machine Not Closed")
            Timer.Seconds = 2
          Else
            If Timer.Finished Then
              State = EState.SampleKierDrainClose
              Timer.Seconds = 5
            End If
          End If

        Case EState.SampleKierDrainClose
          StateString = Translate("Close Drain") &  & Timer.ToString(1)
          If Timer.Finished Then
            State = EState.SampleKierUnIsolate
            Timer.Seconds = MinMax(.Parameters.SampleIsolateTime, 10, 120)
          End If

        Case EState.SampleKierUnIsolate
          StateString = Translate("Sample Unisolating") &  & Timer.ToString(1)
          If Timer.Finished Then
            State = EState.Complete
            Timer.Seconds = 5
            If .PumpControl.PumpEnabled AndAlso Not .PumpControl.IsActive Then
              .PumpControl.StartAuto() : Timer.Seconds = 10
            End If
          End If

#End If

        Case EState.Done
          StateString = Translate("Completing") & Timer.ToString(1)
          ' If pump is enabled, then wait for pump to start, otherwise complete command and step on
          If .PumpControl.IsStarting Then StateString = .PumpControl.StateString
          If .PumpControl.PumpEnabled AndAlso Not .PumpControl.IsActive Then
            If Not .PumpControl.IsRunning Then Timer.Seconds = 5
          End If
          If Timer.Finished Then Cancel()

      End Select
    End With
  End Function

  Public Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
    StateString = ""
    Timer.Cancel()
    TimerOverrun.Cancel()
  End Sub

  ReadOnly Property IsOn As Boolean Implements ACCommand.IsOn
    Get
      IsOn = (State <> EState.Off)
    End Get
  End Property

  ReadOnly Property IsActive As Boolean
    Get
      Return (State > EState.Interlock)
    End Get
  End Property

  ReadOnly Property IsOverrun As Boolean
    Get
      Return IsOn AndAlso TimerOverrun.Finished
    End Get
  End Property

  ReadOnly Property IoLampSignal As Boolean
    Get
      Return (State = EState.Sample) AndAlso flashSlow_
    End Get
  End Property

#If 0 Then


  ReadOnly Property IoSampleUnlock As Boolean
    Get
      Return (State = EState.Sample)
    End Get
  End Property

  ReadOnly Property IsSampleIsolate As Boolean
    Get
      Return (State >= EState.SampleKierIsolate) AndAlso (State <= EState.SampleKierDrainClose)
    End Get
  End Property

  ReadOnly Property IoSampleDrain As Boolean
    Get
      Return (State >= EState.SampleKierDrain) AndAlso (State <= EState.SampleLidCheck)
    End Get
  End Property

  ReadOnly Property IoSampleRinse As Boolean
    Get
      Return (State = EState.SampleKierRinse)
    End Get
  End Property

#End If


#If 0 Then
' TODO

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Sample command
Option Explicit
Implements ACCommand
Public Enum SAState
  SAOff
  SANotSafe
  SASample
  SAComplete
End Enum
Public State As SAState
Public StateString As String
Public SASignalOnRequest As Boolean
Public AdvanceTimer As New acTimer
Public Timer As New acTimer
Private OverrunTimer As New acTimer

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Name=Sample\r\nMinutes=5\r\nHelp=Signals the operator to take a sample and starts recording the sample time. If the sample time exceeds the parameter Standard Sample Time then the overrun time is logged as Sample Delay."
Dim ControlCode As ControlCode: Set ControlCode = ControlObject
  With ControlCode
    .AC.ACCommand_Cancel
    .AT.ACCommand_Cancel: .DR.ACCommand_Cancel: .FI.ACCommand_Cancel
    .HD.ACCommand_Cancel: .HC.ACCommand_Cancel: .LD.ACCommand_Cancel
    .PH.ACCommand_Cancel: .RC.ACCommand_Cancel: .RH.ACCommand_Cancel
    If .RF.FillType = 86 Then .RF.ACCommand_Cancel
    .RI.ACCommand_Cancel: .RT.ACCommand_Cancel: .RW.ACCommand_Cancel
    .TM.ACCommand_Cancel: .UL.ACCommand_Cancel: .WK.ACCommand_Cancel
    .WT.ACCommand_Cancel
    .TemperatureControl.Cancel
    .TemperatureControlContacts.Cancel
    .PumpRequest = False
    .AirpadOn = False
     OverrunTimer = 300
  End With
  AdvanceTimer = 2
  State = SANotSafe
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
Dim ControlCode As ControlCode: Set ControlCode = ControlObject
  With ControlCode
    
    Select Case State
      Case SAOff
        StateString = ""
        
      Case SANotSafe
        StateString = "SA: Not Safe"
        If Not .IO_RemoteRun Then AdvanceTimer = 2
        If AdvanceTimer.Finished Then
          Timer.TimeRemaining = 5
          State = SAComplete
        End If
        If .MachineSafe Then
          State = SASample
          AdvanceTimer = 2
          SASignalOnRequest = True
        End If
     
      Case SASample
        StateString = "Sample, Hold Run to Continue " & TimerString(AdvanceTimer.TimeRemaining)
        If Not .MachineSafe Then
          State = SANotSafe
        End If
        If Not .IO_RemoteRun Then AdvanceTimer = 2
        If AdvanceTimer.Finished Then
          .PumpRequest = True
          .AirpadOn = True
          Timer.TimeRemaining = 5
          State = SAComplete
        End If
        
      Case SAComplete
        StateString = "SA: Completing " & TimerString(Timer.TimeRemaining)
        If Timer.Finished Then
          .PumpRequest = True
          .AirpadOn = True
          ACCommand_Cancel
        End If
           
    End Select
  End With
End Sub
Friend Sub ACCommand_Cancel()
  State = SAOff
  SASignalOnRequest = False
End Sub
Friend Property Get IsOn() As Boolean
  IsOn = (State <> SAOff)
End Property
Friend Property Get IsSampling() As Boolean
  If State = SASample Then IsSampling = True
End Property
Friend Property Get IsUnSafe() As Boolean
  If State = SANotSafe Then IsUnSafe = True
End Property
Friend Property Get IsOverrun() As Boolean
  If IsOn And OverrunTimer.Finished Then IsOverrun = True
End Property
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property


#End If

End Class
