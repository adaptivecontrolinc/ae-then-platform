'American & Efird - Then Platform Package
' Version 2022-03-23

Imports Utilities.Translations

<Command("Pressurize", "Off/On: |0-1|", "", "", ""), _
TranslateCommand("es", "Presurizar", "Off/On: |0-1|"), _
Description("Pressurize the machine (1) or depressurize (0)."), _
TranslateDescription("es", "Presurizar la máquina (1) o despresurizar (0)."), _
Category("Machine Functions")> _
Public Class PR
  Inherits MarshalByRefObject
  Implements ACCommand
  Private Const commandName_ As String = ("PR: ")

  'Command States
  Public Enum EState
    Off
    Interlock
    Active
  End Enum
  Public Property State As EState
  Public Property StateString As String
  Property MachineNotClosed As Boolean

  Friend ReadOnly Timer As New Timer
  Friend TimerAdvance As New Timer
  Public ReadOnly Property TimeAdvance As Integer
    Get
      Return TimerAdvance.Seconds
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

      'Command parameters
      Dim pressurize As Integer

      'Check array bounds just to be on the safe side
      If param.GetUpperBound(0) >= 1 Then pressurize = param(1)

      ' Clear Parent.Signal 
      If .Parent.Signal <> "" Then .Parent.Signal = ""

      'Cancel command if pressurize = 0
      If pressurize = 0 Then
        Cancel()
      Else
        ' Check Lid closed
        If Not .MachineClosed Then
          .Parent.Signal = Translate("Close Machine Lid")
        End If

        ' Set flag for machine not closed
        If Not (.MachineClosed) Then MachineNotClosed = True

        ' Default State
        State = EState.Interlock
      End If

      'Hold At this step if Command Pressurize, and Expansion Tank not closed
      Return False
    End With
  End Function

  Public Function Run() As Boolean Implements ACCommand.Run
    With controlCode

      ' Safe to continue
      Dim safe As Boolean = (Not .EStop) AndAlso .MachineClosed

      Select Case State
        Case EState.Off
          StateString = (" ")

        Case EState.Interlock
          StateString = Translate("Checking Lids")
          If Not .MachineClosed Then StateString = Translate("Close Machine Lid")
          If Not safe Then Timer.Seconds = 5
          ' Flag set where machine/sample/expansion limit switch was lost - operator must hold advance to clear and continue <<Redundancy>>
          If MachineNotClosed Then
            Timer.Seconds = 5
            If Not .AdvancePb Then
              StateString = Translate("Hold Advance") & TimerAdvance.ToString(1)
              TimerAdvance.Seconds = 2
            Else
              'Attempting to Advance
              If .Parent.Signal <> "" Then .Parent.Signal = ""
              StateString = Translate("Advancing") & TimerAdvance.ToString(1)
              If TimerAdvance.Finished Then MachineNotClosed = False
            End If
          End If
          ' Machine is Safe to continue
          If safe AndAlso Timer.Finished Then
            State = EState.Active
            'Continue Stepping on
            Return True
          End If

        Case EState.Active
          StateString = Translate("Active")

      End Select
    End With
  End Function

  Public Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
  End Sub

  Public ReadOnly Property IsOn() As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property

  Public ReadOnly Property IsInterlock As Boolean
    Get
      Return (State = EState.Interlock)
    End Get
  End Property

  Public ReadOnly Property IsActive As Boolean
    Get
      Return (State = EState.Active)
    End Get
  End Property


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
Attribute VB_Name = "PR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Circulate command

Option Explicit
Implements ACCommand
Public Enum PRState
  PROff
  PROn
End Enum
Public State As PRState
Attribute State.VB_VarUserMemId = 0
Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Parameters=01?:|0-1|\r\nName=Pressurize\r\nHelp=Pressurizes (1) or depressurises (0) the machine."
Dim ControlCode As ControlCode: Set ControlCode = ControlObject
  If Param(1) = 0 Then State = PROff
  If Param(1) = 1 Then State = PROn
  StepOn = True
End Sub
Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
End Sub
Friend Sub ACCommand_Cancel()
  State = PROff
End Sub
Friend Property Get IsOn() As Boolean
  IsOn = (State <> PROff)
End Property
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property


#End If
End Class
