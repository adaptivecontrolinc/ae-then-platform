'American & Efird - MX Package
' Version 2022-05-09

Imports Utilities.Translations

<Command("Kitchen Rinse", "Machine:|0-1| Drain:|0-1|", " ", "", ""), _
TranslateCommand("es", "enjuague de cocina", ""), _
Description("Turns off the Drugroom Rinse to Machine and Rinse to Drain steps.  Set parameter to 0 to turn off specific rinse."), 
TranslateDescription("es", "Apaga el enjuague Drugroom a máquina y enjuagar a los pasos de drenaje. Establezca el parámetro en 0 para desactivar la aclaración específica."), 
Category("Kitchen Functions"), TranslateCategory("es", "Funciones de cocina")> 
Public Class KR : Inherits MarshalByRefObject : Implements ACCommand
  Private Const commandName_ As String = ("KR: ")
  Private ReadOnly controlCode As ControlCode

  Public Enum EState
    Off
    Active
  End Enum
  Property State As EState
  Property StateString As String
  Property RinseMachine As Integer
  Property RinseDrain As Integer

  Sub New(ByVal controlCode As ControlCode)
    Me.controlCode = controlCode
    Me.controlCode.Commands.Add(Me)
  End Sub
  Sub ParametersChanged(ByVal ParamArray param() As Integer) Implements ACCommand.ParametersChanged
    'If the command is not running then do nothing the changes will be picked up when the command starts
    If Not IsOn Then Exit Sub

    'Enable command parameter adjustments, if parameter set
    If ControlCode.Parameters.EnableCommandChg = 1 Then
      Start(param)
    End If
  End Sub

  Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With ControlCode

      ' Cancel Commands

      ' Command parameters
      If param.GetUpperBound(0) >= 1 Then RinseMachine = param(1)
      If param.GetUpperBound(0) >= 2 Then RinseDrain = param(2)

      ' Default State
      State = EState.Active

      'This is a background command so step on
      Return True
    End With
  End Function

  Function Run() As Boolean Implements ACCommand.Run
    With ControlCode

      Select Case State

        Case EState.Off
          StateString = (" ")

        Case EState.Active
          If RinseMachine = 0 AndAlso RinseDrain = 0 Then
            StateString = Translate("Disable All Rinses")
          ElseIf RinseMachine = 0 Then
            StateString = Translate("Disable Rinse to Machine")
          ElseIf RinseDrain = 0 Then
            StateString = Translate("Disable Rinse to Drain")
          End If
          If (RinseMachine > 0) AndAlso (RinseDrain > 0) Then Cancel()
          If Not .Parent.IsProgramRunning Then Cancel()

      End Select
    End With
  End Function

  Public Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
    RinseMachine = 0
    RinseDrain = 0
  End Sub
  ReadOnly Property IsOn As Boolean Implements ACCommand.IsOn
    Get
      IsOn = State <> EState.Off
    End Get
  End Property

  ReadOnly Property IsRinseMachineOff As Boolean
    Get
      Return (State = EState.Active) AndAlso (RinseMachine = 0)
    End Get
  End Property

  ReadOnly Property IsRinseDrainOff As Boolean
    Get
      Return (State = EState.Active) AndAlso (RinseDrain = 0)
    End Get
  End Property


  ' TODO VB6
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
Attribute VB_Name = "KR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Kitchen Rinse = Turn Off Tank 1 Rinse to Machine....Only works once
Option Explicit
Implements ACCommand
Public Enum KRState
  KROff
  KROn
End Enum
Public RinseMachine As Integer
Public RinseDrain As Boolean
Public State As KRState

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Parameters=Machine |0-1| Drain |0-1|\r\nName=Kitchen Rinse Control\r\nHelp=Turns off the Drugroom Rinse to Machine and Rinse to Drain steps.  Set parameter to 0 to turn off specific rinse."
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode
    RinseMachine = Param(1)
    RinseDrain = Param(2)
    State = KROn
    StepOn = True
  End With
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
   With ControlCode
    Select Case State
    
     Case KROff
       RinseMachine = 0
       RinseDrain = 0
            
     Case KROn
      If Not .Parent.IsProgramRunning Then State = KROff
         
    End Select
   End With
End Sub

Friend Sub ACCommand_Cancel()
  State = KROff
End Sub
Friend Property Get IsOn() As Boolean
  IsOn = (State <> KROff)
End Property
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property


#End If
End Class
