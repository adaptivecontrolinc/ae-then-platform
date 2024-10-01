'American & Efird - MX Package
' Version 2016-08-17

<Command("End Timing", "", "", "", ""), _
  TranslateCommand("Sincronización Del Final", "", ""), _
  Description("Stops timing associated with SA and BR commands, for production reports."), _
  TranslateDescription("es", "Utilizado conjuntamente con (SA) y comandos (BR) parar aprobaciones, muestras o el nuevo tratamiento de la sincronización."), _
  Category("Production Reports"), TranslateCategory("es", "Production Reports")> _
Public Class ET
  Inherits MarshalByRefObject
  Implements ACCommand

  'Command States
  Public Enum EState
    Off
    Active
  End Enum
  Property State As EState
  Property StateString As String

  Private ReadOnly ControlCode As ControlCode
  Sub New(ByVal controlCode As ControlCode)
    Me.ControlCode = controlCode
  End Sub
  Sub ParametersChanged(ByVal ParamArray param() As Integer) Implements ACCommand.ParametersChanged
  End Sub
  Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With ControlCode
      .BR.Cancel()
      .LD.Cancel()
      .PH.Cancel()
      .SA.Cancel()
      .UL.Cancel()
      Cancel()
    End With
  End Function

  Function Run() As Boolean Implements ACCommand.Run
  End Function

  Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
  End Sub

  ReadOnly Property IsOn As Boolean Implements ACCommand.IsOn
    Get
      Return State <> EState.Off
    End Get
  End Property

  ' TODO CHeck
#If 0 Then
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'End Timing command

Option Explicit
Implements ACCommand
Private Enum ETState
  ETOff
  ETOn
End Enum
Private State As ETState
Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Name=End Timing\r\nHelp=Used in conjunction with BA, BS and BR commands to stop timing approvals, samples or reprocessing."
Dim ControlCode As ControlCode: Set ControlCode = ControlObject

 With ControlCode
   .LD.ACCommand_Cancel
   .UL.ACCommand_Cancel
   .SA.ACCommand_Cancel
   .BR.ACCommand_Cancel
 End With
  State = ETOff
End Sub
Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
End Sub
Friend Sub ACCommand_Cancel()
  State = ETOff
End Sub
Friend Property Get IsOn() As Boolean
  IsOn = (State <> ETOff)
End Property
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property




#End If
End Class
