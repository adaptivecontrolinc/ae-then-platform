Public Class acManualAddDrain
#If 0 Then
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ACManualAddDrain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===============================================================================================
'manual fill from add buttons
'===============================================================================================
'
Option Explicit
Public Enum ManualAddDrainState
  Off
  Interlock
  DrainEmpty
  Rinse
  Drain
  TurnOff
End Enum
Public State As ManualAddDrainState
Public Timer As New acTimer
  
Public Sub Run(AddLevel As Long, AddDrainPB As Boolean, AddDrainFinished As Boolean, _
              DrainTime As Long, RinseTime As Long)

    Select Case State
       
       Case Off
          If AddDrainPB Then
            State = Interlock
          End If
          
       Case Interlock
          If Not AddDrainPB Then State = Off
          If AddLevel > 10 Then
             State = DrainEmpty
             Timer = DrainTime
          Else
             State = Rinse
             Timer = RinseTime
          End If
        
        Case DrainEmpty
          If Not AddDrainPB Then State = Off
          If AddLevel > 10 Then Timer = DrainTime
          If Timer.Finished Then
             State = Rinse
             Timer = RinseTime
          End If
          
        Case Rinse
          If Not AddDrainPB Then State = Off
          If Timer.Finished Then
             State = Drain
             Timer = DrainTime
          End If
          
        Case Drain
          If Not AddDrainPB Then State = Off
          If AddLevel > 10 Then Timer = DrainTime
          If Timer.Finished Then
             Timer.TimeRemaining = 7
             AddDrainFinished = True
             State = TurnOff
          End If
          
        Case TurnOff
          If Not AddDrainPB And Timer.Finished Then
             AddDrainFinished = False
             State = Off
          End If
          
    End Select
End Sub
Public Property Get IsActive() As Boolean
    IsActive = (State <> Off)
End Property
Public Property Get IsRinsing() As Boolean
  If (State = Rinse) Then IsRinsing = True
End Property
Public Property Get IsDraining() As Boolean
  If (State = Drain) Or (State = DrainEmpty) Or (State = Rinse) Then IsDraining = True
End Property
Public Property Get IsTurnOff() As Boolean
  If (State = TurnOff) Then IsTurnOff = True
End Property

#End If
End Class
