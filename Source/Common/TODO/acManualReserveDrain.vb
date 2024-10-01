Public Class acManualReserveDrain
#If 0 Then
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ACManualReserveDrain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===============================================================================================
'manual fill from add buttons
'===============================================================================================
'
  Option Explicit
  Public Enum ManualReserveDrainState
    Off
    Interlock
    DrainEmpty
    Rinse
    Drain
    TurnOff
  End Enum
  Public State As ManualReserveDrainState
  Public Timer As New acTimer
  
Public Sub Run(ReserveLevel As Long, ReserveDrainPB As Boolean, ReserveDrainFinished As Boolean, _
              DrainTime As Long, RinseTime As Long)

    Select Case State
       
       Case Off
          If ReserveDrainPB Then
            State = Interlock
          End If
          
       Case Interlock
          If Not ReserveDrainPB Then State = Off
          If ReserveLevel > 10 Then
            State = DrainEmpty
            Timer = DrainTime
          Else
            State = Rinse
            Timer = RinseTime
          End If
        
        Case DrainEmpty
          If Not ReserveDrainPB Then State = Off
          If ReserveLevel > 10 Then Timer = DrainTime
          If Timer.Finished Then
            State = Rinse
            Timer = RinseTime
          End If
          
        Case Rinse
          If Not ReserveDrainPB Then State = Off
          If Timer.Finished Then
            State = Drain
            Timer = DrainTime
          End If
          
        Case Drain
          If Not ReserveDrainPB Then State = Off
          If ReserveLevel > 10 Then Timer = DrainTime
          If Timer.Finished Then
            Timer.TimeRemaining = 7
            ReserveDrainFinished = True
            State = TurnOff
          End If
          
        Case TurnOff
          If Not ReserveDrainPB And Timer.Finished Then
            ReserveDrainFinished = False
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
