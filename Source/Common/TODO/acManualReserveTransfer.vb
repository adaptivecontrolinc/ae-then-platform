Public Class acManualReserveTransfer
#If 0 Then
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ACManualReserveTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===============================================================================================
'manual fill from add buttons
'===============================================================================================
'
  Option Explicit
  Public Enum ManualReserveTransferState
    Off
    Interlock
    Transfer1
    Rinse
    Transfer2
    RinseToDrain
    Drain
    TurnOff
  End Enum
  Public State As ManualReserveTransferState
  Public Timer As New acTimer
  

Public Sub Run(ByVal ControlObject As Object)
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  With ControlCode:


    Select Case State
       
       Case Off
          If .ManualReserveTransferPB Then
            State = Interlock
          End If
          
       Case Interlock
          If Not .ManualReserveTransferPB Then State = Off
          If .MachineSafe And .LidLocked Then
            State = Transfer1
            Timer = .Parameters.ReserveTimeBeforeRinse
          End If
          
        Case Transfer1
          If Not .ManualReserveTransferPB Then State = Off
          If Not .LidLocked Then State = Interlock
          If .ReserveLevel > 10 Then Timer = .Parameters.ReserveTimeBeforeRinse
          If Timer.Finished Then
             State = Rinse
             Timer = .Parameters.ReserveRinseTime
          End If
          
        Case Rinse
          If Not .ManualReserveTransferPB Then State = Off
          If Timer.Finished Then
             State = Transfer2
             Timer = .Parameters.ReserveTimeAfterRinse
          End If
          
        Case Transfer2
          If Not .ManualReserveTransferPB Then State = Off
          If Not .LidLocked Then State = Interlock
          If .ReserveLevel > 10 Then Timer = .Parameters.ReserveTimeAfterRinse
          If Timer.Finished Then
             State = RinseToDrain
             Timer = .Parameters.ReserveRinseToDrainTime
          End If
        
        Case RinseToDrain
          If Not .ManualReserveTransferPB Then State = Off
          If Timer.Finished Then
             State = Drain
             Timer = .Parameters.ReserveDrainTime
          End If
          
        Case Drain
          If Not .ManualReserveTransferPB Then State = Off
          If .ReserveLevel > 10 Then Timer = .Parameters.ReserveDrainTime
          If Timer.Finished Then
             State = TurnOff
             .ManualReserveTransferFinished = True
          End If
          
        Case TurnOff
           If .ManualReserveTransferPB = False Then
             .ManualReserveTransferFinished = False
             State = Off
           End If
    
    End Select
    End With
End Sub
Public Property Get IsTransferring() As Boolean
 If (State = Transfer1) Or (State = Rinse) Or (State = Transfer2) Then IsTransferring = True
End Property
Public Property Get IsRinsing() As Boolean
  If (State = Rinse) Or (State = RinseToDrain) Then IsRinsing = True
End Property
Public Property Get IsDraining() As Boolean
  If (State = Drain) Or (State = RinseToDrain) Then IsDraining = True
End Property


#End If
End Class
