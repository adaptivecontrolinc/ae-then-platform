Public Class acManualAddTransfer
#If 0 Then
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ACManualAddTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===============================================================================================
'manual fill from add buttons
'===============================================================================================
'
  Option Explicit
  Public Enum ManualAddTransferState
    Off
    Interlock
    Transfer1
    Rinse
    Transfer2
    RinseToDrain
    Drain
    TurnOff
  End Enum
  Public State As ManualAddTransferState
  Public Timer As New acTimer
  

Public Sub Run(AddLevel As Long, AddTransferPB As Boolean, AddTransferFinished As Boolean, _
              TransferTimeBeforeRinse As Long, RinseTime As Long, TransferTimeAfterRinse As Long, RinseToDrainTime As Long, _
              DrainTime As Long)


    Select Case State
       
       Case Off
          If AddTransferPB Then
            State = Interlock
            AddTransferFinished = False
          End If
          
       Case Interlock
          If Not AddTransferPB Then State = Off
          State = Transfer1
          Timer = TransferTimeBeforeRinse
        
        Case Transfer1
          If Not AddTransferPB Then State = Off
          If AddLevel > 10 Then Timer = TransferTimeBeforeRinse
          If Timer.Finished Then
             State = Rinse
             Timer = RinseTime
          End If
          
        Case Rinse
          If Not AddTransferPB Then State = Off
          If Timer.Finished Then
             State = Transfer2
             Timer = TransferTimeAfterRinse
          End If
          
        Case Transfer2
          If Not AddTransferPB Then State = Off
          If AddLevel > 10 Then Timer = TransferTimeAfterRinse
          If Timer.Finished Then
             State = RinseToDrain
             Timer = RinseToDrainTime
          End If
        
        Case RinseToDrain
          If Not AddTransferPB Then State = Off
          If Timer.Finished Then
             State = Drain
             Timer = DrainTime
          End If
          
        Case Drain
          If Not AddTransferPB Then State = Off
          If AddLevel > 10 Then Timer = DrainTime
          If Timer.Finished Then
             State = TurnOff
             AddTransferFinished = True
          End If
          
        Case TurnOff
           If AddTransferPB = False Then
             AddTransferFinished = False
             State = Off
           End If
    
    End Select
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
