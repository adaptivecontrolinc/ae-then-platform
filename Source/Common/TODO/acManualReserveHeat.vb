Public Class acManualReserveHeat
#If 0 Then
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ACManualReserveHeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===============================================================================================
'manual fill from add buttons
'===============================================================================================
'
  Option Explicit
  Public Enum ManualReserveHeatState
    Off
    Interlock
    Heating
  End Enum
  Public State As ManualReserveHeatState
  

Public Sub Run(DesiredTemp As Long, ReserveTemp As Integer, ReserveLevel As Long, Heat As Boolean)


    Select Case State
       
       Case Off
          If Heat And ReserveLevel > 100 Then
             State = Interlock
          End If
          
       Case Interlock
          If Heat Then State = Heating
          
        
       Case Heating
          If Not Heat Then State = Off
          If (ReserveTemp >= (DesiredTemp * 10)) Then
             State = Off
             Heat = False
          End If
         
          
    End Select
End Sub
Public Property Get IsHeating() As Boolean
  If (State = Heating) Then IsHeating = True
End Property


#End If
End Class
