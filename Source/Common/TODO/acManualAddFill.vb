Public Class acManualAddFill
#If 0 Then
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ACManualAddFill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===============================================================================================
'manual fill from add buttons
'===============================================================================================
'
  Option Explicit
  Public Enum ManualAddFillState
    Off
    Interlock
    FillingHot
    FillingCold
  End Enum
  Public State As ManualAddFillState
  

Public Sub Run(FillLevel As Long, AddLevel As Long, FillLevelDB As Long, FillHot As Boolean, _
              FillCold As Boolean)


    Select Case State
       
       Case Off
          If FillCold Or FillHot Then
             State = Interlock
          End If
          
       Case Interlock
          If FillHot Then
             State = FillingHot
          End If
          If FillCold Then
            State = FillingCold
          End If
          
        
       Case FillingHot
          If Not FillHot Then State = Off
          If (AddLevel >= ((FillLevel * 10) - FillLevelDB)) Then
             State = Off
             FillHot = False
          End If
         
        Case FillingCold
         If Not FillCold Then State = Off
          If (AddLevel >= ((FillLevel * 10) - FillLevelDB)) Then
             State = Off
             FillCold = False
          End If
          
    End Select
End Sub
Public Property Get IsFillingHot() As Boolean
  If (State = FillingHot) Then IsFillingHot = True
End Property
Public Property Get IsFillingCold() As Boolean
  If (State = FillingCold) Then IsFillingCold = True

End Property


#End If
End Class
