Public Class LidControlDelete
#If 0 Then
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acLidControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum LidControlState
  Off
  ReleasePin
  OpenBand
  RaiseLid
  LidOpen
  LowerLid
  CloseBand
  EngagePin
End Enum
Public State As LidControlState
Public StateString As String
Public Timer As New acTimer

Private pLidClosed As Boolean
Private pLidOpen As Boolean

'===========================================================================================
Public Sub Run(MachineSafe As Boolean, LevelOkToOpen As Boolean, MainPump As Boolean, _
              LidClosedSwitch, LidLimitSwitch As Boolean, _
              CloseLidPB As Boolean, OpenLidPB As Boolean, _
              EmergencyStop As Boolean, _
              LockingBandOpenTime As Long, _
              LockingBandClosingTime As Long, _
              LidRaisingTime As Long)

  pLidClosed = LidClosedSwitch And LidLimitSwitch
  
  Select Case State
    
    Case Off
      StateString = ""
      'start open the lid if the machine is safe
      If OpenLidPB And MachineSafe And LevelOkToOpen And (Not MainPump) And (Not EmergencyStop) Then
         State = ReleasePin
      End If
      'lower the lid if its not already lower or locked.
      If CloseLidPB And (Not EmergencyStop) Then
        If LidClosedSwitch Then
          If Not LidLimitSwitch Then
            State = CloseBand
            Timer = LockingBandClosingTime
          Else
            State = EngagePin
          End If
        Else
          State = LowerLid
        End If
      End If

    Case ReleasePin
      StateString = "Release Pin " & TimerString(Timer.TimeRemaining)
      If LidLimitSwitch Then Timer = 5
      If Timer.Finished Then
         State = OpenBand
         Timer = LockingBandOpenTime
      End If
    
    Case OpenBand
      StateString = "Open Band " & TimerString(Timer.TimeRemaining)
      If Timer.Finished Then
         State = RaiseLid
         Timer = LidRaisingTime
      End If
    
    Case RaiseLid
      StateString = "Raise Lid " & TimerString(Timer.TimeRemaining)
      If Not OpenLidPB Then
        State = LidOpen
      End If
      If LidClosedSwitch Then Timer = LidRaisingTime
      If Timer.Finished Then
         State = LidOpen
      End If
    
    Case LidOpen
      StateString = "Lid Open "
      If OpenLidPB Then
         State = RaiseLid
         Timer = LidRaisingTime
      End If
      If CloseLidPB Then
        State = LowerLid
      End If
      
    Case LowerLid
      StateString = "Lower Lid " & TimerString(Timer.TimeRemaining)
      If (Not LidClosedSwitch) Then Timer = 3
      If Not CloseLidPB Then
        State = LidOpen
      End If
      If Timer.Finished Then
         State = CloseBand
         Timer = LockingBandClosingTime
      End If
          
    Case CloseBand
      StateString = "Close Band " & TimerString(Timer.TimeRemaining)
      If Timer.Finished Then
         State = EngagePin
      End If
          
     Case EngagePin
      StateString = "Raise Pin " & TimerString(Timer.TimeRemaining)
      If (Not LidLimitSwitch) Then Timer = 3
      If Timer.Finished Then
         State = Off
      End If
          
  End Select
End Sub
'===========================================================================================

Friend Property Get IsActive() As Boolean
  IsActive = (State <> Off)
End Property
Public Property Get IsReleasePin() As Boolean
  IsReleasePin = (State = ReleasePin) Or (State = OpenBand) Or (State = RaiseLid) Or (State = LidOpen) Or (State = LowerLid) Or (State = CloseBand)
End Property
Public Property Get IsOpenLockingBand() As Boolean
  IsOpenLockingBand = (State = OpenBand) Or (State = RaiseLid) Or (State = LidOpen) Or (State = LowerLid)
End Property
Public Property Get IsRaiseLid() As Boolean
  IsRaiseLid = (State = RaiseLid)
End Property
Public Property Get IsLowerLid() As Boolean
  IsLowerLid = (State = LowerLid)
End Property
Public Property Get IsCloseLockingBand() As Boolean
  IsCloseLockingBand = (State = CloseBand) Or (State = EngagePin) Or (State = Off)
End Property



#End If
End Class
