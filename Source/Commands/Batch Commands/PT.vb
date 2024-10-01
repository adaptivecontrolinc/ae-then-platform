'American & Efird - MX Then Multiflex Package
' Version 2015-09-15

' NOTE: A&E Mt. Holly uses package typ 1-8
'   1-8 = A-H, else 0, "NS" for "Not Set"

<Command("Package Type", "|0-4|", " ", "", "", CommandType.BatchParameter), _
TranslateCommand("es", "Type de paquetes:", "|0-4|"), _
Description("Sets the package type (0-4) loaded.  Used to calculated Working Level when parameters for corresponding package type at different levels are set."), _
TranslateDescription("es", ""), _
Category("Batch Commands"), TranslateCategory("es", "Batch Commands")> _
Public Class PT
  Inherits MarshalByRefObject
  Implements ACCommand

  'Command States
  Public Enum EState
    Off
    Active
  End Enum
  Property State As EState
  Property PackageType As Integer
  Private Property timer_ As New Timer
  Property OverrunTimer As New Timer

  'Keep a local reference to the control code object for convenience
  Private ReadOnly controlCode As ControlCode

  Public Sub New(ByVal controlCode As ControlCode)
    Me.controlCode = controlCode
    Me.controlCode.Commands.Add(Me)
  End Sub

  Public Sub ParametersChanged(ByVal ParamArray param() As Integer) Implements ACCommand.ParametersChanged
    'If the command is not running then do nothing the changes will be picked up when the command starts
    If Not IsOn Then Exit Sub
  End Sub

  Function Start(ByVal ParamArray param() As Integer) As Boolean Implements ACCommand.Start
    With controlCode

      'Check array bounds just to be on the safe side
      If param.GetUpperBound(0) >= 1 Then PackageType = param(1)

      .PackageType = PackageType

      ' Determine working level based on NP, PT, & Corresponding Level Parameters
      If (.PackageHeight > 0) AndAlso (.PackageHeight <= 4) Then
        Select Case .PackageType
          Case 0
            .WorkingLevel = .Parameters.PackageType0Level(.PackageHeight)
          Case 1
            .WorkingLevel = .Parameters.PackageType1Level(.PackageHeight)
          Case 2
            .WorkingLevel = .Parameters.PackageType2Level(.PackageHeight)
          Case 3
            .WorkingLevel = .Parameters.PackageType3Level(.PackageHeight)
          Case 4
            .WorkingLevel = .Parameters.PackageType4Level(.PackageHeight)
          Case 5
            .WorkingLevel = .Parameters.PackageType5Level(.PackageHeight)
          Case 6
            .WorkingLevel = .Parameters.PackageType6Level(.PackageHeight)
          Case 7
            .WorkingLevel = .Parameters.PackageType7Level(.PackageHeight)
          Case 8
            .WorkingLevel = .Parameters.PackageType8Level(.PackageHeight)
        End Select

        ' Step to active state
        State = EState.Active
      Else
        State = EState.Off
      End If

      'Foreground Step for moment
      Return True
    End With
  End Function

  Function Run() As Boolean Implements ACCommand.Run
    With controlCode

      Select Case State
        Case EState.Off

        Case EState.Active
          If Not .Parent.IsProgramRunning Then
            .PackageType = 0
            .WorkingLevel = 0
            State = EState.Off

          End If
      End Select
    End With
  End Function

  Sub Cancel() Implements ACCommand.Cancel
    State = EState.Off
  End Sub

  Public ReadOnly Property IsOn() As Boolean Implements ACCommand.IsOn
    Get
      Return (State <> EState.Off)
    End Get
  End Property

#If 0 Then

' TODO Remove

  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Package Type
Option Explicit
Implements ACCommand
Private Enum PTState
  PTOff
  PTOn
End Enum
Public PackageType As String
Public CalculatedWorkingLevel As Long
Private State As PTState

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Parameters=Type: |0-8|\r\nName=Package Type \r\nHelp=Sets the package type (A-G) loaded.  Used to calculated Working Level when parameters for corresponding package type at different levels are set.\r\nEnterWithDyelot=True"
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  
  With ControlCode
  
    If Param(1) = 0 Then
      PackageType = "0"
      .PackageType = 0
    ElseIf Param(1) = 1 Then
      PackageType = "A"
      .PackageType = 1
    ElseIf Param(1) = 2 Then
      PackageType = "B"
      .PackageType = 2
    ElseIf Param(1) = 3 Then
      PackageType = "C"
      .PackageType = 3
    ElseIf Param(1) = 4 Then
      PackageType = "D"
      .PackageType = 4
    ElseIf Param(1) = 5 Then
      PackageType = "E"
      .PackageType = 5
    ElseIf Param(1) = 6 Then
      PackageType = "F"
      .PackageType = 6
    ElseIf Param(1) = 7 Then
      PackageType = "G"
      .PackageType = 7
    ElseIf Param(1) = 8 Then
      PackageType = "H"
      .PackageType = 8
    Else
      'default to 0 - not used
      PackageType = "NS"
      .PackageType = 0
    End If
    
   'Determine Working Level Based on NP, PT, & Corresponding Level Parameter
    If (.PackageHeight > 0) Then
      If .PackageHeight = 1 Then
        If .PackageType = 0 Then
          .WorkingLevel = .Parameters.PackageType0Level1
        ElseIf .PackageType = 1 Then
          .WorkingLevel = .Parameters.PackageType1Level1
        ElseIf .PackageType = 2 Then
          .WorkingLevel = .Parameters.PackageType2Level1
        ElseIf .PackageType = 3 Then
          .WorkingLevel = .Parameters.PackageType3Level1
        ElseIf .PackageType = 4 Then
          .WorkingLevel = .Parameters.PackageType4Level1
        ElseIf .PackageType = 5 Then
          .WorkingLevel = .Parameters.PackageType5Level1
        ElseIf .PackageType = 6 Then
          .WorkingLevel = .Parameters.PackageType6Level1
        ElseIf .PackageType = 7 Then
          .WorkingLevel = .Parameters.PackageType7Level1
        Else  '.PackageType = 8
          .WorkingLevel = .Parameters.PackageType8Level1
        End If
      ElseIf .PackageHeight = 2 Then
        If .PackageType = 0 Then
          .WorkingLevel = .Parameters.PackageType0Level2
        ElseIf .PackageType = 1 Then
          .WorkingLevel = .Parameters.PackageType1Level2
        ElseIf .PackageType = 2 Then
          .WorkingLevel = .Parameters.PackageType2Level2
        ElseIf .PackageType = 3 Then
          .WorkingLevel = .Parameters.PackageType3Level2
        ElseIf .PackageType = 4 Then
          .WorkingLevel = .Parameters.PackageType4Level2
        ElseIf .PackageType = 5 Then
          .WorkingLevel = .Parameters.PackageType5Level2
        ElseIf .PackageType = 6 Then
          .WorkingLevel = .Parameters.PackageType6Level2
        ElseIf .PackageType = 7 Then
          .WorkingLevel = .Parameters.PackageType7Level2
        Else  '.PackageType = 8
          .WorkingLevel = .Parameters.PackageType8Level2
        End If
      ElseIf .PackageHeight = 3 Then
        If .PackageType = 0 Then
          .WorkingLevel = .Parameters.PackageType0Level3
        ElseIf .PackageType = 1 Then
          .WorkingLevel = .Parameters.PackageType1Level3
        ElseIf .PackageType = 2 Then
          .WorkingLevel = .Parameters.PackageType2Level3
        ElseIf .PackageType = 3 Then
          .WorkingLevel = .Parameters.PackageType3Level3
        ElseIf .PackageType = 4 Then
          .WorkingLevel = .Parameters.PackageType4Level3
        ElseIf .PackageType = 5 Then
          .WorkingLevel = .Parameters.PackageType5Level3
        ElseIf .PackageType = 6 Then
          .WorkingLevel = .Parameters.PackageType6Level3
        ElseIf .PackageType = 7 Then
          .WorkingLevel = .Parameters.PackageType7Level3
        Else  '.PackageType = 8
          .WorkingLevel = .Parameters.PackageType8Level3
        End If
      ElseIf .PackageHeight = 4 Then
        If .PackageType = 0 Then
          .WorkingLevel = .Parameters.PackageType0Level4
        ElseIf .PackageType = 1 Then
          .WorkingLevel = .Parameters.PackageType1Level4
        ElseIf .PackageType = 2 Then
          .WorkingLevel = .Parameters.PackageType2Level4
        ElseIf .PackageType = 3 Then
          .WorkingLevel = .Parameters.PackageType3Level4
        ElseIf .PackageType = 4 Then
          .WorkingLevel = .Parameters.PackageType4Level4
        ElseIf .PackageType = 5 Then
          .WorkingLevel = .Parameters.PackageType5Level4
        ElseIf .PackageType = 6 Then
          .WorkingLevel = .Parameters.PackageType6Level4
        ElseIf .PackageType = 7 Then
          .WorkingLevel = .Parameters.PackageType7Level4
        Else  '.PackageType = 8
          .WorkingLevel = .Parameters.PackageType8Level4
        End If
      ElseIf .PackageHeight = 5 Then
        If .PackageType = 0 Then
          .WorkingLevel = .Parameters.PackageType0Level5
        ElseIf .PackageType = 1 Then
          .WorkingLevel = .Parameters.PackageType1Level5
        ElseIf .PackageType = 2 Then
          .WorkingLevel = .Parameters.PackageType2Level5
        ElseIf .PackageType = 3 Then
          .WorkingLevel = .Parameters.PackageType3Level5
        ElseIf .PackageType = 4 Then
          .WorkingLevel = .Parameters.PackageType4Level5
        ElseIf .PackageType = 5 Then
          .WorkingLevel = .Parameters.PackageType5Level5
        ElseIf .PackageType = 6 Then
          .WorkingLevel = .Parameters.PackageType6Level5
        ElseIf .PackageType = 7 Then
          .WorkingLevel = .Parameters.PackageType7Level5
        Else  '.PackageType = 8
          .WorkingLevel = .Parameters.PackageType8Level5
        End If
      ElseIf .PackageHeight = 6 Then
        If .PackageType = 0 Then
          .WorkingLevel = .Parameters.PackageType0Level6
        ElseIf .PackageType = 1 Then
          .WorkingLevel = .Parameters.PackageType1Level6
        ElseIf .PackageType = 2 Then
          .WorkingLevel = .Parameters.PackageType2Level6
        ElseIf .PackageType = 3 Then
          .WorkingLevel = .Parameters.PackageType3Level6
        ElseIf .PackageType = 4 Then
          .WorkingLevel = .Parameters.PackageType4Level6
        ElseIf .PackageType = 5 Then
          .WorkingLevel = .Parameters.PackageType5Level6
        ElseIf .PackageType = 6 Then
          .WorkingLevel = .Parameters.PackageType6Level6
        ElseIf .PackageType = 7 Then
          .WorkingLevel = .Parameters.PackageType7Level6
        Else  '.PackageType = 8
          .WorkingLevel = .Parameters.PackageType8Level6
        End If
      ElseIf .PackageHeight = 7 Then
        If .PackageType = 0 Then
          .WorkingLevel = .Parameters.PackageType0Level7
        ElseIf .PackageType = 1 Then
          .WorkingLevel = .Parameters.PackageType1Level7
        ElseIf .PackageType = 2 Then
          .WorkingLevel = .Parameters.PackageType2Level7
        ElseIf .PackageType = 3 Then
          .WorkingLevel = .Parameters.PackageType3Level7
        ElseIf .PackageType = 4 Then
          .WorkingLevel = .Parameters.PackageType4Level7
        ElseIf .PackageType = 5 Then
          .WorkingLevel = .Parameters.PackageType5Level7
        ElseIf .PackageType = 6 Then
          .WorkingLevel = .Parameters.PackageType6Level7
        ElseIf .PackageType = 7 Then
          .WorkingLevel = .Parameters.PackageType7Level7
        Else  '.PackageType = 8
          .WorkingLevel = .Parameters.PackageType8Level7
        End If
      ElseIf .PackageHeight = 8 Then
        If .PackageType = 0 Then
          .WorkingLevel = .Parameters.PackageType0Level8
        ElseIf .PackageType = 1 Then
          .WorkingLevel = .Parameters.PackageType1Level8
        ElseIf .PackageType = 2 Then
          .WorkingLevel = .Parameters.PackageType2Level8
        ElseIf .PackageType = 3 Then
          .WorkingLevel = .Parameters.PackageType3Level8
        ElseIf .PackageType = 4 Then
          .WorkingLevel = .Parameters.PackageType4Level8
        ElseIf .PackageType = 5 Then
          .WorkingLevel = .Parameters.PackageType5Level8
        ElseIf .PackageType = 6 Then
          .WorkingLevel = .Parameters.PackageType6Level8
        ElseIf .PackageType = 7 Then
          .WorkingLevel = .Parameters.PackageType7Level8
        Else  '.PackageType = 8
          .WorkingLevel = .Parameters.PackageType8Level8
        End If
      Else
        If .PackageType = 0 Then
          .WorkingLevel = .Parameters.PackageType0Level9
        ElseIf .PackageType = 1 Then
          .WorkingLevel = .Parameters.PackageType1Level9
        ElseIf .PackageType = 2 Then
          .WorkingLevel = .Parameters.PackageType2Level9
        ElseIf .PackageType = 3 Then
          .WorkingLevel = .Parameters.PackageType3Level9
        ElseIf .PackageType = 4 Then
          .WorkingLevel = .Parameters.PackageType4Level9
        ElseIf .PackageType = 5 Then
          .WorkingLevel = .Parameters.PackageType5Level9
        ElseIf .PackageType = 6 Then
          .WorkingLevel = .Parameters.PackageType6Level9
        ElseIf .PackageType = 7 Then
          .WorkingLevel = .Parameters.PackageType7Level9
        Else  '.PackageType = 8
          .WorkingLevel = .Parameters.PackageType8Level9
        End If
      End If
    End If
  State = PTOff
  
  End With
  
End Sub

Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
End Sub
Friend Sub ACCommand_Cancel()
  State = PTOff
End Sub
Friend Property Get IsOn() As Boolean
  IsOn = (State <> PTOff)
End Property
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property



#End If
End Class
