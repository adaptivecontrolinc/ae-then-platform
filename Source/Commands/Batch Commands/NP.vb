'American & Efird - Mt. Holly Then Platform to GMX
' Version 2022-02-03

<Command("Number Of Packages", "|0-999|", " ", "", "", CommandType.BatchParameter), _
 TranslateCommand("es", "Número de paquetes:", "|0-999|"), _
 Description("Sets the number of packages for the kier. This is used to calculate the working level."), _
 TranslateDescription("es", "Fija el número de los paquetes para el ma's kier. Esto se utiliza para calcular el sistema de trabajo de level. If a cero o mayor a de 8 entonces la máquina se inunda completamente."), _
  Category("Batch Commands"), TranslateCategory("es", "Batch Commands")> _
Public Class NP
  Inherits MarshalByRefObject
  Implements ACCommand

  'Command States
  Public Enum EState
    Off
    Active
  End Enum
  Property State As EState
  Property NumberPackages As Integer
  Property NumberSpindles As Integer
  Public PackageHeightCalculated As Integer

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

      ' Check array bounds just to be on the safe side
      If param.GetUpperBound(0) >= 1 Then NumberPackages = param(1)

      ' Calculate Package Height
      If (.Parameters.NumberOfSpindales > 0) Then
        ' Capacity to Isolate one kier
        If .HL.KierBLoaded Then
          NumberSpindles = .Parameters.NumberOfSpindales * 2
        Else
          NumberSpindles = .Parameters.NumberOfSpindales
        End If

        ' Calculate package height
        PackageHeightCalculated = CInt(NumberPackages / NumberSpindles)
        If PackageHeightCalculated >= 9 Then
          .PackageHeight = 9
        ElseIf PackageHeightCalculated >= 8 Then
          .PackageHeight = 8
        ElseIf PackageHeightCalculated >= 7 Then
          .PackageHeight = 7
        ElseIf PackageHeightCalculated >= 6 Then
          .PackageHeight = 6
        ElseIf PackageHeightCalculated >= 5 Then
          .PackageHeight = 5
        ElseIf PackageHeightCalculated >= 4 Then
          .PackageHeight = 4
        ElseIf PackageHeightCalculated >= 3 Then
          .PackageHeight = 3
        ElseIf PackageHeightCalculated >= 2 Then
          .PackageHeight = 2
        ElseIf PackageHeightCalculated >= 1 Then
          .PackageHeight = 1
        Else
          .PackageHeight = 0
        End If
      End If

      ' Determine working level based on NP, PT, & Corresponding Level Parameters
      If (.PackageHeight > 0) AndAlso (.PackageHeight <= 9) Then
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
          Case 9
            .WorkingLevel = .Parameters.PackageType9Level(.PackageHeight)
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
            NumberPackages = 0
            NumberSpindles = 0
            .WorkingLevel = 0
            .PackageHeight = 0
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

  ' TODO Check
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
Attribute VB_Name = "NP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'number of packages
Option Explicit
Implements ACCommand
Private Enum NPState
  NPOff
  NPOn
End Enum
Public NumberOfPackages As Long
Public CalculatedPackHeight As Double
Private State As NPState

Public Sub ACCommand_Start(ByVal ControlObject As Object, Param() As Variant, StepOn As Boolean)
Attribute ACCommand_Start.VB_Description = "Parameters= |0-999|\r\nName=Number of Packages: \r\nHelp=Sets the number of packages for the kier. This is used to calculate the working level.\r\nEnterWithDyelot=True"
  Dim ControlCode As ControlCode:  Set ControlCode = ControlObject
  NumberOfPackages = Param(1)
  If NumberOfPackages < 0 Then NumberOfPackages = 0

  With ControlCode
    
    If (.Parameters.NumberOfSpindales > 0) Then
      CalculatedPackHeight = NumberOfPackages / .Parameters.NumberOfSpindales
      If CalculatedPackHeight > 8 Then
         .PackageHeight = 9
      ElseIf CalculatedPackHeight > 7 Then
        .PackageHeight = 8
      ElseIf CalculatedPackHeight > 6 Then
        .PackageHeight = 7
      ElseIf CalculatedPackHeight > 5 Then
        .PackageHeight = 6
      ElseIf CalculatedPackHeight > 4 Then
        .PackageHeight = 5
      ElseIf CalculatedPackHeight > 3 Then
        .PackageHeight = 4
      ElseIf CalculatedPackHeight > 2 Then
        .PackageHeight = 3
      ElseIf CalculatedPackHeight > 1 Then
        .PackageHeight = 2
      ElseIf CalculatedPackHeight = 1 Then
        .PackageHeight = 1
      Else
        .PackageHeight = 0
      End If
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
    
  End With
  State = NPOff

End Sub
Private Sub ACCommand_Run(ByVal ControlObject As Object, StepOn As Boolean)
End Sub
Friend Sub ACCommand_Cancel()
  State = NPOff
End Sub
Friend Property Get IsOn() As Boolean
  IsOn = (State <> NPOff)
End Property
Private Property Get ACCommand_IsOn() As Boolean
  ACCommand_IsOn = IsOn
End Property



#End If
End Class
