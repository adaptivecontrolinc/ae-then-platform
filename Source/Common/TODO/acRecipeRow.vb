Public Class acRecipeRow
#If 0 Then
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acRecipeRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

  Public id As String
  Public AddNumber As Integer
  Public AddSequence As Integer
  Public IngredientID As Integer
  Public IngredientDesc As String
  Public Dispensed As Boolean
'

Public Sub Fill(rs As Recordset)
On Error GoTo Err_Fill

'Make sure we have a recordset
  If rs Is Nothing Then Exit Sub
    
'Fill in the variables from the recordset
  With rs
    id = rs("ID")
    AddNumber = rs("ADD_NUMBER")
    AddSequence = rs("ADD_SEQUENCE")
    IngredientID = rs("INGREDIENT_ID")
    IngredientDesc = rs("INGREDIENT_DESC")
    Dispensed = (Left(rs("DISPENSED"), 1) = "A")
  End With
    
'Avoid Error handler
  Exit Sub

Err_Fill:

End Sub

Public Sub Clear()
  id = 0
  AddNumber = 0
  AddSequence = 0
  IngredientID = 0
  IngredientDesc = ""
  Dispensed = False
End Sub

#End If
End Class
