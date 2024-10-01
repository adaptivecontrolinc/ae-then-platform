Public Class Recipe



  ' TODO - from VB6 - acRecipeRow.cls
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

  ' acRecipeRowCollection:
#If 0 Then
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acRecipeRowCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

  Private CurrentrecipeRows As Collection
'

Private Sub Class_Initialize()
  Set CurrentrecipeRows = New Collection
End Sub

Public Sub Add(recipeRow As acRecipeRow)

'Make sure we have something
  If recipeRow Is Nothing Then Exit Sub
  
'Add recipe row to collection
  CurrentrecipeRows.Add recipeRow ' , CStr(recipeRow.id)
  
End Sub

Public Function Count() As Long
  Count = CurrentrecipeRows.Count
End Function

Public Sub Delete(ByVal Index As Variant)
  CurrentrecipeRows.Remove Index
End Sub

Public Function Item(ByVal Index As Variant) As acRecipeRow
Attribute Item.VB_UserMemId = 0
   Set Item = CurrentrecipeRows.Item(Index)
End Function

' NewEnum must return the IUnknown interface of a collection's enumerator.
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = CurrentrecipeRows.[_NewEnum]
End Function
Public Sub Clear()
 Do While CurrentrecipeRows.Count > 0
  CurrentrecipeRows.Remove (1)
 Loop
End Sub


#End If

  ' acRecipeSteps
#If 0 Then
  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acRecipeSteps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Dim pDyelot As String
Dim pRedye As Integer
Dim pRecords As Integer
Dim pCollectionRows As Integer

Dim CurrentrecipeRows As acRecipeRowCollection
'

Private Sub Class_Initialize()
  Set CurrentrecipeRows = New acRecipeRowCollection
End Sub

Public Property Get Dyelot() As String
  Dyelot = pDyelot
End Property

Public Property Get Redye() As Integer
  Redye = pRedye
End Property

Public Property Get Records() As Integer
  Records = pRecords
End Property
Public Property Get CollectionRows() As Integer
  CollectionRows = pCollectionRows
End Property

Public Sub LoadRecipeSteps(Dyelot As String, Redye As Integer)
On Error GoTo Err_LoadRecipeSteps

'Save Dyelot redye locally
  pDyelot = Dyelot
  pRedye = Redye

'Reset records
  pRecords = 0
'clear collection
  CurrentrecipeRows.Clear

'Create ADO objects
  Dim con As ADODB.Connection
  Dim rs As ADODB.Recordset15
  Dim sql As String

  Set con = New ADODB.Connection
  Set rs = New ADODB.Recordset
  sql = "SELECT * FROM DyelotsBulkedRecipe WHERE BATCH_NUMBER='" & Dyelot & "' AND " & "Redye=" & Redye
  
'Open the connection and get the data
  con.Open "BatchDyeingCentral", "Adaptive", "Control"
  rs.Open sql, con, adOpenStatic, adLockReadOnly
 
'Make sure the recordset is fully populated then set records variable
 If rs.RecordCount > 0 Then
  rs.MoveLast: rs.MoveFirst
  pRecords = rs.RecordCount
 Else
  pRecords = 0
  pCollectionRows = 0
  Exit Sub
 End If
'Populate the recipe row collection
  Dim newRecipeRow As acRecipeRow
  Do While Not rs.EOF
    Set newRecipeRow = New acRecipeRow
    newRecipeRow.Fill rs
    CurrentrecipeRows.Add newRecipeRow
    rs.MoveNext
  Loop
   
 'set the number of collection rows just for info
   pCollectionRows = CurrentrecipeRows.Count
   
'FOR DEBUG ONLY - output basic step info
'  Dim message As String
'  rs.MoveFirst
'  Do While Not rs.EOF
'    message = rs("ID") & "  " & rs("INGREDIENT_ID") & "  " & rs("INGREDIENT_DESC")
'    Debug.Print message
'    rs.MoveNext
'  Loop
'  rs.MoveFirst
  
'Clean up
  con.Close
  Set con = Nothing
  Set rs = Nothing

'Avoid Error Handler
  Exit Sub

Err_LoadRecipeSteps:
  MsgBox Err.Description
  Set con = Nothing
  Set rs = Nothing
End Sub







  ' ***********************************
  RecipeSteps.vb:

  VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "acRecipeSteps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True



Dim pDyelot As String
Dim pReDye As Integer

Public Property Get Dyelot() As String
  Dyelot = pDyelot
End Property

Public Property Get ReDye() As Integer

End Property


Public Sub LoadRecipeSteps(Dyelot As String, ReDye As Integer)
On Error GoTo Err_LoadRecipeSteps

'Save Dyelot redye locally
  pDyelot = Dyelot
  pReDye = ReDye


  Dim con As ADODB.Connection
  Dim rs As ADODB.Recordset15



End Sub

















#End If

End Class
