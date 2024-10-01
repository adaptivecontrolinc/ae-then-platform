Public Class acRecipeSteps
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



#End If
End Class
