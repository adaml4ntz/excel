VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PropertyForm 
   Caption         =   "Property"
   ClientHeight    =   5070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4335
   OleObjectBlob   =   "PropertyForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "PropertyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub propBox_Click()
    
    Dim lookup As String
    Static cnt As Long
     
     'get list box selection(s)
    For X = 0 To propBox.ListCount - 1
        If propBox.Selected(X) = True Then
            lookup = propBox.List(X)
            y = X + 2
        End If
    Next X
    
    Me.Hide
    AccountForm.Show
 
End Sub

Private Sub UserForm_Initialize()

    'declare variables and data types
    Dim LR As Long
    Dim Lrow As Long
    Dim Test As New Collection
    Dim i As Single
    Dim value As Variant
    Dim temp As Range
    
    On Error Resume Next 'ignore errors in macro and continue with next line

    Set temp = Worksheets("META").ListObjects("LOOKUP").ListColumns(2).Range 'save values from first column in 'accts' table to temp object
 
    'iterate through values in first column except header value
    For i = 2 To temp.Cells.Rows.Count
         If Len(temp.Cells(i)) > 0 And Not temp.Cells(i).EntireRow.Hidden Then 'save value to test if the number of characters are more than 0 and the row is not hidden
           Test.Add temp.Cells(i), CStr(temp.Cells(i))
         End If
    Next i 'continue with next value
 
    propBox.Clear 'delete all items in listbox

    For Each value In Test 'iterate through all values saved to test
        acct = value
        propBox.AddItem value 'add values to list box
    Next value 'continue with next value

    Set Test = Nothing 'delete object test
        
End Sub




