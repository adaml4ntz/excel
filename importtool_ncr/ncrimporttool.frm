VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NCRImportTool 
   Caption         =   "NCR IMPORT TOOL"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6300
   OleObjectBlob   =   "ncrimporttool.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NCRImportTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public invDate As String

Private Sub enterButton_Click()
     
    Dim lookup As String
    Static cnt As Long
     
     'get list box selection(s)
    For X = 0 To nameListBox.ListCount - 1
        If nameListBox.Selected(X) = True Then
            lookup = nameListBox.List(X)
        End If
    Next X
    
    If lookup = "" Or NCRImportTool.invTextBox.value = "" Then GoTo noSelection
    
    Sheets("META").Cells(1, 1).value = lookup
    'If lookup = "" Then GoTo errorNoSelection 'check to see if selection was made
   
         'INVOICE_NO
         Sheets("OUTPUT").Cells(cnt + 2, 1).value = NCRImportTool.invTextBox.value
        
         'PO_NO
         Sheets("OUTPUT").Cells(cnt + 2, 2).value = Sheets("META").Cells(1, 2).value
         
         'VENDOR_ID
         Sheets("OUTPUT").Cells(cnt + 2, 3).value = Sheets("META").Cells(1, 3).value
         
         'POSTING_DATE
         Sheets("OUTPUT").Cells(cnt + 2, 4).value = Sheets("META").Cells(1, 4).value
         
         'CREATED_DATE
         Sheets("OUTPUT").Cells(cnt + 2, 5).value = invDate
         
         'DUE_DATE
         Sheets("OUTPUT").Cells(cnt + 2, 6).value = invDate
         
         'DESCRIPTION
         'Sheets("OUTPUT").Cells(cnt + 2, 7).Value = Sheets("META").Cells(1, 7).Value
         Sheets("OUTPUT").Cells(cnt + 2, 7).value = "STAFFING; " & Sheets("META").Cells(1, 1).value
         
         'LINE_NO
         'Sheets("OUTPUT").Cells(cnt + 2, 8).Value = Sheets("META").Cells(1, 8).Value
         
         'MEMO
         'Sheets("OUTPUT").Cells(cnt + 2, 9).Value = Sheets("META").Cells(1, 9).Value
         Sheets("OUTPUT").Cells(cnt + 2, 9).value = "STAFFING; " & Sheets("META").Cells(1, 1).value
         
         'ACCT_NO
         Sheets("OUTPUT").Cells(cnt + 2, 10).value = Sheets("META").Cells(1, 10).value
         
         'LOCATION_ID
         Sheets("OUTPUT").Cells(cnt + 2, 11).value = Sheets("META").Cells(1, 11).value
         
         'AMOUNT
         'Sheets("OUTPUT").Cells(cnt + 2, 12).Value = Sheets("META").Cells(1, 12).Value
         Sheets("OUTPUT").Cells(cnt + 2, 12).value = NCRImportTool.amtTextBox.value
    
    cnt = cnt + 1
    
         'INVOICE_NO
         Sheets("OUTPUT").Cells(cnt + 2, 1).value = NCRImportTool.qainvTextBox.value
        
         'PO_NO
         Sheets("OUTPUT").Cells(cnt + 2, 2).value = Sheets("META").Cells(1, 2).value
         
         'VENDOR_ID
         Sheets("OUTPUT").Cells(cnt + 2, 3).value = Sheets("META").Cells(1, 3).value
         
         'POSTING_DATE
         Sheets("OUTPUT").Cells(cnt + 2, 4).value = Sheets("META").Cells(1, 4).value
         
         'CREATED_DATE
         Sheets("OUTPUT").Cells(cnt + 2, 5).value = invDate
         
         'DUE_DATE
         Sheets("OUTPUT").Cells(cnt + 2, 6).value = invDate
         
         'DESCRIPTION
         'Sheets("OUTPUT").Cells(cnt + 2, 7).Value = Sheets("META").Cells(1, 7).Value
         Sheets("OUTPUT").Cells(cnt + 2, 7).value = Sheets("META").Cells(1, 7).value & " QA SERVICES"
         
         'LINE_NO
         'Sheets("OUTPUT").Cells(cnt + 2, 8).Value = Sheets("META").Cells(1, 8).Value
         
         'MEMO
         'Sheets("OUTPUT").Cells(cnt + 2, 9).Value = Sheets("META").Cells(1, 9).Value
         Sheets("OUTPUT").Cells(cnt + 2, 9).value = Sheets("META").Cells(1, 7).value & " QA SERVICES"
         
         'ACCT_NO
         Sheets("OUTPUT").Cells(cnt + 2, 10).value = Sheets("META").Cells(1, 10).value
         
         'LOCATION_ID
         Sheets("OUTPUT").Cells(cnt + 2, 11).value = Sheets("META").Cells(1, 11).value
         
         'AMOUNT
         'Sheets("OUTPUT").Cells(cnt + 2, 12).Value = Sheets("META").Cells(1, 12).Value
         Sheets("OUTPUT").Cells(cnt + 2, 12).value = NCRImportTool.qaamtTextBox.value
         
    If Not NCRImportTool.adjTextBox.value = "" Then
        
        cnt = cnt + 1
        
        'INVOICE_NO
         Sheets("OUTPUT").Cells(cnt + 2, 1).value = NCRImportTool.qainvTextBox.value
        
         'PO_NO
         Sheets("OUTPUT").Cells(cnt + 2, 2).value = Sheets("META").Cells(1, 2).value
         
         'VENDOR_ID
         Sheets("OUTPUT").Cells(cnt + 2, 3).value = Sheets("META").Cells(1, 3).value
         
         'POSTING_DATE
         Sheets("OUTPUT").Cells(cnt + 2, 4).value = Sheets("META").Cells(1, 4).value
         
         'CREATED_DATE
         Sheets("OUTPUT").Cells(cnt + 2, 5).value = invDate
         
         'DUE_DATE
         Sheets("OUTPUT").Cells(cnt + 2, 6).value = invDate
         
         'DESCRIPTION
         'Sheets("OUTPUT").Cells(cnt + 2, 7).Value = Sheets("META").Cells(1, 7).Value
         Sheets("OUTPUT").Cells(cnt + 2, 7).value = Sheets("META").Cells(1, 7).value & " QA SERVICES"
         
         'LINE_NO
         'Sheets("OUTPUT").Cells(cnt + 2, 8).Value = Sheets("META").Cells(1, 8).Value
         
         'MEMO
         'Sheets("OUTPUT").Cells(cnt + 2, 9).Value = Sheets("META").Cells(1, 9).Value
         Sheets("OUTPUT").Cells(cnt + 2, 9).value = Sheets("META").Cells(1, 7).value & " QA SERVICES - ADJUSTMENT DUE TO OVERPAYMENT"
         
         'ACCT_NO
         Sheets("OUTPUT").Cells(cnt + 2, 10).value = Sheets("META").Cells(1, 10).value
         
         'LOCATION_ID
         Sheets("OUTPUT").Cells(cnt + 2, 11).value = Sheets("META").Cells(1, 11).value
         
         'AMOUNT
         'Sheets("OUTPUT").Cells(cnt + 2, 12).Value = Sheets("META").Cells(1, 12).Value
         Sheets("OUTPUT").Cells(cnt + 2, 12).value = "-" & NCRImportTool.adjTextBox.value
         
    End If
    
    cnt = cnt + 1
    
    NCRImportTool.invTextBox.value = "INV00"
    NCRImportTool.qainvTextBox.value = "INV00"
    NCRImportTool.amtTextBox.value = ""
    NCRImportTool.qaamtTextBox.value = ""
    NCRImportTool.adjTextBox.value = ""
    
noSelection:
    
End Sub

Private Sub errorNoSelection()
    
    Application.StatusBar = "ERROR" 'show 'ERROR' in status bar
    MsgBox "Please select a name", vbCritical, "Selection required"
    Application.StatusBar = False 'restore default status bar text

End Sub

Private Sub exportButton_Click()

    'Call EXPORTCSV
    
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet: Set ws = wb.Sheets("OUTPUT")
    Dim Path As String
    Path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Exporting..."
    
    ws.Visible = xlSheetVisible
    ws.Copy
    ActiveWorkbook.SaveAs Filename:=Path & Format(Now(), "MM-DD-YY") & " NCRIMPORT", FileFormat:=xlCSV, CreateBackup:=False
    ActiveWorkbook.Close
    ws.Visible = xlSheetHidden

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Unload Me
    
    MsgBox "Export complete. File saved to desktop"

End Sub

Private Sub UserForm_Initialize()

    'declare variables and data types
    Dim LR As Long
    Dim Lrow As Long
    Dim Test As New Collection
    Dim i As Single
    Dim value As Variant
    Dim temp As Range
    
    Sheets("META").Cells(1, 1).value = "" 'reset lookup cell to blank
    
    NCRImportTool.invTextBox.value = "INV00"
    NCRImportTool.qainvTextBox.value = "INV00"
    
    On Error Resume Next 'ignore errors in macro and continue with next line

    Set temp = Worksheets("META").ListObjects("LOOKUP").ListColumns(1).Range 'save values from first column in 'accts' table to temp object
 
    'iterate through values in first column except header value
    For i = 2 To temp.Cells.Rows.Count
         If Len(temp.Cells(i)) > 0 And Not temp.Cells(i).EntireRow.Hidden Then 'save value to test if the number of characters are more than 0 and the row is not hidden
           Test.Add temp.Cells(i), CStr(temp.Cells(i))
         End If
    Next i 'continue with next value
 
    acctsListBox.Clear 'delete all items in listbox

    For Each value In Test 'iterate through all values saved to test
        acct = value
        nameListBox.AddItem value 'add values to list box
    Next value 'continue with next value

    Set Test = Nothing 'delete object test
    
    invDate = CalendarForm.GetDate
    
    Sheets("META").Cells(1, 5).value = invDate
    
End Sub



