VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AccountForm 
   Caption         =   "AEP Import Tool"
   ClientHeight    =   1290
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "AccountForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "AccountForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub acctBox_AfterUpdate()
    Me.acctBox.value = Format(Me.acctBox.value, "###-###-###-#-#")
End Sub

Private Sub amtBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        Application.ScreenUpdating = False
        Application.StatusBar = "Exporting..."
             
        'INVOICE_NO
        'Sheets("OUTPUT").Cells(line, 1).value =
        
        'PO_NO
        Sheets("OUTPUT").Cells(line, 2).value = AccountForm.acctBox.value
        
        'VENDOR_ID
        Sheets("OUTPUT").Cells(line, 3).value = "V-001415"
        
        'POSTING_DATE
        'Sheets("OUTPUT").Cells(line, 4).value =
        
        'CREATED_DATE
        Sheets("OUTPUT").Cells(line, 5).value = invDate
        
        'DUE_DATE
        Sheets("OUTPUT").Cells(line, 6).value = duDate
        
        'DESCRIPTION
        Sheets("OUTPUT").Cells(line, 7).value = desc
        
        'LINE_NO
        Sheets("OUTPUT").Cells(line, 8).value = 1
        
        'MEMO
        Sheets("OUTPUT").Cells(line, 9).value = desc
        
        'ACCT_NO
        Sheets("OUTPUT").Cells(line, 10).value = 6450
        
        'LOCATION_ID
        
        Sheets("OUTPUT").Cells(line, 11).value = prop
        
        'AMOUNT
        Sheets("OUTPUT").Cells(line, 12).value = AccountForm.amtBox.value
        
        AccountForm.acctBox.value = ""
        AccountForm.amtBox.value = ""
        
        Application.ScreenUpdating = True
        Application.StatusBar = False
        
        line = line + 1
        
    End If
        
    If KeyCode = vbKeyEscape Then
        
        Call Export
        
    End If
    
End Sub
Private Sub acctBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = vbKeyEscape Then
        
        Call Export
        
    End If
    
End Sub



Private Sub Export()
    
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet: Set ws = wb.Sheets("OUTPUT")
    Dim Path As String
    Path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Exporting..."
    
    ws.Visible = xlSheetVisible
    ws.Copy
    ActiveWorkbook.SaveAs Filename:=Path & Format(Now(), "MM-DD-YY") & " AEP Import", FileFormat:=xlCSV, CreateBackup:=False
    ActiveWorkbook.Close
    ws.Visible = xlSheetHidden

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Unload Me
    
    MsgBox "Export complete. File saved to desktop"

End Sub

Private Sub UserForm_Initialize()
    prop = Sheets("META").Cells(y, 1).value
    Me.acctBox.value = ""
    Me.amtBox.value = ""
End Sub
