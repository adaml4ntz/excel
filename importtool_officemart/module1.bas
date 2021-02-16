Attribute VB_Name = "Module1"
Private Sub OFFICEMART()

    Application.ScreenUpdating = False
    Application.StatusBar = "Loading..."
   
    ActiveSheet.Name = "INPUT"
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "OUTPUT"
    
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim ws As Worksheet: Set ws = wb.Sheets("OUTPUT")
    
    Range("A1").Value = "INVOICE_NO"
    Range("B1").Value = "PO_NO"
    Range("C1").Value = "VENDOR_ID"
    Range("D1").Value = "POSTING_DATE"
    Range("E1").Value = "CREATED_DATE"
    Range("F1").Value = "DUE_DATE"
    Range("G1").Value = "DESCRIPTION"
    Range("H1").Value = "LINE_NO"
    Range("I1").Value = "MEMO"
    Range("J1").Value = "ACCT_NO"
    Range("K1").Value = "LOCATION_ID"
    Range("L1").Value = "AMOUNT"
    
    Sheets(1).Select
    Range("K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C2").Formula = "=IF(ISBLANK(A2),"""",""O0186"")"
    Range("C2:C500").FillDown
    Range("D2").Formula = "=IF(ISBLANK(A2),"""",TODAY())"
    Range("D2:D500").FillDown
    Range("F2").Formula = "=IF(ISBLANK(A2),"""",TODAY())"
    Range("F2:F500").FillDown
    Sheets(1).Select
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("E2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("G2").Formula = "=IF(ISBLANK(A2),"""",""OFFICE SUPPLIES"")"
    Range("G2:G500").FillDown
    'Range("H2").Formula = "=IF(ISBLANK(A2),"""",""1"")"
    'Range("H2:H500").FillDown
    
    For j = 2 To 500
        If Not Range("A" & j).Value = "" Then
            If Range("A" & j).Value = Range("A" & (j - 1)).Value Then
                k = k + 1
                Range("H" & j).Value = k
            Else
                Range("H" & j).Value = 1
                k = 1
            End If
        End If
    Next
    
    
    'Range("I2").Formula = "=IF(ISBLANK(G2),"""",G2)"
    'Range("I2:I500").FillDown
    Sheets(1).Select
    Range("R2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("I2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Formula = "=IF(ISBLANK(A2),"""",""6311"")"
    Range("J2:J500").FillDown
    Sheets(1).Select
    Range("M2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("K2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets(1).Select
    Range("AD2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("L2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    'Search for Sanitizer
    Range("N2").Formula = "=IF(A2="""","""",IFERROR(IF(SEARCH(""Sanitizer"",INPUT!R2)>0,TRUE),FALSE))"
        Range("N2:N500").FillDown
    
    Dim r As Long
    
    For r = 2 To 500
        If Cells(r, 14).Value = True Then
            Cells(r, 7).Value = "COVID SUPPLIES"
            Cells(r, 10).Value = "8300"
        End If
    Next
    
    Range("N2").Value = ""
        Range("N2:N500").FillDown
        
    'Search for "Mask"
        
        Range("N2").Formula = "=IF(A2="""","""",IFERROR(IF(SEARCH(""Mask"",INPUT!R2)>0,TRUE),FALSE))"
        Range("N2:N500").FillDown
    
    For r = 2 To 500
        If Cells(r, 14).Value = True Then
            Cells(r, 7).Value = "COVID SUPPLIES"
            Cells(r, 10).Value = "8300"
        End If
    Next
    
    Range("N2").Value = ""
        Range("N2:N500").FillDown
    
    '---------------------------------------------------
    
    Range("A1:L500").NumberFormat = "General"
    Range("J1:J500").HorizontalAlignment = xlLeft
    Range("D2:F500").NumberFormat = "mm/dd/yyyy"
    
    Range("A1").Select
    
    ws.Visible = xlSheetVeryHidden
    
    Call EXPORTCSV
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
End Sub

Private Sub EXPORTCSV()

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet: Set ws = wb.Sheets("OUTPUT")
    Dim Path As String
    Path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Exporting..."
    
    ws.Visible = xlSheetVisible
    ws.Copy
    ActiveWorkbook.SaveAs Filename:=Path & Format(Now(), "MM-DD-YY") & " OfficeMart Import", FileFormat:=xlCSV, CreateBackup:=False
    ActiveWorkbook.Close
    ws.Visible = xlSheetHidden

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    MsgBox "Export Complete. File Saved to Desktop"

End Sub


