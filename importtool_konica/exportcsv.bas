Attribute VB_Name = "Module1"
Sub EXPORTCSV()
Attribute EXPORTCSV.VB_ProcData.VB_Invoke_Func = "x\n14"

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet: Set ws = wb.Sheets("OUTPUT")
    Dim Path As String
    Path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Exporting..."
    
    ws.Visible = xlSheetVisible
    ws.Copy
    ActiveWorkbook.SaveAs Filename:=Path & Format(Now(), "MM-DD-YY") & " Konica Import", FileFormat:=xlCSV, CreateBackup:=False
    ActiveWorkbook.Close
    ws.Visible = xlSheetHidden

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    MsgBox "Export Complete. File Saved to Desktop"

End Sub
