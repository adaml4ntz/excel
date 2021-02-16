VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Activate()
     
     ActiveSheet.ScrollArea = "A1:J107"
    
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim KeyCells As Range
    Set KeyCells = Range("B4:H102")
    
    ActiveSheet.ScrollArea = "A1:J107"
    Application.ScreenUpdating = False
    Application.StatusBar = "Loading..."
    Application.EnableEvents = False

    
    If Not Application.Intersect(KeyCells, Range(Target.Address)) Is Nothing Then
        Debug.Print ("called")
        Call EASYENTRY
        Call VERIFY
    End If
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Application.EnableEvents = True

End Sub

Private Sub VERIFY()
    
    Dim x As Integer
    Dim c As Long
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim ws As Worksheet: Set ws = wb.Sheets("OUTPUT")
    
    Dim amt As Double
    
    ws.Visible = xlSheetVisible
    
    x = 4


        For Each Cell In Sheets("OUTPUT").Range("k2:k101")
            If Not IsNumeric(Cell.Value) Then
                If IsEmpty(Cell.Value) Then
                    GoTo Continue
                End If
                'Marion'
                If Cell.Value = "Marion" Then
                
                    amt = Sheets("INPUT").Range("G" & x)
                    Sheets("INPUT").Range("G" & x) = Application.WorksheetFunction.RoundUp(amt * 0.72, 2)
                
                    Sheets("INPUT").Range("B" & x + 1) = Sheets("INPUT").Range("B" & x)
                    Sheets("INPUT").Range("C" & x + 1) = Sheets("INPUT").Range("C" & x)
                    Sheets("INPUT").Range("D" & x + 1) = Sheets("INPUT").Range("D" & x)
                    Sheets("INPUT").Range("E" & x + 1) = Sheets("INPUT").Range("E" & x)
                    Sheets("INPUT").Range("F" & x + 1) = Sheets("INPUT").Range("F" & x)
                    Sheets("INPUT").Range("H" & x + 1) = Sheets("INPUT").Range("H" & x)
                    Sheets("INPUT").Range("G" & x + 1) = Application.WorksheetFunction.RoundDown(amt * 0.28, 2)
                    
                    Sheets("OUTPUT").Range("K" & x - 2) = "20172"
                    Sheets("OUTPUT").Range("K" & x - 1) = "20194"
                    
                    Worksheets("INPUT").Range("B" & ActiveCell.Row + 1).Activate
                    
                    Exit Sub
                End If
                    
                'Area 7 Allocation
                If Cell.Value = "Area 7 Allocation" Then
                    
                    Sheets("OUTPUT").Range("K" & x - 2) = ""
                    Sheets("OUTPUT").Range("M" & x - 2) = "Area 7 Allocation"
                    
                    Exit Sub
                End If
                
                'Staunton
                If Cell.Value = "Staunton" Then
                
                    amt = Sheets("INPUT").Range("G" & x)
                    Sheets("INPUT").Range("G" & x) = Application.WorksheetFunction.RoundUp(amt * 0.5, 2)
                
                    Sheets("INPUT").Range("B" & x + 1) = Sheets("INPUT").Range("B" & x)
                    Sheets("INPUT").Range("C" & x + 1) = Sheets("INPUT").Range("C" & x)
                    Sheets("INPUT").Range("D" & x + 1) = Sheets("INPUT").Range("D" & x)
                    Sheets("INPUT").Range("E" & x + 1) = Sheets("INPUT").Range("E" & x)
                    Sheets("INPUT").Range("F" & x + 1) = Sheets("INPUT").Range("F" & x)
                    Sheets("INPUT").Range("H" & x + 1) = Sheets("INPUT").Range("H" & x)
                    Sheets("INPUT").Range("G" & x + 1) = Application.WorksheetFunction.RoundDown(amt * 0.5, 2)
                    
                    Sheets("OUTPUT").Range("K" & x - 2) = "20154"
                    Sheets("OUTPUT").Range("K" & x - 1) = "20190"
                    
                    Exit Sub
                End If
                
                'Devonshire
                If Cell.Value = "Devonshire" Then
                
                    amt = Sheets("INPUT").Range("G" & x)
                    Sheets("INPUT").Range("G" & x) = Application.WorksheetFunction.RoundUp(amt * 0.65, 2)
                
                    Sheets("INPUT").Range("B" & x + 1) = Sheets("INPUT").Range("B" & x)
                    Sheets("INPUT").Range("C" & x + 1) = Sheets("INPUT").Range("C" & x)
                    Sheets("INPUT").Range("D" & x + 1) = Sheets("INPUT").Range("D" & x)
                    Sheets("INPUT").Range("E" & x + 1) = Sheets("INPUT").Range("E" & x)
                    Sheets("INPUT").Range("F" & x + 1) = Sheets("INPUT").Range("F" & x)
                    Sheets("INPUT").Range("H" & x + 1) = Sheets("INPUT").Range("H" & x)
                    Sheets("INPUT").Range("G" & x + 1) = Application.WorksheetFunction.RoundDown(amt * 0.35, 2)
                    
                    Sheets("OUTPUT").Range("K" & x - 2) = "20197"
                    Sheets("OUTPUT").Range("K" & x - 1) = "20113"
                    
                    Exit Sub
                End If
                
                'Camden
                If Cell.Value = "Camden" Then
                
                    amt = Sheets("INPUT").Range("G" & x)
                    Sheets("INPUT").Range("G" & x) = Application.WorksheetFunction.RoundUp(amt * 0.5, 2)
                
                    Sheets("INPUT").Range("B" & x + 1) = Sheets("INPUT").Range("B" & x)
                    Sheets("INPUT").Range("C" & x + 1) = Sheets("INPUT").Range("C" & x)
                    Sheets("INPUT").Range("D" & x + 1) = Sheets("INPUT").Range("D" & x)
                    Sheets("INPUT").Range("E" & x + 1) = Sheets("INPUT").Range("E" & x)
                    Sheets("INPUT").Range("F" & x + 1) = Sheets("INPUT").Range("F" & x)
                    Sheets("INPUT").Range("H" & x + 1) = Sheets("INPUT").Range("H" & x)
                    Sheets("INPUT").Range("G" & x + 1) = Application.WorksheetFunction.RoundDown(amt * 0.5, 2)
                    
                    Sheets("OUTPUT").Range("K" & x - 2) = "20004"
                    Sheets("OUTPUT").Range("K" & x - 1) = "20005"
                    
                    Exit Sub
                End If
                
                'Scottwood and Western Run
                If Cell.Value = "Scottwood" Then
                
                    amt = Sheets("INPUT").Range("G" & x)
                    Sheets("INPUT").Range("G" & x) = Application.WorksheetFunction.RoundUp(amt * 0.4, 2)
                
                    Sheets("INPUT").Range("B" & x + 1) = Sheets("INPUT").Range("B" & x)
                    Sheets("INPUT").Range("C" & x + 1) = Sheets("INPUT").Range("C" & x)
                    Sheets("INPUT").Range("D" & x + 1) = Sheets("INPUT").Range("D" & x)
                    Sheets("INPUT").Range("E" & x + 1) = Sheets("INPUT").Range("E" & x)
                    Sheets("INPUT").Range("F" & x + 1) = Sheets("INPUT").Range("F" & x)
                    Sheets("INPUT").Range("H" & x + 1) = Sheets("INPUT").Range("H" & x)
                    Sheets("INPUT").Range("G" & x + 1) = Application.WorksheetFunction.RoundDown(amt * 0.6, 2)
                    
                    Sheets("OUTPUT").Range("K" & x - 2) = "20027"
                    Sheets("OUTPUT").Range("K" & x - 1) = "20028"
                    
                    Exit Sub
                End If
                
                'Ridgeview
                If Cell.Value = "Ridgeview" Then
                
                    amt = Sheets("INPUT").Range("G" & x)
                    Sheets("INPUT").Range("G" & x) = Application.WorksheetFunction.RoundUp(amt / 3, 2)
                
                    Sheets("INPUT").Range("B" & x + 1) = Sheets("INPUT").Range("B" & x)
                    Sheets("INPUT").Range("C" & x + 1) = Sheets("INPUT").Range("C" & x)
                    Sheets("INPUT").Range("D" & x + 1) = Sheets("INPUT").Range("D" & x)
                    Sheets("INPUT").Range("E" & x + 1) = Sheets("INPUT").Range("E" & x)
                    Sheets("INPUT").Range("F" & x + 1) = Sheets("INPUT").Range("F" & x)
                    Sheets("INPUT").Range("H" & x + 1) = Sheets("INPUT").Range("H" & x)
                    Sheets("INPUT").Range("G" & x + 1) = Application.WorksheetFunction.Round(amt / 3, 2)
                    
                    Sheets("INPUT").Range("B" & x + 2) = Sheets("INPUT").Range("B" & x)
                    Sheets("INPUT").Range("C" & x + 2) = Sheets("INPUT").Range("C" & x)
                    Sheets("INPUT").Range("D" & x + 2) = Sheets("INPUT").Range("D" & x)
                    Sheets("INPUT").Range("E" & x + 2) = Sheets("INPUT").Range("E" & x)
                    Sheets("INPUT").Range("F" & x + 2) = Sheets("INPUT").Range("F" & x)
                    Sheets("INPUT").Range("H" & x + 2) = Sheets("INPUT").Range("H" & x)
                    Sheets("INPUT").Range("G" & x + 2) = Application.WorksheetFunction.RoundDown(amt / 3, 2)
                    
                    Sheets("OUTPUT").Range("K" & x - 2) = "20900" 'Ridgeview 1 20160
                    Sheets("OUTPUT").Range("K" & x - 1) = "20900" 'Ridgeview II 20161
                    Sheets("OUTPUT").Range("K" & x) = "20900"     'Ridgeview II 20174
                    
                    Exit Sub
                End If
Continue:
            End If
            x = x + 1
        Next
    ws.Visible = xlSheetVeryHidden
End Sub

Private Sub EASYENTRY()
    
    Dim r As Long, c As Long
    Dim Tbl As ListObject
    Set Tbl = Worksheets("INPUT").ListObjects("INPUT")
    
    For r = 2 To Tbl.DataBodyRange.Rows.Count
        If Tbl.DataBodyRange.Cells(r - 1, 1).Value = Tbl.DataBodyRange.Cells(r, 1).Value Then
            If Not Tbl.DataBodyRange.Cells(r, 1).Value = "" Then
                For c = 1 To 5
                    Tbl.DataBodyRange.Cells(r, c).Value = Tbl.DataBodyRange.Cells(r - 1, c).Value
                Next
            End If
        End If
    Next
End Sub

Private Sub SHOWLISTS()

    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim ws As Worksheet: Set ws = wb.Sheets("LISTS")
    Dim answer As Integer
    
    If wb.Sheets("LISTS").Visible = True Then
        Call HIDELISTS
    Else
        answer = MsgBox("Lists are linked to a central database. Only make changes if you know what you are doing.", vbOKCancel + vbCritical, "Warning!")

        If answer = vbOK Then
            Application.ScreenUpdating = False
            ws.ScrollArea = "A1:J150"
            ws.Visible = xlSheetVisible
            Application.ScreenUpdating = True
        End If
    End If
    
End Sub

Private Sub HIDELISTS()

    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim ws As Worksheet: Set ws = wb.Sheets("LISTS")
    
    Application.ScreenUpdating = False
    ws.ScrollArea = "A1:J150"
    ws.Visible = xlSheetVeryHidden
    Application.ScreenUpdating = True
    
End Sub
