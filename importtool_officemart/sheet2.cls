VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
 
  Dim lastAction As String
  Dim w As Worksheet
  Dim b As Button

If Not Intersect(Target, Range("A1:Z500" & ThisWorkbook.Worksheets(1).UsedRange.Rows.Count)) Is Nothing Then
        Set w = ActiveSheet
        Set b = w.Buttons.Add(5, 5, 70, 18.75)      ' left, top, width, height
        
        b.Select
        With Selection.Font
            .Name = "Courier New"
            .FontStyle = "Bold"
            .Size = 12
        End With
        b.OnAction = "OFFICEMART"          ' a sub name in module
        b.Characters.Text = "Export"  ' text on the button
    End If

    Application.Run "Module1.OFFICEMART"

End Sub

