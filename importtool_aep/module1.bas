Attribute VB_Name = "Module1"
Public desc As String
Public lookup As Integer
Public y As Integer
Public line As Integer
Public invDate As Date
Public duDate As Date
Public prop As String

Sub AEPImportTool()
    
    desc = ""
    line = 2
    
    CalendarForm.Caption = "Bill Mailing Date"
    invDate = CalendarForm.GetDate
    CalendarForm.Caption = "Due Date"
    duDate = CalendarForm.GetDate
    
    desc = ((invDate - 29) & "-" & (invDate))
    
    DescriptionForm.Show
    
    desc = DescriptionForm.descBox.value
    
End Sub






















