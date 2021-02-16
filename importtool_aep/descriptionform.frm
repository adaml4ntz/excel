VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DescriptionForm 
   Caption         =   "Service Period"
   ClientHeight    =   525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2730
   OleObjectBlob   =   "DescriptionForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "DescriptionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    
    DescriptionForm.descBox.value = desc

End Sub

Private Sub descBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Me.Hide
        PropertyForm.Show
    End If
End Sub

