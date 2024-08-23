VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ' Initialize the label and combo box
    labelInstruction.Caption = "How do you want the ITR documents to be saved?"
    SaveOptions.AddItem "Save all ITR sheets as separate PDF documents"
    SaveOptions.AddItem "Save all ITR sheets in the same PDF document"
    SaveOptions.AddItem "Save all ITR sheets from different cable TYPEs in different documents"
    SaveOptions.ListIndex = 0 ' Set the default selection to the first item
    Me.Tag = 0 ' Initialize the Tag property to 0 to handle form closure
End Sub

Private Sub buttonOK_Click()
    If SaveOptions.ListIndex = -1 Then
        MsgBox "Please select a save option.", vbExclamation, "No Selection"
    Else
        Me.Tag = SaveOptions.ListIndex + 1 ' Store the selected option as an integer (1, 2, or 3)
        Me.Hide
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then ' If the form is closed using the "X" button
        Me.Tag = 0 ' Set Tag to 0 indicating the form was closed
        Me.Hide ' Hide the form to prevent further processing
        Cancel = True ' Cancel the default close action to avoid errors
    End If
End Sub

