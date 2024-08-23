Attribute VB_Name = "Refactoring"
Sub rafactoring_macro()
    
    ' ///// PART 0 /////
    ' Disclaimers, etc.
    Dim disclaimerAnswer As VbMsgBoxResult
    disclaimerAnswer = MsgBox("This is a disclaimer. [INSERT DETAILS HERE]. Do you wish to continue?", vbYesNo + vbQuestion, "Disclaimer")
    
    If disclaimerAnswer <> vbYes Then
        Exit Sub
    End If
    
    ' ///// PART 1 /////
    ' First, prompt the user and check that they have copied their cable schedule
    ' into the correct sheet. If they have not, exit. If they have, continue.
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Have you entered your cable schedule into the 'ITR LIST' spreadsheet?", vbYesNo + vbQuestion, "Cable Schedule Entry")
    
    If answer <> vbYes Then
        MsgBox "Please enter your cable schedule into the 'ITR LIST' spreadsheet.", vbExclamation, "Reminder"
        Exit Sub
    End If
    
    ' ///// PART 2 /////
    ' Find out the details of the project. Job number and name of the job
    
    Dim restartProjectDetails As Boolean
    Do
        restartProjectDetails = False
        
        ' Prompt the user to enter the project job number
        Dim projectJobNumber As String
        If Not GetUserInput("Please enter the project job number:", "Project Job Number", projectJobNumber) Then
            Exit Sub
        End If
         
        ' Prompt the user to enter the project name
        Dim projectName As String
        If Not GetUserInput("Please enter the project name:", "Project Name", projectName) Then
            Exit Sub
        End If
        
        ' Confirm the details with the user
        Dim projectDetails As String
        projectDetails = "Project Job Number: " & projectJobNumber & vbCrLf & "Project Name: " & projectName
        answer = MsgBox("You entered the following project details:" & vbCrLf & projectDetails & vbCrLf & "Is this information correct?", vbYesNo + vbQuestion, "Project Confirmation")
        
        If answer <> vbYes Then
            restartProjectDetails = True
        End If
    Loop Until Not restartProjectDetails
    
    ' ///// PART 3 /////
    ' Get the user to tell the program where the important data is in the cable schedule.
    Dim restartProcess As Boolean
    Do
        restartProcess = False
        
        ' Define an array with a list of strings for all parameters which have to be entered
        Dim cableParameters As Variant
        cableParameters = Array("CABLE NUMBER", "CABLE START", "CABLE FINISH", "CABLE TYPE", "CABLE SIZE (mm^2)", "CABLE LENGTH (m)", "CORES")
        
        ' Create a collection to store the user inputs
        Dim columnNames As Collection
        Set columnNames = New Collection
        
        Dim columnType As Variant
        Dim columnName As String
        ' Loop through the array and get the valid column name for each type
        For Each columnType In cableParameters
            If Not GetValidColumnName(CStr(columnType), columnName) Then
                Exit Sub
            End If
            
            ' Add the valid column name to the collection
            columnNames.Add columnName, CStr(columnType)
        Next columnType
        
        ' Confirmation message with "Restart" button
        Dim msg As String
        msg = "You entered the following column names:" & vbCrLf
        Dim colType As Variant
        For Each colType In cableParameters
            msg = msg & colType & ": " & columnNames(CStr(colType)) & vbCrLf
        Next colType
        
        answer = MsgBox(msg & vbCrLf & "Is this information correct?", vbYesNo + vbQuestion, "Column Confirmation")
        If answer <> vbYes Then
            restartProcess = True
        End If
    Loop Until Not restartProcess
    
    
    
    
    ' ///// PART 4 /////
    ' Ask the user how they would like to save the documents
    '   1 - Save all ITR sheets as separate PDF documents
    '   2 - Save all ITR sheets in the same PDF document
    '   3 - Save all ITR sheets from a certain type in the same document. Eg. INSTRUMENT ITRs in one PDF,
    '       CONTROL ITRs in another PDF, etc.
    Dim saveOption As Integer
    With UserForm1
        .Show
        saveOption = CInt(.Tag) ' Retrieve the selected option as an integer
    End With
    
    If saveOption = 0 Then
        Exit Sub
    End If
    
    
    
    ' ///// PART 5 /////
    ' Ask the user where they would like to save the documents
    Dim savePath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "Select Save Location"
        .AllowMultiSelect = False
        If .Show = -1 Then
            savePath = .SelectedItems(1)
        Else
            MsgBox "No folder selected. Exiting.", vbExclamation, "No Folder Selected"
            Exit Sub
        End If
    End With
    
    MsgBox "Documents will be saved to: " & savePath, vbInformation, "Save Location"
    
    
    
    
    ' ///// PART 6 /////
    ' Save the documents
    Select Case saveOption
        Case 1
            MsgBox "You chose to save all ITR sheets as separate PDF documents.", vbInformation, "Save Option"
            
        Case 2
            MsgBox "You chose to save all ITR sheets in the same PDF document.", vbInformation, "Save Option"
            
        Case 3
            MsgBox "You chose to save all ITR sheets from different cable TYPEs in different documents.", vbInformation, "Save Option"
            
        Case Else
            MsgBox "Please start this process again.", vbExclamation, "Invalid Option"
            Exit Sub
    End Select
    
End Sub

Function GetUserInput(prompt As String, title As String, ByRef userInput As String) As Boolean
    userInput = InputBox(prompt, title)
    If userInput = "" Then
        GetUserInput = False
    Else
        GetUserInput = True
    End If
End Function

Function GetValidColumnName(columnType As String, ByRef columnName As String) As Boolean
    Dim isValid As Boolean
    Dim prompt As String
    Dim title As String
    
    prompt = "Please enter the column containing the " & columnType & ":"
    title = columnType & " Column"
    
    Do
        columnName = InputBox(prompt, title)
        
        ' If the user pressed "Cancel" or closed the input box, exit the loop
        If columnName = "" Then
            GetValidColumnName = False
            Exit Function
        End If
        
        ' Check if input contains only alphabetical characters
        isValid = True
        Dim i As Integer
        For i = 1 To Len(columnName)
            If Not (Mid(columnName, i, 1) Like "[A-Za-z]") Then
                isValid = False
                Exit For
            End If
        Next i
        
        If Not isValid Then
            MsgBox "Invalid input. Please enter only alphabetical characters.", vbExclamation, "Invalid Input"
        End If
    Loop Until isValid
    
    GetValidColumnName = True
End Function

