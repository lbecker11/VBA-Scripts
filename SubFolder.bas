Attribute VB_Name = "SubFolder"
Sub CreateOutlookFolders()
    Dim olApp As Object
    Dim olNs As Object
    Dim rootFolder As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim parentFolderName As String
    Dim newFolderName As String
    Dim parentFolder As Object
    Dim newFolder As Object
    Dim folderDict As Object ' Dictionary to map folder names to Outlook folder objects
    Dim f As Object
    Dim folderFound As Boolean

    ' Initialize Outlook
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    If olApp Is Nothing Then
        MsgBox "Outlook could not be opened.", vbCritical
        Exit Sub
    End If
    
    Set olNs = olApp.GetNamespace("MAPI")
    Set rootFolder = olNs.GetDefaultFolder(6).Parent ' Top of mailbox
    Set ws = ThisWorkbook.Sheets(1)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Set folderDict = CreateObject("Scripting.Dictionary")
    folderDict.Add rootFolder.Name, rootFolder

    Debug.Print "Starting folder creation..."
    Debug.Print "Root folder: " & rootFolder.Name

    ' Loop through Excel rows
    For i = 2 To lastRow
        parentFolderName = Trim(ws.Cells(i, 1).Value)
        newFolderName = Trim(ws.Cells(i, 2).Value)

        If parentFolderName <> "" And newFolderName <> "" Then
            Debug.Print String(50, "-")
            Debug.Print "Row " & i & ": Parent='" & parentFolderName & "' | New='" & newFolderName & "'"

            ' Try to get parent folder from dictionary
            If folderDict.Exists(parentFolderName) Then
                Set parentFolder = folderDict(parentFolderName)
                Debug.Print "Found parent folder in dictionary: " & parentFolder.Name
            Else
                Set parentFolder = FindFolder(rootFolder, parentFolderName)
                If Not parentFolder Is Nothing Then
                    folderDict.Add parentFolderName, parentFolder
                    Debug.Print "Found parent folder via search: " & parentFolder.Name
                Else
                    Debug.Print "? Parent folder NOT FOUND: " & parentFolderName
                    GoTo NextRow
                End If
            End If

            ' List existing subfolders
            Debug.Print "Existing subfolders under '" & parentFolder.Name & "':"
            For Each f In parentFolder.Folders
                Debug.Print " - '" & f.Name & "'"
            Next f

            ' Check if folder already exists manually
            folderFound = False
            For Each f In parentFolder.Folders
                If LCase(Trim(f.Name)) = LCase(Trim(newFolderName)) Then
                    folderFound = True
                    Set newFolder = f
                    Exit For
                End If
            Next

            If Not folderFound Then
                Set newFolder = parentFolder.Folders.Add(newFolderName)
                Debug.Print "? Created folder: '" & newFolderName & "' under '" & parentFolder.Name & "'"
            Else
                Debug.Print "?? Folder already exists: '" & newFolderName & "' under '" & parentFolder.Name & "'"
            End If

            ' Add to dictionary so it can be used as a parent later
            If Not folderDict.Exists(newFolderName) Then
                folderDict.Add newFolderName, newFolder
                Debug.Print "Added '" & newFolderName & "' to folder dictionary."
            End If
        End If
NextRow:
    Next i

    MsgBox "Folders processed. Check Immediate Window (Ctrl+G) for debug info.", vbInformation
End Sub

Function FindFolder(startFolder As Object, folderName As String) As Object
    Dim subFolder As Object
    If LCase(Trim(startFolder.Name)) = LCase(Trim(folderName)) Then
        Set FindFolder = startFolder
        Exit Function
    End If
    
    For Each subFolder In startFolder.Folders
        Set FindFolder = FindFolder(subFolder, folderName)
        If Not FindFolder Is Nothing Then Exit Function
    Next
End Function

