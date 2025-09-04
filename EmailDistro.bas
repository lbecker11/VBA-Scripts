Attribute VB_Name = "EmailDistro"

Sub UpdateDashboard()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim currentRow As Long
    Dim lastRow As Long
    Dim matchCell As Range
    Dim searchValue As String
    Dim OutlookApp As Object
    Dim MailItem As Object
    Dim emailBody As String
    Dim recipientEmail As String
    Dim ccEmail As String
    Dim globalSubject As String

    ' Set your worksheets
    Set wsSource = ThisWorkbook.Sheets("Dashboard") ' Source sheet
    Set wsTarget = ThisWorkbook.Sheets("Data") ' Target sheet

    ' Connect to Outlook
    On Error Resume Next
    Set OutlookApp = GetObject(class:="Outlook.Application")
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject(class:="Outlook.Application")
    End If
    On Error GoTo 0

    

    ' Start from row 20 in column G
    currentRow = 20
    lastRow = wsSource.Cells(wsSource.Rows.Count, "G").End(xlUp).Row

    ' Loop through each row starting from G20
    Do While currentRow <= lastRow
            
        ' Highlight columns G to J in the current row
        wsSource.Range("G" & currentRow & ":J" & currentRow).Interior.Color = RGB(144, 238, 144) ' Light green

        ' Get the value to search for
        searchValue = wsSource.Cells(currentRow, "G").Value

        ' Search for the value in the target sheet (column A, starting from A2)
        Set matchCell = wsTarget.Range("A2:A" & wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row) _
                        .Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)

        If Not matchCell Is Nothing Then
            ' Update column K in the source sheet
            wsSource.Cells(currentRow, "K").Value = "Created"

            ' Update column E in the target sheet
            wsTarget.Cells(matchCell.Row, "E").Value = "Yes"

            ' Create and save draft email
            recipientEmail = wsTarget.Cells(matchCell.Row, "A").Value
            emailBody = wsTarget.Cells(matchCell.Row, "G").Value
            emailBody = Replace(emailBody, "Central_Coordinator", wsTarget.Cells(matchCell.Row, "K").Value)
            emailBody = Replace(emailBody, "Company_Name", wsTarget.Cells(matchCell.Row, "B").Value)
            ccEmail = wsTarget.Cells(matchCell.Row, "D").Value ' Optional
            ' Set subject
            globalSubject = wsTarget.Cells(matchCell.Row, "I").Value

            Set MailItem = OutlookApp.CreateItem(0)
            With MailItem
                .To = recipientEmail
                .Subject = globalSubject
                .HTMLBody = emailBody
                .CC = ccEmail
                .SentOnBehalfOfName = "placeholder@test.com"
                .Save
            End With

            ' Refresh chart
            Application.CalculateFull
            DoEvents
            wsSource.ChartObjects("Chart 2").Chart.Refresh
        End If

        ' Remove the highlight from columns G to J
        wsSource.Range("G" & currentRow & ":J" & currentRow).Interior.ColorIndex = xlNone
        ' Move to the next row
        currentRow = currentRow + 1
    Loop
    MsgBox "Processing complete!"
End Sub






