'Macro opens selected file by file from selected folder, but before a file is opened, the name of the file is checked
'If the name is valid, the file is opened, sequentially the file is saved and closed
'If the name is NOT valid, the file is not opened at all and the macro is checking the next selected file for process

Sub macro()
reporttype = "test"
'OPEN MULTIPLE FILES FROM SELECTED FOLDER
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = True
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    If intChoice <> 0 Then
        For i = 1 To Application.FileDialog(msoFileDialogOpen).SelectedItems.Count
            reportName = Application.FileDialog(msoFileDialogOpen).SelectedItems(i)
            'Check whether file is valid for macro process
            If (InStr(reportName, reporttype) = 0) Then
                MsgBox ("File" & reportName & " is not valid")
                GoTo next_file
            Else:
                Set report = Workbooks.Open(reportName)
            'Save and Close report
             report.Close Savechanges:=True
            End If
                
'Get next file
next_file:
        Next i
    End If
End Sub
