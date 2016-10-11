'Save file by MsgBox
Sub saveFile()
'ask user to save the file
    If MsgBox("Do you want to save the file?", vbYesNo, "Save?") = vbYes Then
        ThisWorkbook.Save
    Else: Exit Sub
    End If
End Sub
