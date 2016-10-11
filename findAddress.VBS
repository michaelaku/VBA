
'Get address (range, row, column) of searching value
' the value is searched from set position
Sub search_for_value()
'DEFINITIONS
neededValue = "string"
position = "A1"
'END OF DEFINITIONS
    'use function to find address of value
    foundAddress = Application.Run("findValue", neededValue, position)
    If Not foundAddress = "" Then
        foundRow = Range(foundAddress).Row
        foundColumn = Range(foundAddress).Column
        MsgBox ("RANGE: " & foundAddress & "   ROW: " & foundRow & "  COLUMN: " & foundColumn)
    Else: Exit Sub
    End If
End Sub
'find address of value
Function findValue(value As String, position As String)
    Set findValue = ActiveSheet.Cells.Find(What:=value, After:=ActiveSheet.Range(position), LookIn:=xlValues)
    If findValue Is Nothing Then
       MsgBox ("The file is missing needed value " & Chr(34) & value & Chr(34) & "!")
       findValue = ""
    Else:
        findValue = findValue.Address
    End If
End Function
