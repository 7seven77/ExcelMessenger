' Functions to get needed data from the spreadsheet

' Change these functions so they return the cells
' Where the data is located

Function getSenderCell() As String
'   This should be the cell where the sender name is
    getSenderCelll = "B3"
End Function

Function getRecipientCell() As String
'   This should be the cell where the recipient name is
    getRecipientCell = "B5"
End Function

Function getMessageCell() As String
'   This should be the cell where the message is
    getMessageCell = "E7"
End Function

Function getMessagesColumn() As String
'   The column which the messages should be displayed in
    getMessagesColumn = "E"
End Function

Function getMessagesStartRow() As Integer
'   The row which the messages should start being displayed on
    getMessagesStartRow = 13
End Function

' Get the value from a cell
Function getCell(position As String) As String
    getCell = Range(position).Value
End Function

Function getSender() As String
    getSender = getCell("B3")
End Function

Function getRecipient() As String
    getRecipient = getCell("B5")
End Function

Function getMessage() As String
    getMessage = getCell(getMessageCell())
End Function

