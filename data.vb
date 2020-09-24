' Should match the maximum number of messages being used in your db
Function getMaximumNumberOfMessages() As Integer
    getMaximumNumberOfMessages = 10
End Function

' Functions to get needed data from the spreadsheet

' Change these functions so they return the cells
' Where the data is located

Function getSenderCell() As String
'   This should be the cell where the sender name is
    getSenderCell = "B3"
End Function

Function getRecipientCell() As String
'   This should be the cell where the recipient name is
    getRecipientCell = "B8"
End Function

Function getMessageCell() As String
'   This should be the cell where the message is
    getMessageCell = "G8"
End Function

Function getMessagesColumn() As String
'   The column which the messages should be displayed in
    getMessagesColumn = "G"
End Function

Function getMessagesStartRow() As Integer
'   The row which the messages should start being displayed on
    getMessagesStartRow = 14
End Function

' Get the value from a cell
Function getCell(position As String) As String
    getCell = Range(position).Value
End Function

Function getSender() As String
    getSender = getCell(getSenderCell())
End Function

Function getRecipient() As String
    getRecipient = getCell(getRecipientCell())
End Function

Function getMessage() As String
    getMessage = getCell(getMessageCell())
End Function

