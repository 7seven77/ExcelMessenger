' Should match the maximum number of messages being used in your db
Function getMaximumNumberOfMessages() As Integer
    getMaximumNumberOfMessages = 20
End Function

' Functions to get needed data from the spreadsheet

' Change these functions so they return the cells
' Where the data is located

Function getSenderCell() As String
'   Change this as needed
'   This should be the cell where the sender name is
    getSenderCell = "B3"
End Function

Function getRecipientCell() As String
'   Change this as needed
'   This should be the cell where the recipient name is
    getRecipientCell = "B8"
End Function

Function getMessageCell() As String
'   Change this as needed
'   This should be the cell where the message is
    getMessageCell = "G8"
End Function

Function getMessagesColumn() As String
'   Change this as needed
'   The column which the messages should be displayed in
    getMessagesColumn = "G"
End Function

Function getMessagesStartRow() As Integer
'   Change this as needed
'   The row which the messages should start being displayed on
    getMessagesStartRow = 14
End Function

' Get the value from a cell
Function getCell(position As String) As String
    getCell = Range(position).value
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

' Check a cell is not empty
' and does not contain forbidden characters
Function isCellValid(cell As String) As Boolean
    Dim value As String
    value = getCell(cell)
    isCellValid = Len(value) > 0 And InStr(value, getSeparator()) = 0
End Function

Function isSenderValid() As Boolean
    isSenderValid = isCellValid(getSenderCell())
    If Not isSenderValid Then
        Call updateStatus("Invalid sender ID")
    End If
End Function

Function isRecipientValid() As Boolean
    isRecipientValid = isCellValid(getRecipientCell())
    If Not isRecipientValid Then
        Call updateStatus("Invalid recipient ID")
    End If
End Function

Function isMessageValid() As Boolean
    isMessageValid = isCellValid(getMessageCell())
    If Not isMessageValid Then
        Call updateStatus("Invalid message")
    End If
End Function

