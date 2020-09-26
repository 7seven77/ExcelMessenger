' Call these subroutines to perform actions
' These can be customised to display and format data
' However you need

Sub showMessages()
    If Not isSenderValid Or Not isRecipientValid Then
        Exit Sub
    End If
    
'   Update status and reset message cells ready for overwrite
    Call setStatus("Fetching messages")
    Call clearMessages
    
    Dim source As String
    Dim messages() As String
    
    source = getMessagesRequest()
    If source = "No results" Then
        Call updateStatus("No messages")
        Exit Sub
    End If
    Dim splitter As String
    splitter = getSeparator() + getSeparator()
    messages = Split(source, splitter)

'   Remove blank space at end of list
    ReDim Preserve messages(UBound(messages) - 1)
    
    Dim count As Integer
    For count = 0 To UBound(messages)
        Call showMessage(messages(count), count)
    Next
    
    Call updateStatus("Messages received")
End Sub

Sub sendMessage()
    If Not isSenderValid Or Not isRecipientValid Or Not isMessageValid Then
        Exit Sub
    End If
    
    Call setStatus("sending message")
    
    Dim response As String
    response = sendMessageRequest()
    Call setStatus(response)
    If InStr(1, response, "ERROR") = 0 Then
        Call showMessages
    End If
End Sub

' Remove all messages so that new ones can be displayed
Sub clearMessages()
    Dim letter As String
    Dim start, maximum As Integer
    
    letter = getMessagesColumn()
    start = getMessagesStartRow()
    
'   Maximum number of messages
    maximum = getMaximumNumberOfMessages()
    
    For number = start To (start + maximum)
        Range(letter & number).value = ""
    Next number

End Sub

' Display a messgae
' Change this subroutine to alter how and what is displayed to the user
Sub showMessage(message As String, offset As Integer)
    Dim messageContents() As String
    Dim letter As String
    Dim number As Integer
    
    messageContents = Split(message, getSeparator())
    
    letter = getMessagesColumn()
    number = getMessagesStartRow() + offset
    
    Dim cell As Range
    Set cell = Range(letter & number)
    
    cell.value = messageContents(2)
    
    Dim sender As String
    sender = getSender()
    
'   If the message was sent by the person viewing
'   Align the message to the right

    If messageContents(0) = sender Then
        cell.HorizontalAlignment = xlRight
        cell.value = messageContents(2) & "  <"
    Else
        cell.HorizontalAlignment = xlLeft
        cell.value = ">  " & messageContents(2)
    End If
    
End Sub
