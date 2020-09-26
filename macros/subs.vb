' Call these subroutines to perform actions
' These can be customised to display and format data
' However you need

' Display all the messages to the user (button command)
Sub showMessages()
    If Not isSenderValid Or Not isRecipientValid Then
        Exit Sub
    End If
    
'   Update status and reset message cells ready for overwrite
    Call setStatus("Fetching messages")
    Call clearMessages
    
    Dim source As String
    Dim messages() As String

'   Get the response from the website
    source = getMessagesRequest()

'   Update status if there are no messages
    If source = "No results" Then
        Call updateStatus("No messages")
        Exit Sub
    End If
    
'   Split the return'd string into an array
'   Using the separator to differentiate between messages
    Dim splitter As String
    splitter = getSeparator() + getSeparator()
    messages = Split(source, splitter)
    
'   Remove blank space at end of list
    ReDim Preserve messages(UBound(messages) - 1)
    
'   Display all the messages using the show sub
    Dim count As Integer
    For count = 0 To UBound(messages)
        Call showMessage(messages(count), count)
    Next
    
    Call updateStatus("Messages received")
End Sub

' Send a message (button command)
Sub sendMessage()
'   Check if the inputs are all valid
    If Not isSenderValid Or Not isRecipientValid Or Not isMessageValid Then
        Exit Sub
    End If
    
    Call setStatus("sending message")

'   Get the response from sending
    Dim response As String
    response = sendMessageRequest()
    Call setStatus(response)

'   If there wasn't an error, show the messages
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
' offset is how recent hte message is (0 being the most recent)
Sub showMessage(message As String, offset As Integer)
    Dim messageContents() As String
    Dim letter As String
    Dim number As Integer
    
'   Split the string into the message components
    messageContents = Split(message, getSeparator())
    
    letter = getMessagesColumn()
    number = getMessagesStartRow() + offset

'   Get the message cell and update the value to show the message
    Dim cell As Range
    Set cell = Range(letter & number)
    
    cell.value = messageContents(2)
  
'   If the message was sent by the person viewing
'   Align the message to the right
    Dim sender As String
    sender = getSender()
    
    If messageContents(0) = sender Then
        cell.HorizontalAlignment = xlRight
        cell.value = messageContents(2) & "  <"
    Else
        cell.HorizontalAlignment = xlLeft
        cell.value = ">  " & messageContents(2)
    End If
    
End Sub
