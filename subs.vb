' Call these subroutines to perform actions
' These can be customised to display and format data
' However you need

Sub showMessages()
    Call clearMessages
    
    Dim source As String
    Dim messages() As String
    
    source = getMessagesRequest()
    If source = "No results" Then Exit Sub
    
    messages = Split(source, ",,")

    ReDim Preserve messages(UBound(messages) - 1)
    
    For Each message In messages
        Debug.Print message
    Next
    
    Dim letter As String
    Dim number As Integer
    Dim messageContents() As String
    
    letter = getMessagesColumn()
    number = getMessagesStartRow()
    For Each message In messages
        messageContents = Split(message, ",")
        Range(letter & number).Value = messageContents(2)
        number = number + 1
    Next
    Debug.Print "Messages output"
End Sub

Sub sendMessage()
    Debug.Print "sending message"
    Dim response As String
    response = sendMessageRequest()
    Debug.Print response
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
        Range(letter & number).Value = ""
    Next number

End Sub
