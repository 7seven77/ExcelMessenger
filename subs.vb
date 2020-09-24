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

'   Remove blank space at end of list
    ReDim Preserve messages(UBound(messages) - 1)
    
    Dim count As Integer
    For count = 0 To UBound(messages)
        Call showMessage(messages(count), count)
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

' Display a messgae
' Change this subroutine to alter how and what is displayed to the user
Sub showMessage(message As String, offset As Integer)
    Dim messageContents() As String
    Dim letter As String
    Dim number As Integer
    
    messageContents = Split(message, ",")
    
    letter = getMessagesColumn()
    number = getMessagesStartRow() + offset
    
    Dim cell As Range
    Set cell = Range(letter & number)
    
    cell.Value = messageContents(2)
    
    Dim sender As String
    sender = getSender()
    
'   If the message was sent by the person viewing
'   Align the message to the right

    If messageContents(0) = sender Then
        cell.HorizontalAlignment = xlRight
        cell
    Else
        cell.HorizontalAlignment = xlLeft
    End If
    
End Sub
