' Call these subroutines to perform actions
' These can be customised to display and format data
' However you need

Sub showMessages()
    Dim source As String
    Dim messages() As String
    
    source = getMessages()
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
