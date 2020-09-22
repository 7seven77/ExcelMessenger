Sub showMessages()

End Sub

Function getMessages() As String
'   Change these values to change where the
'   Information is taken from

    Dim nameCell, recipientCell As String
    nameCell = "B3"
    recipientCell = "B5"
    
    Dim url As String
    Dim name As String
    
    url = "http://7seven77.000webhostapp.com/receive.php"
    
'   Obtain the cell values
    name = Range(nameCell).Value
    recipient = Range(recipientCell).Value
    
'   Create the full url
    url = url + "?sender=" + name + "&recipient=" + recipient
    Debug.Print url
    
'   Make the request to send a message
    getMessages = request(url)
End Function

Sub sendMessage()
'   Change these values to change where the
'   Information is taken from

    Dim nameCell, messageCell, recipientCell As String
    nameCell = "B3"
    recipientCell = "B5"
    messageCell = "E7"
    
    Dim url As String
    Dim name, message As String
    
    url = "http://7seven77.000webhostapp.com/send.php"
    
'   Obtain the cell values
    name = Range(nameCell).Value
    message = Range(messageCell).Value
    recipient = Range(recipientCell).Value
    
'   Create the full url
    url = url + "?sender=" + name + "&recipient=" + recipient + "&message=" + message
    Debug.Print url
'   Make the request to send a message
    result = request(url)
    Debug.Print result

'   Reset the message cell ready for another message
    Range(messageCell).Value = ""
End Sub

Function request(url As String) As String
    Debug.Print "Making Request"
    Dim httprequest As Object
    
    Set httprequest = CreateObject("MSXML2.XMLHTTP")
    With httprequest
        .Open "GET", url, False
        .send
    End With
    
    request = httprequest.responseText
End Function
