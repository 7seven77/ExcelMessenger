Sub sendMessage()
'   Change these values to change where the
'   Information is taken from

    Dim nameCell, messageCell As String
    nameCell = "B3"
    messageCell = "E7"
    
    Dim url As String
    Dim name, message As String
    
    url = "http://7seven77.000webhostapp.com/send.php"
    
'   Obtain the cell values
    name = Range(nameCell).Value
    message = Range(messageCell).Value
    
'   Create the full url
    url = url + "?name=" + name + "&message=" + message

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


