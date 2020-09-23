' Use these functions to make requests

Sub sendMessage()
    Dim url As String
    Dim sender, message As String
    
    url = "http://7seven77.000webhostapp.com/send.php"
    
'   Obtain the cell values
    sender = getSender()
    message = getMessage()
    recipient = getRecipient()
    
'   Create the full url
    url = url + "?sender=" + sender + "&recipient=" + recipient + "&message=" + message
    Debug.Print url
'   Make the request to send a message
    result = request(url)
    Debug.Print result
End Sub

Function getMessages() As String
    Dim url As String
    Dim sender As String
    
    url = "http://7seven77.000webhostapp.com/receive.php"
    
'   Obtain the cell values
    sender = getSender()
    recipient = getRecipient()
    
'   Create the full url
    url = url + "?sender=" + sender + "&recipient=" + recipient + "&random=" & Rnd()
    Debug.Print url
    
'   Make the request to send a message
    getMessages = request(url)
End Function

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

