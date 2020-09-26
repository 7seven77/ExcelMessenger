' Use these functions to make requests

' When messages are recieved they are split with this string
' Each message is split by two concurrent strings
' The message info is split by one
Function getSeparator() As String
    getSeparator = "¬&£@*^%"
End Function

' The base URL from which requests are made
Function getBaseURL() As String
    getBaseURL = "http://7seven77.000webhostapp.com"
End Function

' Send a message to the website
Function sendMessageRequest() As String
    Dim url As String
    Dim sender, message As String
    
    url = getBaseURL() + "/send.php"
    
'   Obtain the cell values
    sender = getSender()
    message = getMessage()
    recipient = getRecipient()
    
'   Create the full url
    url = url + "?sender=" + sender + "&recipient=" + recipient + "&message=" + message
    Debug.Print url
'   Make the request to send a message
    sendMessageRequest = request(url)
End Function

' Get messages from the website
Function getMessagesRequest() As String
    Dim url As String
    Dim sender As String
    
    url = getBaseURL() + "/receive.php"
    
'   Obtain the cell values
    sender = getSender()
    recipient = getRecipient()
    
'   Create the full url
    url = url + "?sender=" + sender + "&recipient=" + recipient + "&random=" & Rnd()
    Debug.Print url
    
'   Make the request to send a message
    Debug.Print request(url)

    getMessagesRequest = request(url)
End Function

' Make a GET request to the website
' Takes in the URL that you would like to go
' Returns the contents of that page
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

