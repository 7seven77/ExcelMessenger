Sub sendMessage()
    Dim url As String
    Dim name, message As String
    
    url = "http://7seven77.000webhostapp.com/send.php"

    name = Range("B3").Value
    message = Range("E7").Value
    
    url = url + "?name=" + name + "&message=" + message
    
    result = request(url)
    
    Debug.Print result
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


