Sub getFromWebsite()
    Dim url As String
    
    url = "http://www.7seven77.infinityfreeapp.com/"
    
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


