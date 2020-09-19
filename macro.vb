Sub request()
    Dim request As Object
    Dim url As String
    
    url = "http://www.7seven77.000webhostapp.cmo/excelim.php"
    
    Set request = CreateObject("MSXML2.XMLHTTP")
    With request
        .Open "GET", url, False
        .send
    End With
    
    Debug.Print request.responseText
End Sub

