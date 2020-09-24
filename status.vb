' Status cell subroutines and functions

Function getStatusCell() As String
    getStatusCell = "G3"
End Function

Function getStatus() As String
    getStatus = getCell(getStatusCell())
End Function

Sub setStatus(newValue As String)
    Range(getStatusCell()).value = newValue
End Sub

Sub updateStatus(newValue As String)
    Call setStatus(newValue)
    
    Dim original As String
    original = getStatus()
    
    Call wait
    
    If getStatus() = original Then
        Call setStatus("")
    End If
    
End Sub

Sub wait()
    Dim waitTime As Date
    waitTime = Now() + TimeValue("00:00:05")
    While Now() < waitTime
        DoEvents
    Wend
End Sub

Sub test()
    Call updateStatus("j")
End Sub
