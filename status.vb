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
    
    Dim WaitTime As Date
    WaitTime = Now() + TimeValue("00:00:05")
    While Now() < WaitTime
        DoEvents
    Wend
    
    If getStatus() = original Then
        Call setStatus("")
    End If
    
End Sub
