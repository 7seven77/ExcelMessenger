' Status cell subroutines and functions

Function getStatusCell() As String
    getStatusCell = "G3"
End Function

Function getStatus() As String
    getStatus = getCell(getStatusCell())
End Function

' Change the status to te string specified
Sub setStatus(newValue As String)
    Range(getStatusCell()).value = newValue
End Sub

' Change the status to the string specified
' It will then be reset after a short time period
' Use setStatus if you want the change to be permanent
Sub updateStatus(newValue As String)
    Call setStatus(newValue)
    
    Dim original As String
    original = getStatus()
    
    Call wait("03")
    
    ' If the status was changed during the wait period
    ' Do not reset it
    If getStatus() = original Then
        Call setStatus("")
    End If
    
End Sub

' Pause execution of a subroutine for the specified number of seconds
Sub wait(seconds As String)
    Dim waitTime As Date
    waitTime = Now() + TimeValue("00:00:" & seconds)
    While Now() < waitTime
        DoEvents
    Wend
End Sub

Function allCellsAreValid() As Boolean
    If Not isSenderValid() Then
        Call updateStatus("Invalid sender ID")
        allCellsAreValid = False
    ElseIf Not isRecipientValid() Then
        Call updateStatus("Invalid recipient ID")
        allCellsAreValid = False
    ElseIf Not isMessageValid() Then
        Call updateStatus("Invalid message")
        allCellsAreValid = False
    End If
    allCellsAreValid = True
End Function
