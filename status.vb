' Status cell subroutines and functions

Function getStatusCell() As String
    getStatusCell = "G3"
End Function

Sub updateStatus(newValue As String)
    Range(getStatusCell()).Value = newValue
End Sub
