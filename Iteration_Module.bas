Attribute VB_Name = "Iteration_Module"
Function Iteration(number As Integer, position As Range) As Variant
    Dim iCount As Integer
    Dim iArray() As Variant
    Dim strA As Variant
    
    ReDim iArray(number)
    
    iCount = 0
    For Each strA In position
        If iCount < number Then
            iCount = iCount + 1
            iArray(iCount) = CVar(strA)
        Else
            Exit For
        End If
    Next strA
    Iteration = iArray
End Function
