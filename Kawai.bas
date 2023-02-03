Attribute VB_Name = "Kawai"
Option Explicit

Sub Kawaii()
Attribute Kawaii.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveSheet.Shapes.Range(Array("Picture 2")).Select
    If Selection.ShapeRange.SoftEdge.Radius = 0 Then
        Selection.ShapeRange.SoftEdge.Radius = 100
    Else
        Selection.ShapeRange.SoftEdge.Radius = 0
    End If
    Range("A1").Select
End Sub
