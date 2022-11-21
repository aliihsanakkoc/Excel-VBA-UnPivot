Attribute VB_Name = "UnPivotUDF"
Function UnPivot(ByVal cll As Range, ByVal sbtstnsys As Integer) As Variant()

Dim arr As Variant
Dim arrmtrs As Variant
Dim x, y, i, j, q, z As Integer
arr = cll.Value

ReDim arrmtrs(1 To (UBound(arr, 1) - 1) * (UBound(arr, 2) - sbtstnsys), 1 To sbtstnsys + 2)

For z = 1 To sbtstnsys
    For x = 2 To UBound(arr, 1)
        For y = ((UBound(arr, 2) - sbtstnsys) * (x - 2)) + 1 To _
        (UBound(arr, 2) - sbtstnsys) * (x - 1)
            arrmtrs(y, z) = arr(x, z)
        Next y
    Next x
Next z

j = 1
For q = 2 To UBound(arr, 1)
    For i = sbtstnsys + 1 To UBound(arr, 2)
        arrmtrs(j, sbtstnsys + 1) = arr(1, i)
        arrmtrs(j, sbtstnsys + 2) = arr(q, i)
    j = j + 1
    Next i
Next q

UnPivot = arrmtrs

End Function

