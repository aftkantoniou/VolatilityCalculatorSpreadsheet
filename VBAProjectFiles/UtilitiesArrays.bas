Attribute VB_Name = "UtilitiesArrays"
Option Explicit

Public Function arrayContains(valueToBeFound As Variant, arr As Variant) As Boolean

' Function to check if a value is in an array of values.
' True if is in array, false otherwise
' Parameter1: Value to search for
' Parameter2: Array of values of any data type.

    arrayContains = False
    
    Dim element As Variant
    
    ' Case Array Empty:
    ' =================
    
    If isArrayAllocated(arr) = False Then
        arrayContains = False
        Exit Function
    End If
    
    ' Case Array not Empty:
    ' =====================
    
    For Each element In arr
        If element = valueToBeFound Then
            arrayContains = True
            Exit Function
        End If
    Next element
        
End Function

Public Function isArrayAllocated(arr As Variant) As Boolean
    
    On Error Resume Next
    isArrayAllocated = IsArray(arr) And Not IsError(LBound(arr, 1)) And LBound(arr, 1) <= UBound(arr, 1)

End Function
