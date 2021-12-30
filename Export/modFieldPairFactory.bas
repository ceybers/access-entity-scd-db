Attribute VB_Name = "modFieldPairFactory"
'@Folder "Common"
Option Compare Database
Option Explicit

Public Function CreateFieldPair(src As String, dst As String, Optional lkp As String = "") As clsFieldPair
    Set CreateFieldPair = New clsFieldPair
    With CreateFieldPair
        .Source = src
        .Destination = dst
        .Lookup = lkp
    End With
End Function

Public Function CreateFieldPairFromArray(arr As Variant) As clsFieldPair
    Set CreateFieldPairFromArray = New clsFieldPair
    With CreateFieldPairFromArray
        .Source = arr(0)
        .Destination = arr(1)
        .Lookup = arr(2)
    End With
End Function

