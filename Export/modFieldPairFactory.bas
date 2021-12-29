Attribute VB_Name = "modFieldPairFactory"
Option Compare Database
Option Explicit

Public Function CreateFieldPair(src As String, dst As String) As clsFieldPair
    Set CreateFieldPair = New clsFieldPair
    With CreateFieldPair
        .Source = src
        .Destination = dst
    End With
End Function

