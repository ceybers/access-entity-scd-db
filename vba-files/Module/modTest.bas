Attribute VB_Name = "modTest"
Option Compare Database
Option Explicit

Public Function Test()
    Dim Detail As IThing
    Set Detail = DetailFactory.Create(1, "name", "table")
    Debug.Print Detail.Name
End Function

