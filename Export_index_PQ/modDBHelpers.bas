Attribute VB_Name = "modDBHelpers"
'@Folder("index_PQ")
Option Compare Database
Option Explicit

Public Sub PrintOpenDatabases()
    Dim c As Integer
    Dim i As Integer
    c = DBEngine(0).Databases.Count
    Debug.Print "There are " & c & "x open database(s)"
    For i = 0 To (c - 1)
        Debug.Print "#" & i & " " & DBEngine(0).Databases(i).Name
    Next i
    Debug.Print
End Sub
