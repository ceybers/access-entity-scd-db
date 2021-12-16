Attribute VB_Name = "modTest"
Option Compare Database
Option Explicit

Public Function Test()
    Dim thing As IThing
    Dim Detail As clsDetail
    
    Set Detail = DetailFactory.Create(1, "hi", "tblHi")
    Set thing = Detail
    
    Debug.Print "detail.TableName " & Detail.TableName
    Debug.Print "thing.ID " & thing.ID
    
    Debug.Print Detail.ID
    ' Debug.Print thing.TableName
End Function

