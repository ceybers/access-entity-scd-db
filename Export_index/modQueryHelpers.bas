Attribute VB_Name = "modQueryHelpers"
'@Folder "Helpers"
Option Compare Database
Option Explicit

Public Function CreateQuery(ByVal queryName As String, ByVal sql As String, Optional ByRef db As DAO.Database) As Boolean
If db Is Nothing Then
        Set db = CurrentDb
    End If

    If DoesQueryExist(queryName, db) Then Exit Function
    
    db.CreateQueryDef queryName, sql
    
    db.QueryDefs.Refresh
    
    CreateQuery = True
End Function

Public Function DoesQueryExist(ByVal queryName As String, Optional ByRef db As DAO.Database) As Boolean
    Dim qdf As QueryDef
    
    If db Is Nothing Then Set db = CurrentDb

    For Each qdf In db.QueryDefs
        If qdf.name = queryName Then
            DoesQueryExist = True
            Exit Function
        End If
    Next qdf
End Function
