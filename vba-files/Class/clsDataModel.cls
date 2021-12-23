VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Constants
Const DETAILS_TABLE As String = "tblDetailTables"
Const ENTITIES_TABLE As String = "tblEntities"
Const ENTITYTYPES_TABLE As String = "lkpEntityTypes"
Const COMMITS_TABLE As String = "tblCommits"

' Private Variables
Private Type TModel
     Details As Collection
     entities As Collection
     EntityTypes As Collection
     Commits As Collection
End Type
Private This As TModel

' Ctor
Private Sub Class_Initialize()
    'Debug.Print "Model initialised"
    ' Changed to lazy loading
    'Set This.entities = LoadThings(ENTITIES_TABLE)
    'Set This.Details = LoadThings(DETAILS_TABLE)
    'Set This.EntityTypes = LoadThings(ENTITYTYPES_TABLE)
    'Set This.Commits = LoadThings(COMMITS_TABLE)
End Sub

' Getters
Public Property Get Details() As Collection
    If This.Details Is Nothing Then
        Set This.Details = LoadThings(DETAILS_TABLE)
    End If
    Set Details = This.Details
End Property

Public Property Get entities() As Collection
    If This.entities Is Nothing Then
        Set This.entities = LoadThings(ENTITIES_TABLE)
    End If
    Set entities = This.entities
End Property

Public Property Get EntityTypes() As Collection
    If This.EntityTypes Is Nothing Then
        Set This.EntityTypes = LoadThings(ENTITYTYPES_TABLE)
    End If
    Set EntityTypes = This.EntityTypes
End Property

Public Property Get Commits() As Collection
    If This.Commits Is Nothing Then
        Set This.Commits = LoadThings(COMMITS_TABLE)
    End If
    Set Commits = This.Commits
End Property

' Public Methods
Public Function GetThingByID(thingCollection As Collection, idx As Double) As IThing
    Dim thisThing As IThing
    For Each thisThing In thingCollection
        If thisThing.ID = idx Then
            Set GetThingByID = thisThing
            Exit Function
        End If
    Next thisThing
    Err.Raise 5, , "GetThingByID failed!"
End Function

' Private Methods
Private Function LoadThings(tableName As String) As Collection
    Dim rs As DAO.Recordset
    Set LoadThings = New Collection

    Set rs = CurrentDb.OpenRecordset("SELECT * FROM " & tableName)
    
    If Not rs.BOF And Not rs.EOF Then
        Do Until rs.EOF = True
            Select Case tableName
                Case DETAILS_TABLE
                    LoadThings.Add DetailFactory.CreateFromRecordset(rs)
                Case ENTITIES_TABLE
                    LoadThings.Add EntityFactory.CreateFromRecordset(rs)
                Case ENTITYTYPES_TABLE
                    LoadThings.Add EntityTypeFactory.CreateFromRecordset(rs)
                Case COMMITS_TABLE
                    LoadThings.Add CommitFactory.CreateFromRecordset(rs)
            End Select
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
End Function
