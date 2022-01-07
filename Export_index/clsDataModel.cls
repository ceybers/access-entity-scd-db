VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule

'@Folder "DataModels"
Option Compare Database
Option Explicit

' Private Variables
Private Type TDataModel
     Details As collection
     Entities As collection
     EntityTypes As collection
     Commits As collection
End Type
Private this As TDataModel

' Ctor

' Getters
Public Property Get Details() As collection
    If this.Details Is Nothing Then
        Set this.Details = LoadThings(DETAILS_TABLE)
    End If
    Set Details = this.Details
End Property

Public Property Get Entities() As collection
    If this.Entities Is Nothing Then
        Set this.Entities = LoadThings(ENTITIES_TABLE)
    End If
    Set Entities = this.Entities
End Property

Public Property Get EntityTypes() As collection
    If this.EntityTypes Is Nothing Then
        Set this.EntityTypes = LoadThings(ENTITYTYPES_TABLE)
    End If
    Set EntityTypes = this.EntityTypes
End Property

Public Property Get Commits() As collection
    If this.Commits Is Nothing Then
        Set this.Commits = LoadThings(COMMITS_TABLE)
    End If
    Set Commits = this.Commits
End Property

' Public Methods
Public Function GetThingByID(thingCollection As collection, idx As Double) As IThing
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
Private Function LoadThings(tableName As String) As collection
    Dim rs As DAO.Recordset
    Set LoadThings = New collection

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
