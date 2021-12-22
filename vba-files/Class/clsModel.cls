VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsModel"
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

' Private Variables
Private Type TModel
     Details As Collection
     entities As Collection
     EntityTypes As Collection
End Type
Private This As TModel

' Ctor
Private Sub Class_Initialize()
    'Debug.Print "Model initialised"
    Set This.entities = LoadThings(ENTITIES_TABLE)
    Set This.Details = LoadThings(DETAILS_TABLE)
    Set This.EntityTypes = LoadThings(ENTITYTYPES_TABLE)
End Sub

' Getters
Public Property Get Details() As Collection
    Set Details = This.Details
End Property

Public Property Get entities() As Collection
    Set entities = This.entities
End Property

Public Property Get EntityTypes() As Collection
    Set EntityTypes = This.EntityTypes
End Property

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
            End Select
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
End Function
