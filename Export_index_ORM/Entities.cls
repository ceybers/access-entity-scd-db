VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Entities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ORM.Entities"
Option Compare Database
Option Explicit

Implements IRecordCollection

' Variables
Private Type TEntities
    Items As Collection
    tableName As String
End Type

Private this As TEntities

' Properties
'@DefaultMember
Public Property Get Items() As Collection
Attribute Items.VB_UserMemId = 0
    Set Items = this.Items
End Property

'@Enumerator
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = this.Items.[_NewEnum]
End Property

Public Property Get Count() As Long
    Count = this.Items.Count
End Property

Public Property Get tableName() As String
    tableName = this.tableName
End Property

Public Property Let tableName(ByVal tableName As String)
    this.tableName = tableName
End Property

Public Property Get Self() As Entities
    Set Self = Me
End Property

' Ctor
Private Sub Class_Initialize()
    Set this.Items = New Collection
End Sub

' Methods
Public Sub Load()
    LoadFromRecordset this.Items, ENTITIES_TABLE, Entity
    LinkParents
End Sub

Public Sub Clear()
    ClearCollection this.Items
End Sub

' Interface Methods
Private Property Get IRecordCollection_Items() As Collection
    Set IRecordCollection_Items = Items
End Property

Public Property Get IRecordCollection_TableName() As String
    IRecordCollection_TableName = this.tableName
End Property

Public Property Let IRecordCollection_TableName(ByVal tableName As String)
    this.tableName = tableName
End Property

Private Sub IRecordCollection_Load()
    Load
End Sub

Private Property Get IRecordCollection_Count() As Long
    IRecordCollection_Count = Count
End Property

Private Sub IRecordCollection_Clear()
    Clear
End Sub

Public Property Get IRecordCollection_Self() As IRecordCollection
    Set IRecordCollection_Self = Me
End Property

' Private Methods
Public Function IRecordCollection_GetByID(ID As Double) As IRecord
    Dim ent As IRecord
    For Each ent In this.Items
        If ent.ID = ID Then
            Set IRecordCollection_GetByID = ent
            Exit Function
        End If
    Next ent
End Function

Private Sub LinkParents()
    Dim ent As Entity
    Dim parentFK As Double
    
    For Each ent In this.Items
        parentFK = GetFieldValue(ENTITIES_TABLE, ent.ID, "ParentFK")
        Set ent.parent = IRecordCollection_GetByID(parentFK)
    Next ent
End Sub

Public Sub MapRefs(externalCollection As IRecordCollection)
    Dim ent As Entity
    Dim parentFK As Double
    
    For Each ent In this.Items
        parentFK = GetFieldValue(ENTITIES_TABLE, ent.ID, "EntityType")
        Set ent.EntityType = externalCollection.GetByID(parentFK)
    Next ent
End Sub
