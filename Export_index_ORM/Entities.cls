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
    TableName As String
End Type

Private this As TEntities

' Properties
'@DefaultMember
Public Property Get Items() As Collection
Attribute Items.VB_UserMemId = 0
    Set Items = this.Items
End Property

'@Enumerator
'@MemberAttribute VB_MemberFlags, "40"
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = this.Items.[_NewEnum]
End Property

Public Property Get Count() As Long
    Count = this.Items.Count
End Property

Public Property Get TableName() As String
    TableName = this.TableName
End Property

Public Property Let TableName(ByVal TableName As String)
    this.TableName = TableName
End Property

Public Property Get Self() As Entities
    Set Self = Me
End Property

' Interface Properties
Private Property Get IRecordCollection_Items() As Collection
    Set IRecordCollection_Items = Items
End Property

Private Property Get IRecordCollection_TableName() As String
    IRecordCollection_TableName = this.TableName
End Property

Private Property Let IRecordCollection_TableName(ByVal TableName As String)
    this.TableName = TableName
End Property

Private Property Get IRecordCollection_Count() As Long
    IRecordCollection_Count = Count
End Property

Private Property Get IRecordCollection_Self() As IRecordCollection
    Set IRecordCollection_Self = Me
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
Private Sub IRecordCollection_Load()
    Load
End Sub

Private Sub IRecordCollection_Clear()
    Clear
End Sub

Private Function IRecordCollection_GetByID(ByVal ID As Double) As IRecord
    Set IRecordCollection_GetByID = this.Items(COLLECTION_INDEX_PREFIX & ID)
    Exit Function
    ' TODO Remove
    Dim ent As Entity
    For Each ent In this.Items
        If ent.ID = ID Then
            Set IRecordCollection_GetByID = ent
            Exit Function
        End If
    Next ent
End Function

' Public Methods
Public Function GetByEntityType(ByRef EntityType As EntityType) As IRecordCollection
    Set GetByEntityType = New Entities
    With GetByEntityType
        .TableName = Me.TableName
    End With
    
    Dim ent As Entity
    For Each ent In this.Items
        If ent.EntityType Is EntityType Then
            GetByEntityType.Items.Add ent, COLLECTION_INDEX_PREFIX & ent.ID
        End If
    Next ent

End Function

Public Function GetByID(ByVal ID As Double) As Entity
    Set GetByID = IRecordCollection_GetByID(ID)
End Function

Public Sub MapRefs(ByRef externalCollection As IRecordCollection)
    Dim ent As Entity
    Dim parentFK As Double
    
    For Each ent In this.Items
        parentFK = GetFieldValue(ENTITIES_TABLE, ent.ID, "EntityType")
        Set ent.EntityType = externalCollection.GetByID(parentFK)
    Next ent
End Sub

' Private Methods
Private Sub LinkParents()
    Dim ent As Entity
    Dim parentFK As Double
    
    For Each ent In this.Items
        parentFK = GetFieldValue(ENTITIES_TABLE, ent.ID, "ParentFK")
        If parentFK > 0 Then
            Set ent.Parent = IRecordCollection_GetByID(parentFK)
        End If
    Next ent
End Sub

