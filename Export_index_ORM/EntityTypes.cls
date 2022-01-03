VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EntityTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ORM.EntityTypes"
Option Compare Database
Option Explicit

Implements IRecordCollection

' Variables
Private Type TEntityTypes
    Items As Collection
    tableName As String
End Type

Private this As TEntityTypes

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

Public Property Get Self() As EntityTypes
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
    LoadFromRecordset this.Items, ENTITYTYPES_TABLE, EntityType
End Sub

Public Sub Clear()
    ClearCollection this.Items
End Sub

' Interface Methods
Private Sub IRecordCollection_Load()
    Load
End Sub

Private Property Get IRecordCollection_Count() As Long
    IRecordCollection_Count = Count
End Property

Private Sub IRecordCollection_Clear()
    Clear
End Sub

' Public Methods
Public Function GetByID(ByVal ID As Double) As EntityType
    Set GetByID = IRecordCollection_GetByID(ID)
End Function

Public Function GetByName(ByVal Name As String) As EntityType
    Dim et As EntityType
    For Each et In this.Items
        If et.Name = Name Then
            Set GetByName = et
            Exit Function
        End If
    Next et
End Function

Public Sub MapRefs(ByRef externalCollection As IRecordCollection)
    Dim et As EntityType
    Dim Entities As Entities
    Set Entities = externalCollection
    For Each et In this.Items
        Set et.Entities = Entities.GetByEntityType(et)
    Next et
End Sub

' Private Methods
Private Function IRecordCollection_GetByID(ByVal ID As Double) As IRecord
    Dim ent As IRecord
    For Each ent In this.Items
        If ent.ID = ID Then
            Set IRecordCollection_GetByID = ent
            Exit Function
        End If
    Next ent
End Function


