VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EntityTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("index_ORM")
Option Compare Database
Option Explicit

Implements IRecordCollection

' Variables
Private Type TEntityTypes
    Items As Collection
    TableName As String
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

Public Property Get TableName() As String
    TableName = this.TableName
End Property

Public Property Let TableName(ByVal TableName As String)
    this.TableName = TableName
End Property

Public Property Get Self() As EntityTypes
    Set Self = Me
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
Private Property Get IRecordCollection_Items() As Collection
    Set IRecordCollection_Items = Items
End Property

Private Sub IRecordCollection_Load()
    Load
End Sub

Private Sub IRecordCollection_Clear()
    Clear
End Sub

' Private Methods
