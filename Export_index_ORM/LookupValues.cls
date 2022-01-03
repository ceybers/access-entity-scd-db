VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LookupValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ORM.LookupValues"
Option Compare Database
Option Explicit

Implements IRecordCollection

' Variables
Private Type TLookupValues
    Items As Collection
    TableName As String
    Parent As LookupTable
End Type

Private this As TLookupValues

' Unique Properties
Public Property Get Parent() As LookupTable
    Set Parent = this.Parent
End Property

Public Property Set Parent(ByVal LookupTable As LookupTable)
    Set this.Parent = LookupTable
End Property

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

Public Property Get Self() As LookupValues
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
    LoadFromRecordset this.Items, LOOKUPS_TABLE, LookupTable
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
    Set IRecordCollection_GetByID = this.Items.Item(COLLECTION_INDEX_PREFIX & ID)
End Function

' Public Methods
Public Function GetByID(ByVal ID As Double) As Commit
    Set GetByID = IRecordCollection_GetByID(ID)
End Function

Public Sub LoadValues()
    LinkTable this.TableName
    LoadFromRecordset this.Items, this.TableName, LookupValue
    SetParent
End Sub

' Private Methods
Private Sub SetParent()
    Dim lkpv As LookupValue
    For Each lkpv In this.Items
        Set lkpv.Parent = Me.Parent
    Next lkpv
End Sub
