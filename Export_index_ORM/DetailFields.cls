VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DetailFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ORM.DetailFields"
Option Compare Database
Option Explicit

Implements IRecordCollection

' Variables
Private Type TDetailFields
    Items As Collection
    TableName As String
    Parent As DetailTable
End Type

Private this As TDetailFields

' Unique Properties
Public Property Get Parent() As DetailTable
    Set Parent = this.Parent
End Property

Public Property Set Parent(ByVal DetailTable As DetailTable)
    Set this.Parent = DetailTable
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

Public Property Get Self() As DetailFields
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
    LoadFromRecordset this.Items, DETAILFIELDS_TABLE, DetailField
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

Public Sub MapRefsDetailTables(ByRef externalCollection As IRecordCollection)
    Dim detFld As DetailField
    Dim detTbl As DetailTable
    Dim DetailTables As DetailTables
    Set DetailTables = externalCollection
    
    For Each detFld In this.Items
        Set detTbl = DetailTables.GetByName(detFld.DetailTable.Name)
        Set detFld.DetailTable = detTbl
    Next detFld
End Sub

Public Sub MapRefsLookupTables(ByRef externalCollection As IRecordCollection)
    Dim detFld As DetailField
    Dim lkpTbl As LookupTable
    Dim LookupTables As LookupTables
    Set LookupTables = externalCollection
    
    For Each detFld In this.Items
        If detFld.LookupTable.Name = vbNullString Then
            Set detFld.LookupTable = Nothing
        Else
            Set lkpTbl = LookupTables.GetByName(detFld.LookupTable.Name)
            Set detFld.LookupTable = lkpTbl
        End If
    Next detFld
End Sub

Public Function GetByDetailTable(ByRef DetailTable As DetailTable) As DetailFields
    Dim result As DetailFields
    Dim detFld As DetailField
    
    Set result = New DetailFields
    
    result.TableName = Me.TableName
    
    For Each detFld In this.Items
        If detFld.DetailTable = DetailTable Then
            result.Items.Add detFld, modConstants.COLLECTION_INDEX_PREFIX & detFld.ID
        End If
    Next detFld
    
    Set GetByDetailTable = result
End Function
' Private Methods
