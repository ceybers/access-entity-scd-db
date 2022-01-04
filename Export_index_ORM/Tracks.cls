VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tracks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ORM.Tracks"
Option Compare Database
Option Explicit

Implements IRecordCollection

' Variables
Private Type TTracks
    Items As Collection
    TableName As String
End Type

Private this As TTracks

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

Public Property Get Self() As Tracks
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
    LoadFromRecordset this.Items, TRACKS_TABLE, Track
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
    Dim trk As Track
    For Each trk In this.Items
        If trk.ID = ID Then
            Set IRecordCollection_GetByID = trk
            Exit Function
        End If
    Next trk
End Function

' Public Methods
Public Function Filter(ByRef Commit As Commit) As IRecordCollection
    Set Filter = New Tracks
    With Filter
        .TableName = Me.TableName
    End With
    
    Dim trk As Track
    For Each trk In this.Items
        If trk.Commit Is Commit Then
            Filter.Items.Add trk
        End If
    Next trk
End Function

Public Function GetByID(ByVal ID As Double) As Track
    Set GetByID = IRecordCollection_GetByID(ID)
End Function

Public Sub MapRefsCommits(ByRef externalCollection As IRecordCollection)
    Dim trk As Track

    For Each trk In this.Items
        Set trk.Commit = externalCollection.GetByID(trk.Commit.ID)
    Next trk
End Sub

' Private Methods