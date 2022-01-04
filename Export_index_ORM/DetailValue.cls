VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DetailValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "ORM.DetailValues"
Option Compare Database
Option Explicit

Implements IRecord

' Variables
Private Type TDetailValue
    ID As Double
    Name As String
    Parent As DetailTable
    Entity As Entity
    Track As Track
End Type

Private this As TDetailValue

' Unique Properties
Public Property Get IsLatest() As Boolean
    Debug.Assert Not this.Track Is Nothing
    IsLatest = (this.Track.ValidUntil = #12/31/9999#)
End Property

Public Property Get Parent() As DetailTable
    Set Parent = this.Parent
End Property

Public Property Set Parent(ByVal RHS As DetailTable)
    Set this.Parent = RHS
End Property

Public Property Get Entity() As Entity
    Set Entity = this.Entity
End Property

Public Property Set Entity(ByVal RHS As Entity)
    Set this.Entity = RHS
End Property

Public Property Get Track() As Track
    Set Track = this.Track
End Property

Public Property Set Track(ByVal RHS As Track)
    Set this.Track = RHS
End Property

' Common Properties

'@DefaultMember
Public Property Get Name() As String
Attribute Name.VB_UserMemId = 0
    Name = this.Name
End Property

Public Property Let Name(ByVal RHS As String)
    this.Name = RHS
End Property

Public Property Get ID() As Double
    ID = this.ID
End Property

Public Property Let ID(ByVal RHS As Double)
    this.ID = RHS
End Property

' Interface
Private Property Get IRecord_ID() As Double
    IRecord_ID = ID
End Property

Private Property Let IRecord_ID(ByVal RHS As Double)
    ID = RHS
End Property

Private Property Get IRecord_Name() As String
    IRecord_Name = Name
End Property

Private Property Let IRecord_Name(ByVal RHS As String)
    Name = RHS
End Property

' Methods
Private Function IRecord_Create(ByRef rs As Recordset) As IRecord
    Dim result As DetailValue
    Set result = New DetailValue
    With result
        .ID = rs.Fields("ID")
        .Name = vbNullString ' These do not have name fields
        Set .Entity = modPlaceholderRecords.CreateEntity(Nz(rs.Fields("EntityFK"), 0))
        Set .Track = modPlaceholderRecords.CreateTrack(Nz(rs.Fields("TrackFK"), 0))
    End With
    Set IRecord_Create = result
End Function


Public Function ToString() As String
    ToString = this.ID & "#"
    If Not this.Parent Is Nothing Then
        ToString = ToString & ", DetailTable = '" & this.Parent.Name & "'"
    End If
    If Not this.Entity Is Nothing Then
        ToString = ToString & ", Entity = '" & this.Entity.Name & "'"
    End If
    If Not this.Track Is Nothing Then
        ToString = ToString & ", Track = #" & this.Track.ID
    End If
End Function
