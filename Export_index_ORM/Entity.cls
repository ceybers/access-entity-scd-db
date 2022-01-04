VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Entity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "ORM.Entities"
Option Compare Database
Option Explicit

Implements IRecord

' Variables
Private Type TEntity
    ID As Double
    Name As String
    Parent As Entity
    EntityType As EntityType
    Details As DetailValues
End Type

Private this As TEntity

' Unique Properties
Public Property Get Details() As DetailValues
     Set Details = this.Details
End Property

Public Property Set Details(ByRef DetailValues As DetailValues)
    Set this.Details = DetailValues
End Property

Public Property Get Parent() As Entity
     Set Parent = this.Parent
End Property

Public Property Set Parent(ByVal Parent As Entity)
    Set this.Parent = Parent
End Property

Public Property Get EntityType() As EntityType
     Set EntityType = this.EntityType
End Property

Public Property Set EntityType(ByVal EntityType As EntityType)
    Set this.EntityType = EntityType
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
    Dim result As Entity
    Set result = New Entity
    Dim et As EntityType
    With result
        .ID = rs.Fields("ID")
        .Name = rs.Fields(ENTITY_FIELDNAME)
        Set .EntityType = CreateEntityType(rs.Fields(modConstants.ENTITYTYPE_FIELDNAME))
    End With
    Set IRecord_Create = result
End Function


Public Function ToString() As String
    ToString = this.ID & "# " & this.Name
    If Not this.Parent Is Nothing Then
        ToString = ToString & ", Parent = " & this.Parent.Name
    End If
    If Not this.EntityType Is Nothing Then
        ToString = ToString & ", EntityType = " & this.EntityType.Name
    End If
End Function
