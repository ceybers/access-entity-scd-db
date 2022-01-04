VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EntityType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "ORM.EntityTypes"
Option Compare Database
Option Explicit

Implements IRecord

' Variables
Private Type TEntityType
    ID As Double
    Name As String
    Entities As Entities
End Type

Private this As TEntityType

' Unique Properties
Public Property Get Entities() As Entities
    Set Entities = this.Entities
End Property

Public Property Set Entities(ByVal RHS As Entities)
    Set this.Entities = RHS
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
    Set IRecord_Create = New EntityType
    With IRecord_Create
        .ID = rs.Fields("ID")
        .Name = rs.Fields(ENTITYTYPE_FIELDNAME)
    End With
End Function


Public Function ToString() As String
    ToString = this.ID & "# " & this.Name
    If Not this.Entities Is Nothing Then
        ToString = ToString & ", Entities = " & this.Entities.Count
    End If
End Function
