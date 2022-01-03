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
    parent As Entity
    EntityType As EntityType
End Type

Private this As TEntity

' Properties
'@DefaultMember
Public Property Get Name() As String
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

Public Property Get parent() As Entity
     parent = this.parent
End Property

Public Property Set parent(ByVal parent As Entity)
    Set this.parent = parent
End Property

Public Property Get EntityType() As EntityType
     EntityType = this.EntityType
End Property

Public Property Set EntityType(ByVal EntityType As EntityType)
    Set this.EntityType = EntityType
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
Public Function IRecord_Create(ByRef rs As Recordset) As IRecord
    Set IRecord_Create = New Entity
    With IRecord_Create
        .ID = rs.Fields("ID")
        .Name = rs.Fields("Entity") ' TODO Const this
    End With
End Function

Public Function ToString() As String
    ToString = this.ID & "# " & this.Name
    If Not this.parent Is Nothing Then
        ToString = ToString & ", Parent = " & this.parent.Name
    End If
    If Not this.EntityType Is Nothing Then
        ToString = ToString & ", EntityType = " & this.EntityType.Name
    End If
End Function
