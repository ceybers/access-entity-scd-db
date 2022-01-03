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
'@Folder("index_ORM")
Option Compare Database
Option Explicit

Implements IRecord

' Variables
Private Type TEntityType
    ID As Double
    Name As String
End Type

Private this As TEntityType

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
    Set IRecord_Create = New EntityType
    With IRecord_Create
        .ID = rs.Fields("ID")
        .Name = rs.Fields(ENTITYTYPE_FIELDNAME)
    End With
End Function
