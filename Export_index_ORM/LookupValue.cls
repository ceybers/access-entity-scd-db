VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LookupValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "ORM.LookupValues"
Option Compare Database
Option Explicit

Implements IRecord

' Variables
Private Type TLookupValue
    ID As Double
    Name As String
    Parent As LookupTable
End Type

Private this As TLookupValue

' Unique Properties
Public Property Get Parent() As LookupTable
    Set Parent = this.Parent
End Property

Public Property Set Parent(ByVal RHS As LookupTable)
    Set this.Parent = RHS
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
    Dim result As LookupValue
    Set result = New LookupValue
    With result
        .ID = rs.Fields("ID")
        .Name = rs.Fields(1) ' Always the second field
    End With
    Set IRecord_Create = result
End Function

Public Function ToString() As String
    ToString = this.ID & "# " & this.Name
    'ToString = this.ID & "# " & this.Name & ", Parent = " & this.Parent.ToString
End Function
