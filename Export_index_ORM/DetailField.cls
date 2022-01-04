VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DetailField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "ORM.DetailFields"
Option Compare Database
Option Explicit

Implements IRecord

' Variables
Private Type TDetailField
    ID As Double
    Name As String
    Parent As DetailTable
    FieldType As String 'Field.Type
    LookupTable As LookupTable
    ' Caption,Default,Suffix,Width,Height,Format,TextAlign
End Type

Private this As TDetailField

' Unique Properties
Public Property Get DetailTable() As DetailTable
    Set DetailTable = this.Parent
End Property

Public Property Set DetailTable(ByVal RHS As DetailTable)
    Set this.Parent = RHS
End Property

Public Property Get LookupTable() As LookupTable
    Set LookupTable = this.LookupTable
End Property

Public Property Set LookupTable(ByVal RHS As LookupTable)
    Set this.LookupTable = RHS
End Property

Public Property Get FieldType() As String
    FieldType = this.FieldType
End Property

Public Property Let FieldType(ByVal RHS As String)
    this.FieldType = RHS
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
    Dim result As DetailField
    Set result = New DetailField
    With result
        .ID = rs.Fields("ID")
        .Name = rs.Fields("FieldName") ' TODO Const this
        Set .DetailTable = New DetailTable
        .DetailTable.Name = rs.Fields("TableName")
        .FieldType = rs.Fields("FieldType")
        Set .LookupTable = New LookupTable
        .LookupTable.Name = Nz(rs.Fields("LookupTable"), vbNullString)
    End With
    Set IRecord_Create = result
End Function


Public Function ToString() As String
    ToString = this.ID & "# " & this.Name
    ToString = ToString & ", DetailTable = " & this.Parent.Name
    If Not this.LookupTable Is Nothing Then
        ToString = ToString & ", LookupTable = " & this.LookupTable.Name
    End If
End Function
