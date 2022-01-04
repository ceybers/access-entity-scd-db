VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DetailTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "ORM.DetailTables"
Option Compare Database
Option Explicit

Implements IRecord

' Variables
Private Type TDetailTable
    ID As Double
    Name As String
    DetailFields As DetailFields
    DetailValues As DetailValues
End Type

Private this As TDetailTable

' Unique Properties
Public Property Get DetailFields() As DetailFields
    Set DetailFields = this.DetailFields
End Property

Public Property Set DetailFields(DetailFields As DetailFields)
    Set this.DetailFields = DetailFields
End Property

Public Property Get DetailValues() As DetailValues
    Set DetailValues = this.DetailValues
End Property

Public Property Set DetailValues(DetailValues As DetailValues)
    Set this.DetailValues = DetailValues
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
    Dim result As DetailTable
    Set result = New DetailTable

    Dim DetailValues As DetailValues
    Set DetailValues = New DetailValues
    DetailValues.TableName = rs.Fields(DETAILTABLE_FIELDNAME)
    Set DetailValues.Parent = result
    DetailValues.LoadValues

    With result
        .ID = rs.Fields("ID")
        .Name = rs.Fields(DETAILTABLE_FIELDNAME)
        Set .DetailValues = DetailValues
    End With
    Set IRecord_Create = result
End Function

Public Function ToString() As String
    ToString = this.ID & "# " & this.Name
    ToString = ToString & ", Fields = " & this.DetailFields.Count
    If Not this.DetailValues Is Nothing Then
        ToString = ToString & ", Values = " & this.DetailValues.Count
    End If
End Function

Public Sub MapRefsEntities(ByRef externalCollection As IRecordCollection)
    this.DetailValues.MapRefsEntities externalCollection
End Sub

Public Sub MapRefsTracks(ByRef externalCollection As IRecordCollection)
    this.DetailValues.MapRefsTracks externalCollection
End Sub
