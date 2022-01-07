VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SubformControlSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "Provisioning"
Option Compare Database
Option Explicit

Private Type TSubformControlSet
    FieldName As String
    Caption As String
    Width As String
    LookupTable As String
    Suffix As String
    Format As String
    Textalign As String
    RecordsetQuery As String
End Type

Private this As TSubformControlSet

Public Property Get FieldName() As String
    FieldName = this.FieldName
End Property

Public Property Let FieldName(ByVal RHS As String)
    this.FieldName = RHS
End Property

Public Property Get Caption() As String
    Caption = this.Caption
End Property

Public Property Let Caption(ByVal RHS As String)
    this.Caption = RHS
End Property

Public Property Get Width() As String
    Width = this.Width
End Property

Public Property Let Width(ByVal RHS As String)
    this.Width = RHS
End Property

Public Property Get LookupTable() As String
    LookupTable = this.LookupTable
End Property

Public Property Let LookupTable(ByVal RHS As String)
    this.LookupTable = RHS
End Property

Public Property Get Suffix() As String
    Suffix = this.Suffix
End Property

Public Property Let Suffix(ByVal RHS As String)
    this.Suffix = RHS
End Property

Public Property Get Format() As String
    Format = this.Format
End Property

Public Property Let Format(ByVal RHS As String)
    this.Format = RHS
End Property

Public Property Get Textalign() As String
    Textalign = this.Textalign
End Property

Public Property Let Textalign(ByVal RHS As String)
    this.Textalign = RHS
End Property

Public Property Get RecordsetQuery() As String
    RecordsetQuery = this.RecordsetQuery
End Property

Public Property Let RecordsetQuery(ByVal RHS As String)
    this.RecordsetQuery = RHS
End Property

Public Function Create(ByVal FieldName As String, ByVal Caption As String, ByVal Width As Long) As subformcontrolset
    Set Create = New subformcontrolset
    With Create
        .FieldName = FieldName
        .Caption = Caption
        .Width = Width
    End With
End Function

Public Function CreateFromRecordset(ByRef rs As Recordset) As subformcontrolset
    Set CreateFromRecordset = New subformcontrolset
    With CreateFromRecordset
        .FieldName = rs.fields("FieldName").Value
        .Caption = Nz(rs.fields("Caption").Value, .FieldName)
        .Width = Nz(rs.fields("Width").Value, vbNullString)
        .LookupTable = Nz(rs.fields("LookupTable").Value, vbNullString)
        .Suffix = Nz(rs.fields("Suffix").Value)
        .Format = Nz(rs.fields("Format").Value)
        .Textalign = Nz(rs.fields("TextAlign").Value)
        .RecordsetQuery = rs.name
    End With
End Function
