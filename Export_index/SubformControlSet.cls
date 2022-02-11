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

Private This As TSubformControlSet

Public Property Get FieldName() As String
    FieldName = This.FieldName
End Property

Public Property Let FieldName(ByVal rhs As String)
    This.FieldName = rhs
End Property

Public Property Get Caption() As String
    Caption = This.Caption
End Property

Public Property Let Caption(ByVal rhs As String)
    This.Caption = rhs
End Property

Public Property Get Width() As String
    Width = This.Width
End Property

Public Property Let Width(ByVal rhs As String)
    This.Width = rhs
End Property

Public Property Get LookupTable() As String
    LookupTable = This.LookupTable
End Property

Public Property Let LookupTable(ByVal rhs As String)
    This.LookupTable = rhs
End Property

Public Property Get Suffix() As String
    Suffix = This.Suffix
End Property

Public Property Let Suffix(ByVal rhs As String)
    This.Suffix = rhs
End Property

Public Property Get Format() As String
    Format = This.Format
End Property

Public Property Let Format(ByVal rhs As String)
    This.Format = rhs
End Property

Public Property Get Textalign() As String
    Textalign = This.Textalign
End Property

Public Property Let Textalign(ByVal rhs As String)
    This.Textalign = rhs
End Property

Public Property Get RecordsetQuery() As String
    RecordsetQuery = This.RecordsetQuery
End Property

Public Property Let RecordsetQuery(ByVal rhs As String)
    This.RecordsetQuery = rhs
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
        .RecordsetQuery = rs.Name
    End With
End Function
