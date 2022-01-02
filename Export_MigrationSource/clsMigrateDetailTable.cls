VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMigrateDetailTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MigrateDetails")
Option Compare Database
Option Explicit

Private Type TClass1
    Source As clsDetailTableStub
    Destination As clsDetailTableStub
    Fields As Collection
End Type

Private this As TClass1

Public Property Get Source() As clsDetailTableStub
    Set Source = this.Source
End Property

Public Property Set Source(ByVal RHS As clsDetailTableStub)
    Set this.Source = RHS
End Property

Public Property Get Destination() As clsDetailTableStub
    Set Destination = this.Destination
End Property

Public Property Set Destination(ByVal RHS As clsDetailTableStub)
    Set this.Destination = RHS
End Property

Public Property Get Fields() As Collection
    Set Fields = this.Fields
End Property

Public Property Set Fields(ByVal RHS As Collection)
    Set this.Fields = RHS
End Property
