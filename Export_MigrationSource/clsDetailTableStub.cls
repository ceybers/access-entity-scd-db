VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDetailTableStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MigrateDetails")
Option Compare Database
Option Explicit

Private Type TClsDetailTableStub
    tableName As String
    ID As String
    TrackFK As String
End Type

Private this As TClsDetailTableStub

Public Property Get tableName() As String
    tableName = this.tableName
End Property

Public Property Let tableName(ByVal RHS As String)
    this.tableName = RHS
End Property

Public Property Get ID() As String
    ID = this.ID
End Property

Public Property Let ID(ByVal RHS As String)
    this.ID = RHS
End Property

Public Property Get TrackFK() As String
    TrackFK = this.TrackFK
End Property

Public Property Let TrackFK(ByVal RHS As String)
    this.TrackFK = RHS
End Property
