VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMigrateCommits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MigrateCommits"
Option Compare Database
Option Explicit

Private Type TClsMigrateCommit
    SourceTableName As String
    SourceIDField As String
    DestinationTableName As String
    DestinationIDField As String
    Fields As Collection
End Type

Private this As TClsMigrateCommit

Public Property Get SourceTableName() As String
    SourceTableName = this.SourceTableName
End Property

Public Property Let SourceTableName(ByVal RHS As String)
    this.SourceTableName = RHS
End Property

Public Property Get SourceIDField() As String
    SourceIDField = this.SourceIDField
End Property

Public Property Let SourceIDField(ByVal RHS As String)
    this.SourceIDField = RHS
End Property

Public Property Get DestinationTableName() As String
    DestinationTableName = this.DestinationTableName
End Property

Public Property Let DestinationTableName(ByVal RHS As String)
    this.DestinationTableName = RHS
End Property

Public Property Get DestinationIDField() As String
    DestinationIDField = this.DestinationIDField
End Property

Public Property Let DestinationIDField(ByVal RHS As String)
    this.DestinationIDField = RHS
End Property

Public Property Get Fields() As Collection
    Set Fields = this.Fields
End Property

Public Property Set Fields(ByVal RHS As Collection)
    Set this.Fields = RHS
End Property
