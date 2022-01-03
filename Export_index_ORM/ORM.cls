VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("index_ORM")
Option Compare Database
Option Explicit

Private Type TORM
    EntityTypes As EntityTypes
    IsReady As Boolean
End Type

Private this As TORM

Public Property Get EntityTypes() As EntityTypes
    Set EntityTypes = this.EntityTypes
End Property

Public Property Get IsReady() As Boolean
    IsReady = this.IsReady
End Property

Public Sub Class_Initialize()
    LoadEntityTypes
    LoadEntities
    this.IsReady = True
End Sub

Private Sub LoadEntityTypes()
    With New EntityTypes
        Set this.EntityTypes = .Self
        .TableName = modConstants.ENTITYTYPES_TABLE
        .Load
    End With
End Sub

Private Sub LoadEntities()

End Sub
