VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Hello"
'@ModuleDescription "Hello"
'@Folder("ORM")
Option Compare Database
Option Explicit

' Variables
Private Type TORM
    EntityTypes As EntityTypes
    Entities As Entities
    
    IsReady As Boolean
End Type

Private this As TORM

' Properties
Public Property Get EntityTypes() As EntityTypes
    Set EntityTypes = this.EntityTypes
End Property

Public Property Get Entities() As Entities
    Set Entities = this.Entities
End Property

Public Property Get IsReady() As Boolean
    IsReady = this.IsReady
End Property

' Ctor
Public Sub Class_Initialize()
    LoadEntityTypes
    LoadEntities
    this.IsReady = True
End Sub

' Private methods
Private Sub LoadEntityTypes()
    With New EntityTypes
        Set this.EntityTypes = .Self
        .tableName = modConstants.ENTITYTYPES_TABLE
        .Load
    End With
End Sub

Private Sub LoadEntities()
    With New Entities
        Set this.Entities = .Self
        .TableName = modConstants.ENTITIES_TABLE
        .Load
        .MapRefs this.EntityTypes
        this.EntityTypes.MapRefs .Self
    End With
    
End Sub