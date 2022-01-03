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
    Commits As Commits
    Tracks As Tracks
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

Public Property Get Commits() As Commits
    Set Commits = this.Commits
End Property

Public Property Get Tracks() As Tracks
    Set Tracks = this.Tracks
End Property

Public Property Get IsReady() As Boolean
    IsReady = this.IsReady
End Property

' Ctor
Public Sub Class_Initialize()
    LoadEntityTypes
    LoadEntities
    LoadCommits
    LoadTracks
    this.IsReady = True
End Sub

' Private methods
Private Sub LoadEntityTypes()
    With New EntityTypes
        Set this.EntityTypes = .Self
        .TableName = modConstants.ENTITYTYPES_TABLE
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

Private Sub LoadCommits()
    With New Commits
        Set this.Commits = .Self
        .TableName = modConstants.COMMITS_TABLE
        .Load
    End With
End Sub

Private Sub LoadTracks()
    With New Tracks
        Set this.Tracks = .Self
        .TableName = modConstants.TRACKS_TABLE
        .Load
        .MapRefs this.Commits
        this.Commits.MapRefs .Self
    End With
End Sub
