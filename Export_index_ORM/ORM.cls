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
    Tracks As Tracks ' Do we need these here? They are under Commits anyway
    LookupTables As LookupTables
    DetailTables As DetailTables
    DetailFields As DetailFields
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

Public Property Get LookupTables() As LookupTables
    Set LookupTables = this.LookupTables
End Property

Public Property Get DetailTables() As DetailTables
    Set DetailTables = this.DetailTables
End Property

Public Property Get DetailFields() As DetailFields
    Set DetailFields = this.DetailFields
End Property

Public Property Get IsReady() As Boolean
    IsReady = this.IsReady
End Property

' Ctor
Public Sub Class_Initialize()
    'InitializeWithTimers
    InitializeWithoutTimers
    Debug.Print "ORM loaded."
    this.IsReady = True
End Sub

Private Sub InitializeWithTimers()
    Dim start As Long
    start = VBA.DateTime.Timer
    Debug.Print "Loading ORM..."
    Debug.Print Format$(VBA.DateTime.Timer - start, " 0.0000000 ") '& "LoadEntityTypes"
    LoadEntityTypes
    Debug.Print Format$(VBA.DateTime.Timer - start, " 0.0000000 ") '& "LoadEntityTypes"
    LoadEntities
    Debug.Print Format$(VBA.DateTime.Timer - start, " 0.0000000 ") '& "LoadEntities"
    LoadCommits
    Debug.Print Format$(VBA.DateTime.Timer - start, " 0.0000000 ") '& "LoadCommits"
    LoadTracks
    Debug.Print Format$(VBA.DateTime.Timer - start, " 0.0000000 ") '& "LoadTracks"
    LoadLookupTables
    Debug.Print Format$(VBA.DateTime.Timer - start, " 0.0000000 ") '& "LoadLookupTables"
    LoadDetailTables
    Debug.Print Format$(VBA.DateTime.Timer - start, " 0.0000000 ") '& "LoadDetailTables"
    LoadDetailFields
    Debug.Print Format$(VBA.DateTime.Timer - start, " 0.0000000 ") '& "LoadDetailTables"
End Sub

Private Sub InitializeWithoutTimers()
    LoadEntityTypes
    LoadEntities
    LoadCommits
    LoadTracks
    LoadLookupTables
    LoadDetailTables
    LoadDetailFields
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
        .MapRefsEntityTypes this.EntityTypes
        this.EntityTypes.MapRefsEntities .Self
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
        .MapRefsCommits this.Commits
        this.Commits.MapRefsTracks .Self
    End With
End Sub

Private Sub LoadLookupTables()
    With New LookupTables
        Set this.LookupTables = .Self
        .TableName = modConstants.LOOKUPS_TABLE
        .Load
    End With
End Sub

Private Sub LoadDetailTables()
    With New DetailTables
        Set this.DetailTables = .Self
        .TableName = modConstants.DETAILS_TABLE
        .Load
        .MapRefsEntities this.Entities
        .MapRefsTracks this.Tracks
    End With
End Sub

Private Sub LoadDetailFields()
    With New DetailFields
        Set this.DetailFields = .Self
        .TableName = modConstants.DETAILFIELDS_TABLE
        .Load
        .MapRefsDetailTables this.DetailTables
        .MapRefsLookupTables this.LookupTables
        this.DetailTables.MapRefsFields .Self
    End With
End Sub

