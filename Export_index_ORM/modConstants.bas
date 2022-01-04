Attribute VB_Name = "modConstants"
'@Folder "Common"
Option Compare Database
Option Explicit

Public Const ENTITYTYPES_TABLE As String = "lkpEntityTypes"
Public Const ENTITIES_TABLE As String = "tblEntities"
Public Const COMMITS_TABLE As String = "tblCommits"
Public Const TRACKS_TABLE As String = "tblTrack"
Public Const LOOKUPS_TABLE As String = "metaLookupTables"
Public Const DETAILS_TABLE As String = "metaDetailTables"
Public Const DETAILFIELDS_TABLE As String = "metaSchema"

Public Const ENTITYTYPE_FIELDNAME As String = "EntityType"
Public Const ENTITY_FIELDNAME As String = "Entity"
Public Const TRACK_FIELDNAME As String = "ID" ' No name/title field in this table
Public Const COMMIT_FIELDNAME As String = "Title"
Public Const LOOKUPTABLE_FIELDNAME As String = "TableName"
Public Const DETAILTABLE_FIELDNAME As String = "TableName"

Public Const TRACK_COMMITFK_FIELDNAME As String = "CommitFK"
Public Const TRACK_VALIDFROM_FIELDNAME As String = "ValidFrom"
Public Const TRACK_VALIDUNTIL_FIELDNAME As String = "ValidUntil"
Public Const COMMIT_CLOSED_FIELDNAME As String = "Closed"

Public Const COLLECTION_INDEX_PREFIX As String = "ID#"

Public Const BE_DATABASE_FILENAME As String = "C:\Users\User\Documents\access-entity-scd-db\index_BE.accdb"
Public Const LINKED_DB_CONNECT As String = ";DATABASE="
