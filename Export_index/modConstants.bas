Attribute VB_Name = "modConstants"
'@Folder "Common"
Option Compare Database
Option Explicit

Public Enum RAGColors
    Default = 16777215
    Red = 13421823
    Yellow = 13431551
    Green = 14282978
    Grey = 15921906
End Enum

Public Const ENTITYTYPES_TABLE As String = "metaEntityTypes"
Public Const ENTITIES_TABLE As String = "metaEntities"
Public Const COMMITS_TABLE As String = "metaCommits"
Public Const TRACKS_TABLE As String = "metaTrack"
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

Public Const QUERY_TRACK_LATEST As String = "qryTrack_Latest"
Public Const SCHEMA_TABLE As String = "metaSchema"

Public Const BE_DATABASE_FILENAME As String = "C:\Users\User\Documents\access-entity-scd-db\index_BE.accdb"
Public Const LINKED_DB_CONNECT As String = ";DATABASE="

Public Const COMMIT_FORMNAME As String = "sfrmCommits"
Public Const DETAIL_FORMNAME As String = "sfrmDetails"
Public Const ENTITY_FORMNAME As String = "sfrmEntities"
Public Const NEW_COMMIT_FORM As String = "fdlgCommitNew"

Public Const SCHEMA_FILENAME As String = "C:\Users\User\Documents\access-entity-scd-db\Schema\schema.csv" ' TODO Check if still required

Public Const CM_TO_TWIP As Long = 567
Public Const DEFAULT_HEIGHT As Long = 360
