Attribute VB_Name = "modConstants"
'@Folder("index")
Option Compare Database
Option Explicit

' Constants
Public Const DETAILS_TABLE As String = "tblDetailTables"
Public Const ENTITIES_TABLE As String = "tblEntities"
Public Const ENTITYTYPES_TABLE As String = "lkpEntityTypes"
Public Const COMMITS_TABLE As String = "tblCommits"

Public Const BACKCOLOR_YELLOW As Double = 13431551
Public Const BACKCOLOR_DEFAULT As Double = 16777215

Public Enum RAGColors
    Default = 16777215
    Red = 13421823
    Yellow = 13431551
    Green = 14282978
End Enum

Public Const COMMIT_FORMNAME As String = "sfrmCommits"
Public Const DETAIL_FORMNAME As String = "sfrmDetails"
Public Const ENTITY_FORMNAME As String = "sfrmEntities"
Public Const NEW_COMMIT_FORM As String = "fdlgCommitNew"

Public Const SCHEMA_FILENAME As String = "C:\Users\User\Documents\access-entity-scd-db\Schema\schema.csv"

Public Const CM_TO_TWIP As Integer = 567
Public Const DEFAULT_HEIGHT As Integer = 360

Public Const QUERY_TRACK_LATEST As String = "qryTrack_Latest"
Public Const SCHEMA_TABLE As String = "metaSchema"

Public Const BE_DATABASE_FILENAME As String = "C:\Users\User\Documents\access-entity-scd-db\index_BE.accdb"
Public Const LINKED_DB_CONNECT As String = ";DATABASE="