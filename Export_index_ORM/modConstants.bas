Attribute VB_Name = "modConstants"
'@Folder "Common"
Option Compare Database
Option Explicit

Public Const ENTITYTYPES_TABLE As String = "lkpEntityTypes"
Public Const ENTITIES_TABLE As String = "tblEntities"
Public Const DETAILS_TABLE As String = "tblDetailTables" ' TODO rename to metaDetailTables
Public Const COMMITS_TABLE As String = "tblCommits"
Public Const TRACKS_TABLE As String = "tblTrack"
' metaLookupTables

Public Const ENTITYTYPE_FIELDNAME As String = "EntityType"

Public Const BE_DATABASE_FILENAME As String = "C:\Users\User\Documents\access-entity-scd-db\index_BE.accdb"
Public Const LINKED_DB_CONNECT As String = ";DATABASE="