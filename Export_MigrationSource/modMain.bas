Attribute VB_Name = "modMain"
'@Folder("MigrationSource")
Option Compare Database
Option Explicit

Public Sub Main()
    ResetMigration
    StartMigration
End Sub

Public Sub StartMigration()
    ' Todo add start/stop messages
    ' Optional duplicate sub with timers
    MigrateEntities
    MigrateCommits
    MigrateTracks
    MigrateLookups
    MigrateDetails
End Sub

Public Sub ResetMigration()
    ' TODO Ask first
    ResetSourceTables
    ResetDestinationTables
End Sub

Private Sub ResetSourceTables()
    Dim tables As String
    Dim tbl As Variant
    Dim sql As String
    
    tables = "lkpDesignElevation;lkpDesignFloating;tblBusStream;tblDepot;tblDesign;tblDivision;tblTankID;tblTracking;tblUpdRef"

    For Each tbl In Split(tables, ";")
        sql = "UPDATE " & tbl & " SET MigrationID = NULL"
        Call CurrentDb.Execute(sql)
    Next tbl
End Sub

Private Sub ResetDestinationTables()
    Dim tables As String
    Dim tbl As Variant
    Dim sql As String
    
    tables = "lkpTankDesignElevation;lkpTankDesignFloatingRoof;tblCommits;tblDetailTankDesign;tblEntities;tblTrack"

    For Each tbl In Split(tables, ";")
        sql = "DELETE * FROM " & tbl
        Call CurrentDb.Execute(sql)
    Next tbl
End Sub
