Attribute VB_Name = "modManageLinkedData"
'@Folder("Provisioning")
Option Compare Database
Option Explicit

Public Sub RemoveLinkedTables()
    Dim tdf As TableDef
    For Each tdf In CurrentDb.TableDefs
        If tdf.Name Like "lkp*" Or tdf.Name Like "tblDetail*" Then
            If Len(tdf.Connect) > 0 Then
                CurrentDb.TableDefs.Delete tdf.Name
            End If
        End If
    Next tdf
End Sub

Public Sub AddLinkedTables()
    Dim tdf As TableDef
    Dim db As Database
    Set db = OpenDatabase(BE_DATABASE_FILENAME, False, True)
    For Each tdf In db.TableDefs
        If tdf.Name Like "lkp*" Or tdf.Name Like "tblDetail*" Then
            LinkTable tdf.Name
        End If
    Next tdf
End Sub
