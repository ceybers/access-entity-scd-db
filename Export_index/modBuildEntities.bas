Attribute VB_Name = "modBuildEntities"
'@Folder "Main"
Option Compare Database
Option Explicit

Public Sub BuildSampleEntities()
    Dim db As Database
    Dim divID As Long, strmID As Long, deptID As Long
    
    If vbNo = MsgBox("Are you sure?", vbExclamation + vbYesNo + vbDefaultButton2, "Build Sample Entities") Then
        Exit Sub
    End If
    
    Set db = CurrentDb
    db.Execute "DELETE * FROM metaEntities;"

    db.Execute "INSERT INTO metaEntities ([Entity], [ParentFK], [EntityType]) VALUES ('Root', 0, 1)"
    
    Dim i As Long, j As Long, k As Long, l As Long
    For i = 1 To 8
        db.Execute "INSERT INTO metaEntities ([Entity], [ParentFK], [EntityType]) VALUES ('Division " & i & "', 1, 1)"
        divID = db.OpenRecordset("SELECT @@Identity FROM metaEntities")(0)
        For j = 1 To 2
            db.Execute "INSERT INTO metaEntities ([Entity], [ParentFK], [EntityType]) VALUES ('Stream " & i & "." & j & "', " & divID & ", 2)"
            strmID = db.OpenRecordset("SELECT @@Identity FROM metaEntities")(0)
            For k = 1 To 14
                db.Execute "INSERT INTO metaEntities ([Entity], [ParentFK], [EntityType]) VALUES ('Depot " & i & "." & j & "." & k & "', " & strmID & ", 3)"
                deptID = db.OpenRecordset("SELECT @@Identity FROM metaEntities")(0)
                For l = 1 To 8
                    db.Execute "INSERT INTO metaEntities ([Entity], [ParentFK], [EntityType]) VALUES ('Tank " & i & "." & j & "." & k & "." & l & "', " & deptID & ", 4)"
                Next l
            Next k
        Next j
    Next i
    
    Set db = Nothing
    
    Dim n As Long
    n = 8 * 2 * 14 * 8
    MsgBox "Built " & n & " entities OK", vbInformation + vbOKOnly, "Build Sample Entities"
End Sub
