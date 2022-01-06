Attribute VB_Name = "modTest"
'@Folder("index")
Option Compare Database
Option Explicit

Public Sub TEST()
    Dim db As Database
    Dim rs As Recordset
    Dim divID As Long, strmID As Long, deptID As Long
    
    Set db = CurrentDb
    db.Execute "DELETE * FROM metaEntities;"

    Dim i As Integer, j As Integer, k As Integer, l As Integer
    For i = 1 To 8
        db.Execute "INSERT INTO metaEntities ([Entity], [ParentFK], [EntityType]) VALUES ('Division " & i & "', 0, 1)"
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
    
    'rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub
