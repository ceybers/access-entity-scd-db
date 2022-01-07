Attribute VB_Name = "modTrackHelpers"
'@Folder("Helpers")
Option Compare Database
Option Explicit

Public Function CreateNewTrackRecord(CommitFK As Long) As Long
    Dim db As Database
    Dim rs As Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM " & TRACKS_TABLE, dbOpenDynaset, dbSeeChanges)
    
    rs.AddNew
    CreateNewTrackRecord = rs.fields("ID").Value
    rs.fields(TRACK_COMMITFK_FIELDNAME) = CommitFK
    rs.fields(TRACK_VALIDFROM_FIELDNAME) = Now()
    rs.fields(TRACK_VALIDUNTIL_FIELDNAME) = #12/31/9999#
    rs.Update
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function
