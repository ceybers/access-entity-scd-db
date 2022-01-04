Attribute VB_Name = "modPlaceholderRecords"
'@Folder("ORM")
Option Compare Database
Option Explicit

Public Function CreateCommit(ByVal ID As Double) As IRecord
    Set CreateCommit = New Commit
    CreateCommit.ID = ID * -1
End Function

Public Function CreateDetailField(ByVal ID As Double) As IRecord
    Set CreateDetailField = New DetailField
    CreateDetailField.ID = ID * -1
End Function

Public Function CreateDetailTable(ByVal ID As Double) As IRecord
    Set CreateDetailTable = New DetailTable
    CreateDetailTable.ID = ID * -1
End Function

Public Function CreateEntity(ByVal ID As Double) As IRecord
    Set CreateEntity = New Entity
    CreateEntity.ID = ID * -1
End Function

Public Function CreateEntityType(ByVal ID As Double) As IRecord
    Set CreateEntityType = New EntityType
    CreateEntityType.ID = ID * -1
End Function

Public Function CreateLookupTable(ByVal ID As Double) As IRecord
    Set CreateLookupTable = New LookupTable
    CreateLookupTable.ID = ID * -1
End Function

Public Function CreateLookupValue(ByVal ID As Double) As IRecord
    Set CreateLookupValue = New LookupValue
    CreateLookupValue.ID = ID * -1
End Function

Public Function CreateTrack(ByVal ID As Double) As IRecord
    Set CreateTrack = New Track
    CreateTrack.ID = ID * -1
End Function
