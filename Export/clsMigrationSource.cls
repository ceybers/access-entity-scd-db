VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMigrationSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Type TClsMigrationSource
    tableName As String
    idField As String
    nameField As String
    parentField As String
    parentTableName As String
    entityTypeID As Double
    parentPK As String
End Type

Private this As TClsMigrationSource

Public Property Get tableName() As String
    tableName = this.tableName
End Property

Public Property Get idField() As String
    idField = this.idField
End Property

Public Property Get nameField() As String
    nameField = this.nameField
End Property

Public Property Get parentField() As String
    parentField = this.parentField
End Property

Public Property Get parentTableName() As String
    parentTableName = this.parentTableName
End Property

Public Property Get entityTypeID() As Double
    entityTypeID = this.entityTypeID
End Property

Public Property Get parentPK() As String
    parentPK = this.parentPK
End Property

Public Function Self() As clsMigrationSource
    Set Self = Me
End Function

Public Function SetValues(tableName As String, idField As String, nameField As String, entityTypeID As Double, Optional parentField As String = "", Optional parentTableName As String = "", Optional parentPK As String = "")
    With this
        .tableName = tableName
        .idField = idField
        .nameField = nameField
        .parentField = parentField
        .parentTableName = parentTableName
        .entityTypeID = entityTypeID
        .parentPK = parentPK
    End With
End Function
    
