Attribute VB_Name = "modBuildTablesForDetails"

'@Folder "Provisioning"
Option Compare Database
Option Explicit

Private Type TControlSet
    FieldName As String
    Caption As String
    Width As String
    LookupTable As String
    Suffix As String
    Format As String
    Textalign As String
End Type

Public Sub BuildTablesForDetails()
    'If MsgBox("Build tables?", vbYesNo + vbDefaultButton2) = vbNo Then
    '    Exit Sub
    'End If
    
    Dim tables As collection
    
    Debug.Print "Getting list of detail tables from metaSchema..."
    Set tables = GetListOfTablesFromSchema
    Debug.Print " " & tables.count & " table(s) found"
    Debug.Print
    
    Debug.Print "Filtering to include only empty tables..."
    'Set tables2 = FilterEmptyTablesOnly(tables)
    'Debug.Print " " & tables2.count & " table(s) found"
    'Debug.Print
    
    Debug.Print "Removing tables with 0 records..."
    'dropResult = DropTables(tables2)
    'Debug.Print " " & dropResult & " table(s) dropped"
    'Debug.Print
        
    Debug.Print "Creating tables..."
    'createResult = CreateTables(tables2)
    'Debug.Print " " & createResult & " table(s) created and linked"
    'Debug.Print
        
    Dim formResult As Long
    If MsgBox("Build forms?", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
        Debug.Print "Creating forms..."
        formResult = CreateForms(tables)
        Debug.Print " " & formResult & " forms(s) built"
        Debug.Print
    End If
    
    Debug.Print "END"
End Sub

Private Function CreateForms(tables As collection) As Long
    Dim tbl As Variant
    For Each tbl In tables
        BuildFormForDetail Replace(tbl, "tblDetail", vbNullString)
    Next tbl
    CreateForms = -1 ' TODO Implement or refactor into Sub
End Function

Private Function CreateTables(tables As collection) As Long
    Dim tbl As Variant
    
    For Each tbl In tables
        If CreateTable(CStr(tbl)) Then
            LinkTable (CStr(tbl))
            CreateTables = CreateTables + 1
        End If
    Next tbl
End Function

Private Function CreateTable(tableName As String) As Boolean
    Dim db As Database
    Dim tbl As TableDef
    
    Set db = OpenDatabase(BE_DATABASE_FILENAME, False, False)
    
    Set tbl = db.CreateTableDef(tableName)

    CreateIDField tbl
    
    CreateGenericField tbl, "EntityFK", dbLong ' TODO Const this
    CreateGenericField tbl, "TrackFK", dbLong ' TODO Const this
    
    AddFieldsToTableDefFromMetaSchema tbl, tableName
    
    db.TableDefs.Append tbl
    db.TableDefs.Refresh
    
    db.Close
    Set db = Nothing
    
    CreateTable = True
End Function

Private Sub AddFieldsToTableDefFromMetaSchema(tblDef As TableDef, tableName As String)
    Dim db As Database
    Dim rs As Recordset
    Dim sql As String
    
    sql = "SELECT * FROM " & SCHEMA_TABLE & " WHERE TableName = '" & tableName & "';"
    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            AddFieldToTableDefFromMetaSchema tblDef, rs
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub

Private Function AddFieldToTableDefFromMetaSchema(tblDef As TableDef, rs As Recordset)
    Dim fldType As Long
    Dim prop As DAO.Property
    Dim fld As Field
    
    fldType = dbText
    
    Select Case rs.fields("fieldType")
        Case "Number"
            fldType = dbLong
        Case "Date/Time"
            fldType = dbDate
        Case "Double"
            fldType = dbDouble
    End Select
    
    Set fld = tblDef.CreateField(rs!FieldName, fldType)
    
    If Nz(rs!Format) <> vbNullString Then
        Set prop = fld.CreateProperty("Format", dbText)
        prop.Value = rs!Format
        'fld.Properties.Append prop
    End If
    
    If Nz(rs!defaultValue) <> vbNullString Then
        fld.defaultValue = rs!defaultValue
    End If
    
    tblDef.fields.Append fld
End Function

Private Function CreateGenericField(tbl As TableDef, FieldName As String, Optional fieldType As Long = dbText)
    Dim fld As Field
    Set fld = tbl.CreateField(FieldName, fieldType)
    tbl.fields.Append fld
End Function

Private Function CreateIDField(tbl As TableDef, Optional FieldName As String = "ID")
    Dim fld As DAO.Field
    Dim idx As DAO.index
    
    Set fld = CreateAutoNumberField(tbl, FieldName)
    tbl.fields.Append fld
    
    Set idx = tbl.CreateIndex
    
    With idx
        .name = "Primary Key"
        .fields.Append .CreateField(FieldName)
        .Unique = True
        .Primary = True
    End With
    tbl.Indexes.Append idx
End Function

Private Function CreateAutoNumberField(tbl As TableDef, FieldName As String) As Field
    Set CreateAutoNumberField = tbl.CreateField(FieldName, dbLong, 4)
    With CreateAutoNumberField
         .Attributes = dbAutoIncrField
    End With
End Function

Private Function GetListOfTablesFromSchema() As collection
    Dim rs As Recordset
    Dim sql As String
    
    Set GetListOfTablesFromSchema = New collection
    sql = "SELECT DISTINCT TableName FROM " & SCHEMA_TABLE & ";"
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            GetListOfTablesFromSchema.Add CStr(rs!tableName)
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Function FilterEmptyTablesOnly(tables As collection) As collection
    Dim tableName As String
    Dim tbl As Variant
    Dim db As Database
    
    Set db = OpenDatabase(BE_DATABASE_FILENAME, False, False)
    
    Set FilterEmptyTablesOnly = New collection
    
    For Each tbl In tables
        tableName = CStr(tbl)
        If DoesTableExist(tableName, db) Then
            If IsTableEmpty(tableName, db) Then
                FilterEmptyTablesOnly.Add tbl
            End If
        Else
            FilterEmptyTablesOnly.Add tbl
        End If
    Next
    
    db.Close
    Set db = Nothing
End Function
