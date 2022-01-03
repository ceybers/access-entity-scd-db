# Workspace

ORM for SCE data model.

---

## Structure
### Workspace
 #### Interface
  * Item() as Collection<IThing>  [function]
  * Load(tableName) as Boolean [function]
  * Ready() as Boolean [prop]
  * TableName as String [prop]
  * GetByID(n) as IThing [function]
  ' We need a multistage load. One to load individual records from recordsets, then a second to link ByRefs, then possibly a third if anything has a dependency that *all* the items in a collection (or other collections) need ByRefs loaded.

 #### EntityTypes
  * ValueFieldName as String ' i.e. 'EntityType'. Needs this for doing JOINs

 ####  Entities
  * GetByType(`EntityType`) as Collection<Entity> ' Filters the Item collection

 ####  LookupTables
  * n/a

 ####  DetailTables
  *  GetByEntityType(`EntityType`) as Collection<DetailTable> 
  ' This collection level method differs from the item level - it returns a list of DetailTables that are used in an EntityType, whereas the item level returns a list of EntityTypes that are used by a particular DetailTable.
 
 ####  Commits
  * GetOpen() as Collection<Commit>

 ####  Tracks
  * n/a

---

## Classes
### EntityType (abstract)
 * ID
 * Name
 * DetailTables() as Collection<DetailTables> (`tblEntityTypeToDetail`)
 * Entities() as Collection<Entity> 
 ' Can go Workspace.EntityType("Tank").Entities -> list of all entities which are tanks
 ' Alternatively, `Workspace.Entities.GetByType(EntityType("Tank"))`
 * PermissibleParents() as Collection<EntityType> ' NYI
 * PermissibleChildren() as Collection<EntityType> ' NYI
 
### Entity (instance)
 * ID
 * Name
 * EntityType as `ByRef EntityType`
 * Parent as `ByRef Entity`
 * Children as Collection<Entity>
 * Details as Collection<Detail> ' Includes historic
 * LatestDetails as Collection<Detail> ' Latest only! Use same backing field as above
 * ProvisionedDetailTables as Collection<DetailTable>
 ' Do we need a list of DetailTables that have been instantiated for this Entity? i.e. Map(.LatestDetails.DetailTable)

### LookupTable (group)
 * ID ' Use TableDefs index?
 * (Table)Name as String
 * ValueFieldName as String ' Useful for doing SQL JOINS
 * LookupValues as Collection<LookupValue>
 * UsedBy as Collection<DetailField> ' How do we get the parent table's name?
 ' UsedByTables(), UsedByFields() perhaps

### LookupValue (member)
 * ID as Double
 * Value as String ' Note, not `Name` ' This breaks our consistent naming convention. Just use `Name`.
 * Owner as `LookupTable`
 * Details as Collection<Detail> 
 ' All the detail records in all the detail tables where the detail field value == this ID. 
 ' Probably don't need the actual Details, only the count?

### DetailTable (group)
 * ID from TableDefs index
 * (Table)Name as String
 * Fields as Collection<DetailField> 
 * UsedBy as Collection<EntityType>
 * Details as Collection<Detail>
 * LatestDetails as Collection<Detail> ' Same backing collection as Details

### DetailField (record) 
 * ID as Double
 * Name (internal fieldname) as String
 * FieldType as enum
 * Caption as String
 * Lookup as `ByRef LookupTable`
 * Suffix as String
 * DefaultValue, Width, Height, TextAlign, NumberFormat
 * Owner as `DetailTable` ' Necessary, see: LookupTable.UsedBy()
 ' Could apply *TWIPS_TO_CM to w, h
 ' Could apply enums to FieldType, TextAlign
 ' These are the records in `metaSchema` atm

### Commit
 * ID
 * Name
 * IsClosed as `Boolean` ' Useful for filtering for UI ComboBox
 * CommitType? (would need to pull in the special lookup table for it too)
 * Tracks as Collection<Track>

### Track
 * ID
 * ValidFrom
 * ValidUntil
 * Commit as Commit

### Detail (instance of Ent*Det)
 * ID (this is the ID in its' DetailTable) 
 * Entity as Entity
 * DetailTable as DetailTable
 * Track as Track
 * IsLatest as Boolean (i.e. .Track.ValidUntil == #9999/12/31#)
 * IsProvisioned as Boolean (false if no pre-existing records for this Ent*Det)
 ' `SELECT * FROM tblDetailABC WHERE ID = 123`
 ' Some kind of Function for getting the record in the recordset? Or at least a deep clone of the Fields collection?

---

## Code

### EntityTypes

```
RecordsetToCollection(ByVal tableName as string, ByRef something as ISomething, Optional ByRef db as Database)
    Dim rs as Recordset
    Dim result as New Collection
    rs = db.OpenRecordset(tableName, dbOpenSnapshot, dbReadOnly)
    If Not RS.BOF and Not RS.EOF
        Do While Not RS.EOF
            result.Add something.DoSomething(rs)
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
```
We want a common method for recordset -> collection
ISomething_DoSomething(ByRef rs as Recordset) as Variant

Then, EntityTypes can go:
`RecordsetToCollection("lkpEntityTypes", clsEntityTypeFactory)`
and Entities can go:
`RecordsetToCollection("tblEntities", clsEntityFactory)`
Or, we put all the methods in one class, and use CallByName
`CallByName(clsFactory, "CreateEntityType", Array(rs))` 
!! Assuming we can pass ByRef objects to CallByName!

Then we need to get parents/children for setting ByRef properties. This will be unique to classes as the fields differ.


For Each EntityType in EntityTypes
    Set EntityType.Entities = New Collection
    For Each Entity in Entities
        If Entity.??? ' This won't work because the POCO aren't storing the FKs
    Next Entity
    ' `SELECT ID FROM tblEntity WHERE EntityTypeFK = n`
    ' This would need to be run from SCD level, not EntityType level
    ' So, MapEntityTypesToEntities(EntityTypes, Entities)
Next EntityType

Public Function GetIThingInCollectionByName(ByRef coll as Collection, criteria as String) as IThing
    ' Common across all 9x collection classes
End Function

### Dependency Order
Alternatively,
1. Load EntityTypes
2. Load Entities. Since EntityTypes are loaded, EntityFactory can look in EntityTypes.GetByID() to get ByRef of the EntityType instance
3. We can then loop through EntityTypes, and for each EntityType, loop through Entities, and check LHS is RHS on their .EntityType property

Caches:
1. metaLookupTables - load from TableDefs
2. metaDetailTables - load from TableDefs

In practice,
1. EntityTypes RS
2. Entities RS, .EntityType ByRef
3. EntityTypes .Entities ByRef()
4. Commits RS
5. Tracks RS
6. Tracks .Commit ByRef
7. Commits .Tracks ByRef()
8. LookupTables RS
9. LookupValues RS, .Owner ByRef
10. LookupTables .Values ByRef()

Pattern is thus:
1. Load RS, and if it has single ByRefs, load them from dependency via GetByID(n)
2. Load collection ByRefs once the dependency has the ByRefs set

e.g. 
```
EntityType.Entities = ()

Entity.EntityType = EntityTypes.GetByID(rs!EntityTypeFK)

Sub EntityTypes.LoadByRefs()
    For Each EntityType in EntityTypes
        For Each Entity in Entities
            If Entity.EntityType is EntityType Then
                EntityTypes.Entities.Add Entity
            End if
        Next Entity
    Next EntityType
```
---

## Notes

Workspace.Entities.Item("AAA-101"), or Workspace.Entities("AAA-101")

Load needs to:
* Pull `ID, Name, FK` from `Recordset`
* Apply `.Parent` with `ByRef`
* Calculate `FullPath` (if required)

We can compare instances of class objects using `is` operator, the same one that is used for null checking.
e.g. `Entities("AAA-101").EntityType is EntityTypes("Tank")`

```
Enum CustomError
    Error1 = vbObjectError + 1
End Enum

Private Function CustomErrorDescription(errNumber As CustomError) As String
    Select Case errNumber
        Case Error1
            CustomErrorDescription = "My message"
    End Select
End Function
```

Excel Macros *can* be undone, if you implement `Application.Undo Text="", Procedure:="sub"` 

Application.Run can run Public Functions of regular modules: `Call Application.Run("MyFunc", "hi")`

`CallByName` let's us get/set/let properties in an instance of a class via string names.
e.g. 
```
Dim result As Variant
Dim x As String
result = CallByName(something, "Name", VbLet, "Test")
x = CallByName(something, "Name", VbGet)
```

It can also run a predeclared class' methods (i.e. `clsDetail` instead of `set something = new clsDetail`). Use `vbMethod`.