# TODO for ORM

Current performance: 961ms to load ORM. 

Benchmark reduced to 260ms by replacing `f(tableName, ID, "FieldName")` with dummy objects, where the ID is stored negatively, and the actual object is mapped later. This lets us temporarily store the Foreign Key without needing separate fields.

- [x] Fix abysmal performance by not open recordsets 1-by-1 on every iteration of a for-each loop on a collection.
- [ ] Check if Tracks module is a performance issue due to Date field type casting.
- [ ] Entity class needs .Details 
- [ ] Entity class possibly needs .DetailTables 
- [x] DetailTable needs .DetailFields
- [x] DetailField needs LookupTable