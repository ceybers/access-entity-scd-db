# TODO for ORM

Current performance: 961ms to load ORM. 

- [ ] Fix abysmal performance by not open recordsets 1-by-1 on every iteration of a for-each loop on a collection.
- [ ] Check if Tracks module is a performance issue due to Date field type casting.
- [ ] Entity class needs .Details 
- [ ] Entity class possibly needs .DetailTables 
- [ ] DetailTable needs .DetailFields
- [ ] DetailField needs LookupTable