# TODO

## Export System
- [x] Since we don't have tblFact, we need to be able to export presets for EntityTypes, including resolving each Detail table's Lookups

## EntityType Management Interface
- [x] Initially Listbox
- [ ] TreeView would be nice, to visualise parent-child hierarchy
- [ ] Set bridge table for `EntityType`-`EntityType` permissible children
- [ ] Set bridge table to `EntityType`-`Details`
    * ListBox with multi-select?
    * ListView with Checkboxes

## Entity Management Interface
- [ ] Use same UI as `frmExplorer?`
- [ ] Provision instances of children
    * List of permissible `EntityTypes` that we may create
    * Create 1* new
    * Create n* new:
        * GUID titles
        * Prefix name 000
        * Paste names into multi-line textbox
    * Create instances of *all* 1+ children that have not yet been instanced and are always/usually required (anonymous floor, shell strakes, roof)
- [ ] So frmEntities needs to show `lstDetails`, and `lstChildEntities`

## Lookup Table Management Interface
- [ ] Need action to Add New Lookuptable
- [ ] Currently we can only add to existing tables

## Entity Structure
- [ ] Will reports be n* instances of Entities, with the Tank as parent?
- [ ] Can we have a `tblDetailsTankReports` where we have fields that lookup to the `Entities` table itself, instead of a `lkp*` table? 
- [ ] Similarly `Floating Roof`, `Tank Strakes`, `MAT CR RL`

## Data Migration
- [ ] Need Squash/Split for when the tables in TankDB3 don't match in SCE DB5, e.g. `Service` + `Dims.OpCapacity` -> `NewService`

## MVVM

 * View (UserForm) -> Model, Commands
 * Model -> Nothing
 * Commands -> Model
 * Presenter -> View, Model

## Dependency Injection

* Concrete: Get Thing, Set Thing
* Interface: Get Thing, DoSomething()

## TODO
- [ ] Add optional "Hidden as Boolean" for `modDBHelpers.LinkTable()`
- [ ] Merge `modDBHelpers` and `modCommon`
- [ ] Split `modCommon.DropTables` into `DropTables` and `DropTable`
- [ ] Implement `DoesQueryExist