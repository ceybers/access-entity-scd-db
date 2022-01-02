# clsDetailForm3

## Host Form (frmTest3)

Main form with buttons, and Sub-form with the sfrmDetail*

frmTest3 has 7 buttons:
* Add New, Save New, Cancel New
* View Existing, Edit Existing, Save Edits, Cancel Edits

frmTest3 needs to set Properties in clsDetailForm3/DetailVM:
* TableName or SubformName or SubformInstance 
* EntityID
* CommitID
(#1 and #2 reset State to Ready/Check, #3 does not)

Currently VM's Events (EditingStarted/Stopped/Cancelled; and StateChanged) just log details to console and update buttons - if we give the VM references to the buttons we can update them in clsDetailForm3 instead.

clsEventButton, with property for Access.CommandButton, and DetailState3 (i.e. action name)

`CancelEditing()` and `SaveEdits()` should be changed to States (Cancelling, Saving). They are already common between New and Existing.


## View Model (clsDetailForm3)

### TODO

* Remove events (or change to Log()s)
* Previous makes it easier for us to wrap VM in a Static instance Getter
* Add EventCommandButton collection
* Error handle `CheckForExistingRecords()`
* LHS and RHS loops seem to only use .Locked, .Visible, and .BackColor
* Suffix loop only used .Left
* Error handling on `RemoveCancelledNewRecord()`
* Refactor `GetButtonState(View)`
* Add function to get instance in EventCmdBtn collection by "name" (state)
