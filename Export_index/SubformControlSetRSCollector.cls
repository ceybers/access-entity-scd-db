VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SubformControlSetRSCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "Provisioning"
Option Compare Database
Option Explicit

Implements IRecordsetCollector

Private Sub IRecordsetCollector_AddRecord(ByRef rs As Recordset, ByRef coll As collection)
    Dim cs As subformcontrolset
    Set cs = subformcontrolset.CreateFromRecordset(rs)
    coll.Add cs
End Sub
