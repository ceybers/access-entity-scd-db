VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ViewModelExplorer2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "Explorer2"
Option Compare Database
Option Explicit

Public Event StartEditing()
Public Event StopEditing()
Public Event CancelEditing()
Public Event EntityChanged(ByVal entityID As Long)
Public Event DetailChanged(ByVal DetailTableName As String)

Private Const ENTITY_PREFIX As String = "Entity#"
Private Const ENTITYTYPE_PREFIX As String = "EntityType#"

Private Type TViewModelExplorer2
    entityID As Double
    entityTypeid As Double
    DetailTableName As String
    SearchCriteria As String
    Events As collection
    Subform As Subform
    CommitID As Double
    IsEditing As Boolean
End Type

Private This As TViewModelExplorer2

' Ctor
Private Sub Class_Initialize()
    ' TODO below is test env only. Handle unselcted/uninit entityTypeID properly
    This.entityTypeid = 1
    Set This.Events = New collection
    This.CommitID = 1
End Sub

' Properties
Public Property Get IsEditing() As Boolean
    IsEditing = This.IsEditing
End Property

'@Ignore MisleadingByRefParameter
Public Property Set Subform(ByRef rhs As Subform)
    Set This.Subform = rhs
End Property

Public Property Get entityID() As Double
    entityID = This.entityID
End Property

Public Property Let entityID(ByVal rhs As Double)
    If This.IsEditing Then
        MsgBox "Cannot change Entity while editing!"
        RaiseEvent EntityChanged(This.entityID) ' This resets the TreeView control
        Exit Property
    End If
    
    If This.entityID = rhs Then
        Exit Property
    End If
    
    This.entityID = rhs
    RaiseEvent EntityChanged(rhs)
End Property

Public Property Get entityTypeid() As Double
    entityTypeid = This.entityTypeid
End Property

Public Property Let entityTypeid(ByVal rhs As Double)
    This.entityTypeid = rhs
End Property

Public Property Get DetailSubFormName() As String
    DetailSubFormName = Replace(This.DetailTableName, "tbl", "sfrm")
End Property

Public Property Get DetailTableName() As String
    DetailTableName = This.DetailTableName
End Property

Public Property Let DetailTableName(ByVal rhs As String)
    If This.DetailTableName = rhs Then
        Exit Property
    End If
    
    This.DetailTableName = rhs
    RaiseEvent DetailChanged(rhs)
End Property

Public Property Get SearchCriteria() As String
    SearchCriteria = This.SearchCriteria
End Property

Public Property Let SearchCriteria(ByVal rhs As String)
    This.SearchCriteria = rhs
End Property

Public Function GetSQL() As String
    Debug.Assert This.entityTypeid > 0
    GetSQL = "SELECT * FROM metaEntities WHERE EntityType <= " & entityTypeid & " ORDER BY EntityType ASC, Entity ASC"
End Function

Public Function GetRecordset() As Recordset
    Set GetRecordset = CurrentDb.OpenRecordset(GetSQL, dbOpenSnapshot, dbReadOnly)
End Function

' ---
Public Function GetCollectionFromRecordset() As collection
    Dim rs As Recordset
    Dim result As collection
    Dim ent As Explorer2Entity
    
    Set result = New collection
    Set rs = GetRecordset
    
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            Set ent = RecordToEntity(rs)
            result.Add ent, ent.ID
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
    
    Set GetCollectionFromRecordset = result
End Function

Private Function RecordToEntity(ByRef rs As Recordset) As Explorer2Entity
    Set RecordToEntity = New Explorer2Entity
    With RecordToEntity
        .ID = (ENTITY_PREFIX & rs.fields("ID").Value)
        .Entity = rs.fields("Entity").Value
        .Parent = ENTITY_PREFIX & rs.fields("ParentFK").Value
        .EntityType = ENTITYTYPE_PREFIX & rs.fields("EntityType").Value
    End With
End Function


Public Sub AddEvents(cb As CommandButton, State As String)
    With New Explorer2EventCommandButton
        Set .CommandButton = cb
        Set .Context = Me
        .State = State
        This.Events.Add .Self
    End With
    'Debug.Print "AddEvents(" & cb.name & ", " & State & ")"
End Sub

Public Sub HandleEvent(ByVal State As String)
    'Debug.Print "HandleEvent(" & State & ")"
    Select Case State
        Case "EditNew"
            DetailEditNew
        Case "Save"
            DetailSave
        Case "ViewMany"
            Explorer2DetailSubform.ViewMany This.Subform
        Case "ViewOne"
            Explorer2DetailSubform.ViewOne This.Subform
    End Select
End Sub

Private Sub DetailSave()
    RaiseEvent StopEditing
    This.IsEditing = False
    This.Subform.Requery
    Explorer2DetailSubform.ViewOne This.Subform
End Sub

Private Sub DetailEditNew()
    Dim frm As Form
    
    Explorer2DetailSubform.EditNew This.Subform
    Set frm = This.Subform.Form
    frm.Undo
    frm.dataentry = True
    frm.Recordset.AddNew
    'DoCmd.GoToRecord acDataForm, frm.name, acNewRec
    frm.controls("ValidUntil").Value = #12/31/9999#
    frm.controls("CommitFK").Value = This.CommitID
    
    RaiseEvent StartEditing
    This.IsEditing = True
End Sub

Private Function DoesSubformHaveZeroRecords() As Boolean
    Debug.Assert Not This.Subform Is Nothing
    Debug.Assert Not This.Subform.Form Is Nothing
    Dim recordCount As Long
    recordCount = This.Subform.Form.RecordsetClone.recordCount
    DoesSubformHaveZeroRecords = (recordCount = 0)
End Function
