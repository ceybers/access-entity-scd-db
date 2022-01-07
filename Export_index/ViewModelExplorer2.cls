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

Private Const ENTITY_PREFIX As String = "Entity#"
Private Const ENTITYTYPE_PREFIX As String = "EntityType#"

Private Type TViewModelExplorer2
    EntityID As Double
    entityTypeid As Double
    DetailTableName As String
    SearchCriteria As String
    Events As collection
    Subform As Subform
    CommitID As Double
    IsEditing As Boolean
End Type

Private this As TViewModelExplorer2

' Ctor
Private Sub Class_Initialize()
    ' TODO below is test env only. Handle unselcted/uninit entityTypeID properly
    this.entityTypeid = 1
    Set this.Events = New collection
    this.CommitID = 1
End Sub

' Properties
Public Property Get IsEditing() As Boolean
    IsEditing = this.IsEditing
End Property

'@Ignore MisleadingByRefParameter
Public Property Set Subform(ByRef RHS As Subform)
    Set this.Subform = RHS
End Property

Public Property Get EntityID() As Double
    EntityID = this.EntityID
End Property

Public Property Let EntityID(ByVal RHS As Double)
    this.EntityID = RHS
End Property

Public Property Get entityTypeid() As Double
    entityTypeid = this.entityTypeid
End Property

Public Property Let entityTypeid(ByVal RHS As Double)
    this.entityTypeid = RHS
End Property

Public Property Get DetailSubFormName() As String
    DetailSubFormName = Replace(this.DetailTableName, "tbl", "sfrm")
End Property

Public Property Get DetailTableName() As String
    DetailTableName = this.DetailTableName
End Property

Public Property Let DetailTableName(ByVal RHS As String)
    this.DetailTableName = RHS
End Property

Public Property Get SearchCriteria() As String
    SearchCriteria = this.SearchCriteria
End Property

Public Property Let SearchCriteria(ByVal RHS As String)
    this.SearchCriteria = RHS
End Property

Public Function GetSQL() As String
    Debug.Assert this.entityTypeid > 0
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


Public Function AddEvents(cb As CommandButton, State As String)
    With New Explorer2EventCommandButton
        Set .CommandButton = cb
        Set .Context = Me
        .State = State
        this.Events.Add .Self
    End With
    'Debug.Print "AddEvents(" & cb.name & ", " & State & ")"
End Function

Public Sub HandleEvent(ByVal State As String)
    'Debug.Print "HandleEvent(" & State & ")"
    Select Case State
        Case "EditNew"
            DetailEditNew
        Case "Save"
            DetailSave
        Case "ViewMany"
            Explorer2DetailSubform.ViewMany this.Subform
        Case "ViewOne"
            Explorer2DetailSubform.ViewOne this.Subform
    End Select
End Sub

Private Sub DetailSave()
    RaiseEvent StopEditing
    this.IsEditing = False
    this.Subform.Requery
    Explorer2DetailSubform.ViewOne this.Subform
End Sub

Private Sub DetailEditNew()
    Dim frm As Form
    
    Explorer2DetailSubform.EditNew this.Subform
    Set frm = this.Subform.Form
    frm.Undo
    
    'DoCmd.GoToRecord acDataForm, frm.name, acNewRec
    frm.controls("ValidUntil").Value = #12/31/9999#
    frm.controls("CommitFK").Value = this.CommitID
    
    RaiseEvent StartEditing
    this.IsEditing = False
End Sub

Private Function DoesSubformHaveZeroRecords() As Boolean
    Debug.Assert Not this.Subform Is Nothing
    Debug.Assert Not this.Subform.Form Is Nothing
    Dim recordCount As Long
    recordCount = this.Subform.Form.RecordsetClone.recordCount
    DoesSubformHaveZeroRecords = (recordCount = 0)
End Function
