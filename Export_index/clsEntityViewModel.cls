VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEntityViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder "ViewModels"
Option Compare Database
Option Explicit

' Events
Public Event DetailChanged()
Public Event EntityChanged()
Public Event EntityTypeChanged()
Public Event CommitChanged()

' Private Variables
Private Type TViewModel
     SelectedDetail As clsDetail
     SelectedEntity As clsEntity
     SelectedEntityType As clsEntityType
     Model As clsDataModel
End Type
Private This As TViewModel

' Ctor
Private Sub Class_Initialize()
    Set This.Model = New clsDataModel
End Sub

' Getters
Public Property Get IsValid() As Boolean
    If This.SelectedDetail Is Nothing Then Exit Property
    If This.SelectedEntity Is Nothing Then Exit Property
    IsValid = True
End Property

Public Property Get Detail() As clsDetail
    Debug.Assert Not This.SelectedDetail Is Nothing
    Set Detail = This.SelectedDetail
End Property

Public Property Get Entity() As clsEntity
    Debug.Assert Not This.SelectedEntity Is Nothing
    Set Entity = This.SelectedEntity
End Property

Public Property Get EntityType() As clsEntityType
    Debug.Assert Not This.SelectedEntityType Is Nothing
    Set EntityType = This.SelectedEntityType
End Property

' Letters
Public Property Let Detail(ByRef vNewValue As clsDetail)
    If vNewValue Is Nothing Then Exit Property
    If Not This.SelectedDetail Is Nothing Then
        If Detail.ID = vNewValue.ID Then
            Exit Property
        End If
    End If

    Set This.SelectedDetail = vNewValue
    RaiseEvent DetailChanged
End Property

Public Property Let Entity(ByRef vNewValue As clsEntity)
    If vNewValue Is Nothing Then Exit Property
    If Not This.SelectedEntity Is Nothing Then
        If Entity.ID = vNewValue.ID Then
            Exit Property
        End If
    End If
    
    Set This.SelectedEntity = vNewValue
    RaiseEvent EntityChanged
End Property

Public Property Let EntityType(ByRef vNewValue As clsEntityType)
    If vNewValue Is Nothing Then Exit Property
    If Not This.SelectedEntityType Is Nothing Then
        If EntityType.ID = vNewValue.ID Then
            Exit Property
        End If
    End If
    
    Set This.SelectedEntityType = vNewValue
    RaiseEvent EntityTypeChanged
End Property

' Methods
Public Function SetDetailByID(idx As Double)
    Detail = This.Model.GetThingByID(This.Model.Details, idx)
End Function

Public Function SetEntityByID(idx As Double)
    Entity = This.Model.GetThingByID(This.Model.Entities, idx)
End Function

Public Function SetEntityTypeByID(idx As Double)
    EntityType = This.Model.GetThingByID(This.Model.EntityTypes, idx)
End Function

' Apply Functions for Form Controls
Public Function ApplyEntitiesSourceToListbox(ByRef lb As ListBox)
    lb.RowSource = "SELECT * FROM " & ENTITIES_TABLE & " WHERE EntityType = " & EntityType.ID & ";"
    lb.RowSourceType = "Table/Query"
End Function

Public Function ApplyDetailsSourceToListbox(ByRef lb As ListBox)
    lb.RowSource = "SELECT DetailTableFK,DetailTable FROM qryEntityTypeToDetail WHERE EntityTypeFK = " & EntityType.ID & ";"
End Function

Public Function ApplyCommitsSourceToCombobox(ByRef cmb As ComboBox)
    Err.Raise 5, , "DEPREC"
    cmb.RowSource = vbNullString
    cmb.RowSource = "SELECT TOP 3 ID, Title, Modified FROM " & COMMITS_TABLE & " WHERE Closed Is Null ORDER BY Created DESC;"
End Function

Public Function SearchEntity(tb As textbox, lb As ListBox)
    Dim filteredEntities As Collection
    Dim Entity As clsEntity
    Dim i As Integer
    Dim criteria As String
    
    If IsNull(tb) Then
        ApplyEntitiesSourceToListbox lb
        Exit Function
    End If
    
    If tb.Value = vbNullString Then
        ApplyEntitiesSourceToListbox lb
        Exit Function
    End If
    
    Set filteredEntities = New Collection
    tb.SetFocus
    criteria = tb.text & "*"
    
    For Each Entity In This.Model.Entities
        If Entity.name Like criteria Then
            If (This.SelectedEntityType Is Nothing) Then
                filteredEntities.Add Entity
            ElseIf Entity.EntityType = This.SelectedEntityType.ID Then
                filteredEntities.Add Entity
            End If
        End If
    Next Entity
    
    lb.RowSource = vbNullString
    lb.RowSourceType = "Value List"
    
    For i = lb.ListCount - 1 To 0 Step -1
        lb.RemoveItem i
    Next i
    
    For Each Entity In filteredEntities
        lb.AddItem CStr(Entity.ID) & ";" & Entity.name
    Next Entity
End Function
