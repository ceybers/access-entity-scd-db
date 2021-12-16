VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Events
Public Event DetailChanged()
Public Event EntityChanged()
Public Event EntityTypeChanged()

' Private Variables
Private Type TViewModel
     SelectedDetail As clsDetail
     SelectedEntity As clsEntity
     SelectedEntityType As clsEntityType
     Model As clsModel
End Type
Private This As TViewModel

' Ctor
Private Sub Class_Initialize()
    Debug.Print "ViewModel initialised"
    Set This.Model = New clsModel
End Sub

' Getters
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
    Detail = GetThingByID(This.Model.Details, idx)
End Function

Public Function SetEntityByID(idx As Double)
    Entity = GetThingByID(This.Model.Entities, idx)
End Function

Public Function SetEntityTypeByID(idx As Double)
    EntityType = GetThingByID(This.Model.EntityTypes, idx)
End Function

Public Function ReplaceSubform(ByRef sfrm As SubForm)
    sfrm.SourceObject = Replace(Detail.TableName, "tbl", "sfrm")
    sfrm.Form.Filter = "EntityFK = " & Entity.ID
    sfrm.Form.FilterOn = True
    sfrm.SetFocus
    DoCmd.GoToRecord Record:=acLast
End Function

Public Function ApplyEntitiesSourceToListbox(ByRef lb As ListBox)
    lb.RowSource = "SELECT * FROM tblEntities WHERE EntityType = " & EntityType.ID & ";"
End Function

Public Function ApplyDetailsSourceToListbox(ByRef lb As ListBox)
    lb.RowSource = "SELECT DetailTableFK,DetailTable FROM qryEntityTypeToDetail WHERE EntityTypeFK = " & EntityType.ID & ";"
End Function

' Private Methods
Private Function GetThingByID(thingCollection As Collection, idx As Double) As IThing
    Dim thisThing As IThing
    For Each thisThing In thingCollection
        If thisThing.ID = idx Then
            Set GetThingByID = thisThing
            Exit Function
        End If
    Next thisThing
    Err.Raise 5, , "GetThingByID failed!"
End Function
