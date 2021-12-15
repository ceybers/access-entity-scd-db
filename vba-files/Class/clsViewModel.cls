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
Public Property Get Detail() As IThing
    Debug.Assert Not This.SelectedDetail Is Nothing
    Set Detail = This.SelectedDetail
End Property

Public Property Get Entity() As IThing
    Debug.Assert Not This.SelectedEntity Is Nothing
    Set Entity = This.SelectedEntity
End Property

Public Property Get EntityType() As IThing
    Debug.Assert Not This.SelectedEntityType Is Nothing
    Set EntityType = This.SelectedEntityType
End Property

' Letters
Public Property Let Detail(ByRef vNewValue As IThing)
    Set This.SelectedDetail = vNewValue
    RaiseEvent DetailChanged
End Property

Public Property Let Entity(ByRef vNewValue As IThing)
    Set This.SelectedEntity = vNewValue
    RaiseEvent EntityChanged
End Property

Public Property Let EntityType(ByRef vNewValue As IThing)
    Set This.SelectedEntityType = vNewValue
    RaiseEvent EntityTypeChanged
End Property

' Methods
Public Function SetDetailByID(idx As Double)
    Set Detail = SetThingByID(This.Model.Entities, idx)
End Function

Public Function SetEntityByID(idx As Double)
    Set Entity = SetThingByID(This.Model.Entities, idx)
End Function

Public Function SetEntityTypeByID(idx As Double)
    ' TODO Fix this
    Debug.Print "EntityType is setting..."
    Dim thing As IThing
    Set thing = GetThingByID(This.Model.EntityTypes, idx)
    Debug.Assert Not thing Is Nothing
    'Set EntityType = GetThingByID(This.Model.EntityTypes, idx)
    Debug.Print "EntityType is set"
    Set This.SelectedEntityType = thing
    RaiseEvent EntityTypeChanged
End Function

' Private Methods
Private Function GetThingByID(thingCollection As Collection, idx As Double) As IThing
    Dim thisThing As IThing
    For Each thisThing In thingCollection
        If thisThing.ID = idx Then
            Debug.Print "Found a thing"
            Set GetThingByID = thisThing
            Exit Function
        End If
    Next thisThing
    Err.Raise 5, , "GetThingByID failed!"
End Function
