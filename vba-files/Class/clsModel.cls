VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Event DetailChanged()
Public Event EntityChanged()
Public Event EntityTypeChanged()

Private mSelectedDetail As clsDetail
Private mSelectedEntity As clsEntity
Private mSelectedEntityType As clsEntityType

Public Property Get Detail() As clsDetail
    Set Detail = mSelectedDetail
End Property

Public Property Let Detail(ByRef vNewValue As clsDetail)
    Set mSelectedDetail = vNewValue
    RaiseEvent DetailChanged
End Property

Public Property Get Entity() As clsEntity
    Set Entity = mSelectedEntity
End Property

Public Property Let Entity(ByRef vNewValue As clsEntity)
    Set mSelectedEntity = vNewValue
    RaiseEvent EntityChanged
End Property

Public Property Get EntityType() As clsEntityType
    Set EntityType = mSelectedEntityType
End Property

Public Property Let EntityType(ByRef vNewValue As clsEntityType)
    Set mSelectedEntityType = vNewValue
    RaiseEvent EntityTypeChanged
End Property

