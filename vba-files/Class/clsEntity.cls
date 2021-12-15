VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mID As Double
Private mName As String

Public Property Get ID() As Double
    ID = mID
End Property

Public Property Let ID(ByVal vNewValue As Double)
    mID = vNewValue
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal vNewValue As String)
    mName = vNewValue
End Property

Public Function Create(entityId As Double, detailName As String) As clsEntity
    Dim result As New clsEntity
    result.ID = entityId
    result.Name = detailName
    Set Create = result
End Function
