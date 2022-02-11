VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId

'@Folder "Things"
Option Compare Database
Option Explicit

Implements IThing

Private Type TEntity
    ID As Double
    Name As String
    EntityType As Double
End Type
Private This As TEntity

Public Property Get Self() As clsEntity
    Set Self = Me
End Property

Public Property Get ID() As Double
    ID = This.ID
End Property

Public Property Get Name() As String
    Name = This.Name
End Property

Public Property Get EntityType() As Double
    EntityType = This.EntityType
End Property

Private Property Get IThing_ID() As Double
    IThing_ID = This.ID
End Property

Private Property Get IThing_Name() As String
    IThing_Name = This.Name
End Property

Public Property Let ID(ByVal vNewValue As Double)
    This.ID = vNewValue
End Property

Public Property Let Name(ByVal vNewValue As String)
    This.Name = vNewValue
End Property

Public Property Let EntityType(ByVal vNewValue As Double)
    This.EntityType = vNewValue
End Property

