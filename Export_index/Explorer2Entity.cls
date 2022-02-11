VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Explorer2Entity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Explorer2")
Option Compare Database
Option Explicit

Private Type TExplorer2Entity
    ID As String
    Entity As String
    Parent As String
    EntityType As String
End Type

Private This As TExplorer2Entity

Public Property Get ID() As String
    ID = This.ID
End Property

Public Property Let ID(ByVal rhs As String)
    This.ID = rhs
End Property

Public Property Get Entity() As String
    Entity = This.Entity
End Property

Public Property Let Entity(ByVal rhs As String)
    This.Entity = rhs
End Property

Public Property Get Parent() As String
    Parent = This.Parent
End Property

Public Property Let Parent(ByVal rhs As String)
    This.Parent = rhs
End Property

Public Property Get EntityType() As String
    EntityType = This.EntityType
End Property

Public Property Let EntityType(ByVal rhs As String)
    This.EntityType = rhs
End Property
