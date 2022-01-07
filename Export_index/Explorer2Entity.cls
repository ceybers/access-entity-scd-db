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

Private this As TExplorer2Entity

Public Property Get ID() As String
    ID = this.ID
End Property

Public Property Let ID(ByVal RHS As String)
    this.ID = RHS
End Property

Public Property Get Entity() As String
    Entity = this.Entity
End Property

Public Property Let Entity(ByVal RHS As String)
    this.Entity = RHS
End Property

Public Property Get Parent() As String
    Parent = this.Parent
End Property

Public Property Let Parent(ByVal RHS As String)
    this.Parent = RHS
End Property

Public Property Get EntityType() As String
    EntityType = this.EntityType
End Property

Public Property Let EntityType(ByVal RHS As String)
    this.EntityType = RHS
End Property
