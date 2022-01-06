VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId

'@Folder "Things"
Option Compare Database
Option Explicit

Implements IThing

Private Type TDetail
    ID As Double
    name As String
    tableName As String
End Type
Private this As TDetail

Public Property Get Self() As clsDetail
    Set Self = Me
End Property

Public Property Get ID() As Double
    ID = this.ID
End Property

Public Property Get name() As String
    name = this.name
End Property

Private Property Get IThing_ID() As Double
    IThing_ID = this.ID
End Property

Private Property Get IThing_Name() As String
    IThing_Name = this.name
End Property

Public Property Get tableName() As String
    tableName = this.tableName
End Property

Public Property Let ID(ByVal vNewValue As Double)
    this.ID = vNewValue
End Property

Public Property Let name(ByVal vNewValue As String)
    this.name = vNewValue
End Property

Public Property Let tableName(ByVal vNewValue As String)
    this.tableName = vNewValue
End Property

