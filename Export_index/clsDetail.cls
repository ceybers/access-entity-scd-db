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
    TableName As String
End Type
Private This As TDetail

Public Property Get Self() As clsDetail
    Set Self = Me
End Property

Public Property Get ID() As Double
    ID = This.ID
End Property

Public Property Get name() As String
    name = This.name
End Property

Private Property Get IThing_ID() As Double
    IThing_ID = This.ID
End Property

Private Property Get IThing_Name() As String
    IThing_Name = This.name
End Property

Public Property Get TableName() As String
    TableName = This.TableName
End Property

Public Property Let ID(ByVal vNewValue As Double)
    This.ID = vNewValue
End Property

Public Property Let name(ByVal vNewValue As String)
    This.name = vNewValue
End Property

Public Property Let TableName(ByVal vNewValue As String)
    This.TableName = vNewValue
End Property

