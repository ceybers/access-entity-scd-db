VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRecordCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Folder("index_ORM")
'@Interface

Public Property Get Items() As Collection
End Property

Public Property Get NewEnum() As IUnknown
End Property

Public Property Get Count() As Long
End Property

Public Property Get TableName() As String
End Property

Public Property Let TableName(ByVal TableName As String)
End Property

Public Property Get Self() As EntityTypes
End Property

Public Sub Load()
End Sub

Public Sub Clear()
End Sub
