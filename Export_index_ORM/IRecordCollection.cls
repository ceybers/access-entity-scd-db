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

'@Folder "ORM.Interfaces"
'@Interface

Public Property Get Items() As Collection
End Property

Public Property Get Count() As Long
End Property

Public Property Get TableName() As String
End Property

Public Property Let TableName(ByVal TableName As String)
End Property

Public Property Get Self() As IRecordCollection
End Property

Public Sub Load()
End Sub

Public Sub Clear()
End Sub

Public Function GetByID(ByVal ID As Double) As IRecord
End Function
