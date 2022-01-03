VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Folder("index_ORM")
'@Interface

Public Property Get ID() As Double
End Property

Public Property Let ID(ByVal RHS As Double)
End Property

Public Property Get Name() As String
End Property

Public Property Let Name(ByVal RHS As String)
End Property

Public Function Create(ByRef rs As Recordset) As IRecord
End Function

