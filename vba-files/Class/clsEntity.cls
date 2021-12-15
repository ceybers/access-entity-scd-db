VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mID As Double
Private mName As String

Public Property Get id() As Double
    id = mID
End Property

Public Property Let id(ByVal vNewValue As Double)
    mID = vNewValue
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal vNewValue As String)
    mName = vNewValue
End Property

