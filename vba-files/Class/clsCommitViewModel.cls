VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommitViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Event RecordIDChanged()

Private Type TCommitViewModel
    AutofillType As String
    AutofillData As String
    RecordID As Double
End Type
Private This As TCommitViewModel

Public Property Get AutofillType() As String
    AutofillType = This.AutofillType
End Property

Public Property Let AutofillType(ByVal RHS As String)
    This.AutofillType = RHS
End Property

Public Property Get AutofillData() As String
    AutofillData = This.AutofillData
End Property

Public Property Let AutofillData(ByVal RHS As String)
    This.AutofillData = RHS
End Property

Public Property Get RecordID() As Double
    RecordID = This.RecordID
End Property

Public Property Let RecordID(RHS As Double)
    This.RecordID = RHS
    RaiseEvent RecordIDChanged
End Property

Public Function IsValid() As Boolean
    If AutofillType = vbNullString Then Exit Function
    If AutofillData = vbNullString Then Exit Function
    IsValid = True
End Function

Private Sub Class_Initialize()
    This.AutofillType = -1
End Sub
