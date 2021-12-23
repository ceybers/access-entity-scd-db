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

Private Const COMMITS_TO_DISPLAY As Integer = 5

Public Event CommitChanged()

Private Type TCommitViewModel
    AutofillType As String
    AutofillData As String
    RecordID As Double
    Commit As clsCommit
    Model As clsDataModel
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

Public Property Get Commit() As clsCommit
    Set Commit = This.Commit
End Property

Public Property Set Commit(ByRef RHS As clsCommit)
    Set This.Commit = RHS
    RaiseEvent CommitChanged
End Property

Public Function IsValid() As Boolean
    If Commit Is Nothing Then Exit Function
    IsValid = True
End Function

Public Function CanAutofill() As Boolean
    If AutofillType = vbNullString Then Exit Function
    If AutofillData = vbNullString Then Exit Function
    CanAutofill = True
End Function

Private Sub Class_Initialize()
    Set This.Model = New clsDataModel
    This.AutofillType = -1
End Sub

Public Function SetCommitByID(idx As Double)
    Dim aCommit As clsCommit
    Set This.Model = New clsDataModel
    Set aCommit = This.Model.GetThingByID(This.Model.Commits, idx)
    Set Commit = aCommit
End Function

Public Function ApplyCommitsSourceToCombobox(ByRef cmb As ComboBox)
    cmb.RowSource = ""
    cmb.RowSource = "SELECT TOP " & COMMITS_TO_DISPLAY & " ID, Title, Modified FROM tblCommits WHERE Closed Is Null ORDER BY Created DESC;"
End Function
