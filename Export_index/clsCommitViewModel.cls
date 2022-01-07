VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommitViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule

'@Folder "ViewModels"
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
Private this As TCommitViewModel

Public Property Get AutofillType() As String
    AutofillType = this.AutofillType
End Property

Public Property Let AutofillType(ByVal RHS As String)
    this.AutofillType = RHS
End Property

Public Property Get AutofillData() As String
    AutofillData = this.AutofillData
End Property

Public Property Let AutofillData(ByVal RHS As String)
    this.AutofillData = RHS
End Property

Public Property Get Commit() As clsCommit
    Set Commit = this.Commit
End Property

Public Property Set Commit(ByRef RHS As clsCommit)
    Set this.Commit = RHS
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
    Set this.Model = New clsDataModel
    this.AutofillType = -1
End Sub

Public Function SetCommitByID(idx As Double)
    Dim aCommit As clsCommit
    Set this.Model = New clsDataModel
    Set aCommit = this.Model.GetThingByID(this.Model.Commits, idx)
    Set Commit = aCommit
End Function

Public Function ApplyCommitsSourceToCombobox(ByRef cmb As ComboBox)
    cmb.RowSource = vbNullString
    cmb.RowSource = "SELECT TOP " & COMMITS_TO_DISPLAY & " ID, Title, Modified FROM " & COMMITS_TABLE & " WHERE Closed Is Null ORDER BY Created DESC;"
End Function
