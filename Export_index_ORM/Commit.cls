VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Commit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "ORM.Commits"
Option Compare Database
Option Explicit

Implements IRecord

' Variables
Private Type TCommit
    ID As Double
    Name As String
    IsClosed As Boolean
    'Tracks As Tracks
End Type

Private this As TCommit

' Unique Properties
Public Property Get IsClosed() As Boolean
    IsClosed = this.IsClosed
End Property

' Common Properties
'@DefaultMember
Public Property Get Name() As String
Attribute Name.VB_UserMemId = 0
    Name = this.Name
End Property

Public Property Let Name(ByVal RHS As String)
    this.Name = RHS
End Property

Public Property Get ID() As Double
    ID = this.ID
End Property

Public Property Let ID(ByVal RHS As Double)
    this.ID = RHS
End Property

' Interface
Private Property Get IRecord_ID() As Double
    IRecord_ID = ID
End Property

Private Property Let IRecord_ID(ByVal RHS As Double)
    ID = RHS
End Property

Private Property Get IRecord_Name() As String
    IRecord_Name = Name
End Property

Private Property Let IRecord_Name(ByVal RHS As String)
    Name = RHS
End Property

' Methods
Private Function IRecord_Create(ByRef rs As Recordset) As IRecord
    Set IRecord_Create = New Commit
    With IRecord_Create
        .ID = rs.Fields("ID")
        .Name = rs.Fields(COMMIT_FIELDNAME)
        .IsClosed = IsNull(rs.Fields("Closed")) ' TODO Const this
    End With
End Function

Public Function ToString() As String
    ToString = this.ID & "# " & this.Name
    'If Not this.Parent Is Nothing Then
    '    ToString = ToString & ", Parent = " & this.Parent.Name
    'End If
End Function
