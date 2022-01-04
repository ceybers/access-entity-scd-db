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
    Tracks As Tracks
End Type

Private this As TCommit

' Unique Properties
Public Property Get IsClosed() As Boolean
    IsClosed = this.IsClosed
End Property

Public Property Let IsClosed(RHS As Boolean)
    this.IsClosed = RHS
End Property

Public Property Get Tracks() As Tracks
    Set Tracks = this.Tracks
End Property

Public Property Set Tracks(RHS As Tracks)
    Set this.Tracks = RHS
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
    Dim result As Commit
    Set result = New Commit
    With result
        .ID = rs.Fields("ID")
        .Name = rs.Fields(COMMIT_FIELDNAME)
        .IsClosed = IsNull(rs.Fields(COMMIT_CLOSED_FIELDNAME))
    End With
    Set IRecord_Create = result
End Function


Public Function ToString() As String
    ToString = this.ID & "# " & this.Name
    ToString = ToString & ", IsClosed = " & Iif(IsClosed, "Yes", "No")
    If Not this.Tracks Is Nothing Then
        ToString = ToString & ", Tracks = " & this.Tracks.Count
    End If
End Function
