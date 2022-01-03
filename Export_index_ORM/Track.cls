VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Track"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "ORM.Tracks"
Option Compare Database
Option Explicit

Implements IRecord

' Variables
Private Type TTrack
    ID As Double
    Name As String
    ValidFrom As Date
    ValidUntil As Date
    Commit As Commit
End Type

Private this As TTrack

' Unique Properties
Public Property Get ValidFrom() As Date
    ValidFrom = this.ValidFrom
End Property

Public Property Let ValidFrom(RHS As Date)
    this.ValidFrom = RHS
End Property

Public Property Get ValidUntil() As Date
    ValidUntil = this.ValidUntil
End Property

Public Property Let ValidUntil(RHS As Date)
    this.ValidUntil = RHS
End Property

Public Property Get Commit() As Commit
    Set Commit = this.Commit
End Property

Public Property Set Commit(RHS As Commit)
    Set this.Commit = RHS
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
    Dim result As Track
    Set result = New Track
    
    With result
        .ID = rs.Fields("ID")
        .Name = rs.Fields(TRACK_FIELDNAME)
        .ValidFrom = rs.Fields(TRACK_VALIDFROM_FIELDNAME)
        .ValidUntil = rs.Fields(TRACK_VALIDUNTIL_FIELDNAME)
        Set .Commit = New Commit
        .Commit.ID = CDbl(rs.Fields("CommitFK")) * -1
    End With
    
    Set IRecord_Create = result
End Function

Public Function ToString() As String
    ToString = this.ID & "# "
    ToString = ToString & ", Valid<" & Format$(this.ValidFrom, "yyyy/mm/dd") & ";" & Format$(this.ValidUntil, "yyyy/mm/dd") & ">"
    If Not this.Commit Is Nothing Then
        ToString = ToString & ", Commit = " & this.Commit.ToString
    End If
End Function
