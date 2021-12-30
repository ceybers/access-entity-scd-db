VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFieldPair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Common"
Option Compare Database
Option Explicit

Private Type TClsFieldPair
    Source As String
    Destination As String
    Lookup As String
End Type

Private this As TClsFieldPair

Public Property Get Source() As String
    Source = this.Source
End Property

Public Property Let Source(ByVal RHS As String)
    this.Source = RHS
End Property

Public Property Get Destination() As String
    Destination = this.Destination
End Property

Public Property Let Destination(ByVal RHS As String)
    this.Destination = RHS
End Property

Public Property Get Lookup() As String
    Lookup = this.Lookup
End Property

Public Property Let Lookup(ByVal RHS As String)
    this.Lookup = RHS
End Property

Public Function Self() As clsFieldPair
    Set Self = Me
End Function
