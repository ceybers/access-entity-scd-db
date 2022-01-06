VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ViewModelExplorer2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("index")
Option Compare Database
Option Explicit

Private Type TViewModelExplorer2
    EntityID As Double
    entityTypeid As Double
    DetailTableName As String
    SearchCriteria As String
End Type

Private this As TViewModelExplorer2

Public Property Get EntityID() As Double
    EntityID = this.EntityID
End Property

Public Property Let EntityID(ByVal RHS As Double)
    this.EntityID = RHS
End Property

Public Property Get entityTypeid() As Double
    entityTypeid = this.entityTypeid
End Property

Public Property Let entityTypeid(ByVal RHS As Double)
    this.entityTypeid = RHS
End Property

Public Property Get DetailSubFormName() As String
    DetailSubFormName = Replace(this.DetailTableName, "tbl", "sfrm")
End Property

Public Property Get DetailTableName() As String
    DetailTableName = this.DetailTableName
End Property

Public Property Let DetailTableName(ByVal RHS As String)
    this.DetailTableName = RHS
End Property

Public Property Get SearchCriteria() As String
    SearchCriteria = this.SearchCriteria
End Property

Public Property Let SearchCriteria(ByVal RHS As String)
    this.SearchCriteria = RHS
End Property

Public Function GetSQL() As String
    Debug.Assert this.entityTypeid > 0
    GetSQL = "SELECT * FROM metaEntities WHERE EntityType <= " & entityTypeid & " ORDER BY EntityType ASC, Entity ASC"
End Function

Public Function GetRecordset() As Recordset
    Set GetRecordset = CurrentDb.OpenRecordset(GetSQL, dbOpenSnapshot, dbReadOnly)
End Function

Private Sub Class_Initialize()
    ' TODO below is test env only. Handle unselcted/uninit entityTypeID properly
    this.entityTypeid = 1
End Sub
