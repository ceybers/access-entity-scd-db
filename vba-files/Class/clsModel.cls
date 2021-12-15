VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Event DetailChanged()
Public Event EntityChanged()
Public Event EntityTypeChanged()

Private mSelectedDetail As clsDetail
Private mSelectedEntity As clsEntity
Private mSelectedEntityType As clsEntityType

Private details As Collection
Private entities As Collection

Public Property Get Detail() As clsDetail
    Set Detail = mSelectedDetail
End Property

Public Property Let Detail(ByRef vNewValue As clsDetail)
    Set mSelectedDetail = vNewValue
    RaiseEvent DetailChanged
End Property

Public Property Get Entity() As clsEntity
    Set Entity = mSelectedEntity
End Property

Public Property Let Entity(ByRef vNewValue As clsEntity)
    Set mSelectedEntity = vNewValue
    RaiseEvent EntityChanged
End Property

Public Property Get EntityType() As clsEntityType
    Set EntityType = mSelectedEntityType
End Property

Public Property Let EntityType(ByRef vNewValue As clsEntityType)
    Set mSelectedEntityType = vNewValue
    RaiseEvent EntityTypeChanged
End Property

Private Sub Class_Initialize()
    LoadData
End Sub

Public Sub LoadData()
    LoadDetails
    LoadEntities
End Sub

Private Function LoadEntities()
    Set entities = New Collection
    
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblEntities")
    
    If Not rs.BOF And Not rs.EOF Then
        Do Until rs.EOF = True
            entities.Add clsEntity.Create(rs!ID, rs!Entity)
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
    Debug.Print "LoadEntities() complete"
End Function

Private Function LoadDetails()
    Set details = New Collection
    
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblDetailTables")
    
    If Not rs.BOF And Not rs.EOF Then
        Do Until rs.EOF = True
            details.Add clsDetail.Create(rs!ID, rs!DetailTable, rs!tableName)
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
    Debug.Print "LoadDetails() complete"
End Function

Public Function SetEntityByID(idx As Double)
    Dim thisEntity As clsEntity
    For Each thisEntity In entities
        If thisEntity.ID = idx Then
            Entity = thisEntity
            Exit Function
        End If
    Next thisEntity
End Function

Public Function SetDetailByID(idx As Double)
    Dim thisDetail As clsDetail
    For Each thisDetail In details
        If thisDetail.ID = idx Then
            Detail = thisDetail
            Exit Function
        End If
    Next thisDetail
End Function
