Attribute VB_Name = "modPerspectives"
'@Folder("index_PQ")
Option Compare Database
Option Explicit

Public Sub BuildPerspectives()
    BuildPerspective "Default"
End Sub

Private Function DoesPerspectiveExist(ByVal perspectiveName As String) As Boolean
    Dim rs As Recordset
    Dim sql As String
    sql = "SELECT * FROM tblPerspective WHERE PerspectiveName = '" & perspectiveName & "';"
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
    DoesPerspectiveExist = (rs.RecordCount > 0)
    rs.Close
    Set rs = Nothing
End Function

Private Sub BuildPerspective(ByVal perspectiveName As String)
    Dim sql As String
    Dim queryName As String
    
    Debug.Assert DoesPerspectiveExist(perspectiveName)
    
    queryName = "per" & perspectiveName
    sql = GenerateSQLforPerspective(perspectiveName)
    Call modDBHelpers.CreateQuery(queryName, sql)
    'DoCmd.OpenQuery queryName, acViewNormal, acReadOnly
End Sub

Private Function GenerateSQLforPerspective(ByVal perspectiveName As String) As String
    Dim sql As String
    Dim fields As String
    Dim joins As String
    ' Dim entityTable As String
    
    ' TODO FIX
    ' entityTable = "qryTanks"
    
    fields = GetFields(perspectiveName)
    ' TODO Const
    fields = "tblEntitiesDepots.Entity,tblEntitiesTanks.Entity," & fields
    
    joins = GetJoins(perspectiveName)
    
    sql = "SELECT " & fields & " FROM " & joins
    
    GenerateSQLforPerspective = sql
End Function

Private Function GetJoins(ByVal perspectiveName As String) As String
    Dim db As Database
    Dim rs As Recordset
    Dim sql As String
    Dim joins As String
    
    sql = "SELECT DISTINCT EntityQuery, DetailQuery FROM tblPerspective WHERE PerspectiveName = '" & perspectiveName & "';"
    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
    
    'TODO FIX
    joins = "qryTanks"
    
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            joins = JoinTable(joins, rs.fields("EntityQuery"), rs.fields("DetailQuery"))
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    GetJoins = joins
End Function

Private Function JoinTable(ByVal previous As String, ByVal LHS As String, ByVal RHS As String) As String
    JoinTable = "(" & previous & " LEFT JOIN " & RHS & " ON " & LHS & ".ID = " & RHS & ".EntityFK)"
End Function

Private Function GetFields(ByVal perspectiveName As String) As String
    Dim db As Database
    Dim rs As Recordset
    Dim sql As String
    Dim fields As String
    
    sql = "SELECT * FROM tblPerspective WHERE PerspectiveName = '" & perspectiveName & "' ORDER BY [Order];"
    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
    
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            fields = fields & CStr(rs.fields("DetailQuery") & "." & rs.fields("FieldName")) & ","
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    If Len(fields) = 0 Then Exit Function
    GetFields = Left(fields, Len(fields) - 1)
End Function
