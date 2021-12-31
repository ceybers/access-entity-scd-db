Attribute VB_Name = "modHideTables"
'@Folder("index_PQ")
Option Compare Database
Option Explicit

Public Sub HideTables()
    Dim tables As Variant
    Set tables = GetListOfTables
    
    DoHideTables tables, "tblDetail*"
    DoHideTables tables, "lkp*"
    DoHideTables tables, "metaSchema"
End Sub

Private Sub DoHideTables(ByRef tables As Variant, ByVal criteria As String)
    Dim table As Variant
    For Each table In tables
        If table Like criteria Then
            Application.SetHiddenAttribute acTable, table, True
        End If
    Next table
End Sub
