Attribute VB_Name = "Explorer2ListView"
'@Folder("Explorer2")
Option Compare Database
Option Explicit

Public Sub InitializeListView(ByRef listview As listview, ByRef largeIcons As ImageList, ByRef smallIcons As ImageList)
    Debug.Assert Not listview Is Nothing
    Debug.Assert Not largeIcons Is Nothing
    Debug.Assert Not smallIcons Is Nothing
    
    With listview
        .ColumnHeaders.Add , , "Details", 2880
        .FullRowSelect = True
        .GridLines = False
        .HideColumnHeaders = False
        .HideSelection = False
        .Appearance = ccFlat
        .LabelEdit = lvwManual
        .LabelWrap = False
        .BorderStyle = ccNone
        .Arrange = lvwAutoLeft
        
        '.View = lvwIcon
        '.View = lvwList
        .view = lvwReport
        '.View = lvwSmallIcon
        
        .Icons = largeIcons
        .smallIcons = smallIcons
        ' NOTE To get 32x32 icons with icon on left and text on right, use lvwList and set .SmallIcons to a 32x32 ImageList
    End With
End Sub

Public Sub PopulateListView(ByRef lv As MSComctlLib.listview, ByRef largeIcons As ImageList, ByRef smallIcons As ImageList)
    lv.ListItems.Clear
    lv.ColumnHeaders.Clear
    
    InitializeListView lv, largeIcons, smallIcons
    
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM metaDetailTables ORDER BY SortOrder ASC;", dbOpenSnapshot, dbReadOnly)
    
    Dim key As String
    Dim text As String
    
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
                key = rs.fields("tableName").Value ' TODO Const this
                text = rs.fields("DetailTable").Value ' TODO Const this
                lv.ListItems.Add , key, text, "Ktemplate_empty-0", "Ktemplate_empty-0" ' TODO Const this
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
End Sub
