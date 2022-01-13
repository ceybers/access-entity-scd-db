Attribute VB_Name = "modBuildFormsForDetails"
'@Folder "Provisioning"
Option Compare Database
Option Explicit

Private Const TOP_MARGIN As Long = 31
Private Const SPECIAL_EFFECT As Long = 2

Private Type TControlSet
    FieldName As String
    Caption As String
    Width As String
    LookupTable As String
    Suffix As String
    Format As String
    Textalign As String
End Type

Public Sub BuildFormForDetail(ByVal detailName As String)
    Dim controlSets As collection
    Dim tableName As String
    Dim formName As String
    Dim frm As Form
    
    tableName = "tblDetail" & detailName
    formName = "sfrmDetail" & detailName
    
    CloseFormInDesignMode formName
    DeleteExistingForm formName
    CreateBlankForm (formName)
    OpenFormInDesignMode formName
    Set frm = Forms(formName)
    'RemoveAllControlsFromForm frm
    
    InitializeFormProperties frm
    
    Dim sql As String
    sql = "SELECT * FROM " & SCHEMA_TABLE & " WHERE TableName = '" & tableName & "';"
    Set controlSets = modRecordsetHelpers.RecordsetToCollection(sql, SubformControlSetRSCollector)
    AddSCDFields controlSets
    DrawFields formName, controlSets

    SetSCDFields formName
    CloseFormInDesignMode formName
    HideForm formName
End Sub

Private Sub InitializeFormProperties(ByRef frm As Form)
    With frm
        .NavigationButtons = False
        .RecordSelectors = False
        .ScrollBars = 0 ' Neither Horiz nor Vert
        '.dataentry = False
        .AllowAdditions = True
        .AllowEdits = True
        .AllowDeletions = False
        .recordSource = GetSQL(Replace(frm.name, "sfrm", "tbl"))
    End With
End Sub

Private Function GetSQL(ByVal tableName As String) As String
     GetSQL = "SELECT * FROM ((" & tableName & " AS tblDetail LEFT JOIN " _
        & ENTITIES_TABLE & " ON tblDetail.EntityFK = " & ENTITIES_TABLE & ".ID) LEFT JOIN " _
        & TRACKS_TABLE & " ON tblDetail.TrackFK = " & TRACKS_TABLE & ".ID) LEFT JOIN " & _
        COMMITS_TABLE & " ON " & TRACKS_TABLE & "." & TRACK_COMMITFK_FIELDNAME & " = " & _
        COMMITS_TABLE & ".ID ORDER BY metaTrack.ValidUntil DESC;"
End Function

Private Sub AddSCDFields(ByRef coll As collection)
    coll.Add subformcontrolset.Create("TrackFK", "Track ID", 2)
    coll.Add subformcontrolset.Create(TRACK_VALIDFROM_FIELDNAME, "Valid From", 4)
    coll.Add subformcontrolset.Create(TRACK_VALIDUNTIL_FIELDNAME, "Valid Until", 4)
    coll.Add subformcontrolset.Create(TRACK_COMMITFK_FIELDNAME, "Commit ID", 2)
End Sub

Private Sub DrawFields(ByVal formName As String, ByRef coll As collection)
    Dim cs As subformcontrolset
    Dim x As Long
    Dim i As Long
    
    For Each cs In coll
        i = i + 1
        x = ((DEFAULT_HEIGHT + 60) * (i - 1)) + 120
        
        If cs.RecordsetQuery = vbNullString Then
            CreateLabel formName, "lbl", cs, (0.25 * CM_TO_TWIP), x
            CreateTextBoxSCD formName, vbNullString, cs, (3.5 * CM_TO_TWIP), x
        ElseIf cs.LookupTable = vbNullString Then
            CreateTextBox formName, "txtLHS", cs, (3.5 * CM_TO_TWIP), x
            CreateTextBox formName, "txtRHS", cs, (7.75 * CM_TO_TWIP), x
            CreateLabel formName, "lblSuffix", cs, (7.75 * CM_TO_TWIP), x
        Else
            CreateComboBox formName, "cmbLHS", cs, (3.5 * CM_TO_TWIP), x
            CreateComboBox formName, "cmbRHS", cs, (7.75 * CM_TO_TWIP), x
        End If
            
    Next cs
End Sub

Private Sub CreateLabel(ByVal formName As String, ByVal prefix As String, ByRef cs As subformcontrolset, ByVal left As Long, ByVal top As Long)
    Dim lbl As Label
    Set lbl = CreateControl(formName:=formName, ControlType:=acLabel, left:=left, top:=top, Width:=(3 * CM_TO_TWIP), height:=DEFAULT_HEIGHT)
    
    With lbl
        .name = prefix & cs.FieldName
        .Caption = cs.Caption
        .TopMargin = TOP_MARGIN
    End With
End Sub

Private Sub CreateTextBoxSCD(ByVal formName As String, ByVal prefix As String, ByRef cs As subformcontrolset, ByVal left As Long, ByVal top As Long)
    Dim tb As textbox
    Set tb = CreateControl(formName:=formName, ControlType:=acTextBox, left:=left, top:=top, Width:=(4 * CM_TO_TWIP), height:=DEFAULT_HEIGHT)
    
    With tb
        .name = prefix & cs.FieldName
        .SpecialEffect = SPECIAL_EFFECT ' Sunken
        .TopMargin = TOP_MARGIN
        .ControlSource = cs.FieldName
        .Textalign = 1 'Left
        .ColumnWidth = 4 * CM_TO_TWIP
        .BackColor = RAGColors.grey
        .ColumnHidden = True
    End With
End Sub

Private Sub CreateTextBox(ByVal formName As String, ByVal prefix As String, ByRef cs As subformcontrolset, ByVal left As Long, ByVal top As Long)
    Dim tb As textbox
    Set tb = CreateControl(formName:=formName, ControlType:=acTextBox, left:=left, top:=top, Width:=(CDbl(cs.Width) * CM_TO_TWIP), height:=DEFAULT_HEIGHT)
    
    tb.name = prefix & cs.FieldName
    tb.SpecialEffect = SPECIAL_EFFECT
    tb.TopMargin = TOP_MARGIN

    If cs.Textalign <> vbNullString Then
        tb.Textalign = cs.Textalign
    End If
    
    If cs.Format <> vbNullString Then
        tb.Format = cs.Format
    End If
    
    If prefix = "txtLHS" Then
        tb.ControlSource = cs.FieldName
        With CreateControl(formName, acLabel, acDetail, tb.name, , 0.25 * CM_TO_TWIP, top, 3 * CM_TO_TWIP, DEFAULT_HEIGHT)
            .name = "lbl" & cs.FieldName
            .Caption = cs.Caption
        End With
    Else
        tb.ColumnHidden = True
    End If
End Sub

Private Sub CreateComboBox(ByVal formName As String, ByVal prefix As String, ByRef cs As subformcontrolset, ByVal left As Long, ByVal top As Long)
    Dim cb As ComboBox
    Set cb = CreateControl(formName:=formName, ControlType:=acComboBox, left:=left, top:=top, Width:=(4 * CM_TO_TWIP), height:=DEFAULT_HEIGHT)
    With cb
        .name = prefix & cs.FieldName
        .SpecialEffect = SPECIAL_EFFECT
        .TopMargin = TOP_MARGIN
        '.ControlSource = cs.FieldName
        .RowSource = cs.LookupTable
        .ColumnWidths = "0;" & (4 * CM_TO_TWIP) '2835" '2835 = 5cm
        .ColumnCount = 2
    End With
    
    If prefix = "cmbLHS" Then
        cb.ControlSource = cs.FieldName
        With CreateControl(formName, acLabel, acDetail, cb.name, , 0.25 * CM_TO_TWIP, top, 3 * CM_TO_TWIP, DEFAULT_HEIGHT)
            .name = "lbl" & cs.FieldName
            .Caption = cs.Caption
        End With
    Else
        cb.ColumnHidden = True
    End If
End Sub

Private Sub SetSCDFields(ByVal formName As String)
    Dim frm As Form
    Set frm = Forms(formName)
    
    frm.controls("lblTrackFK").Visible = False
    frm.controls("TrackFK").Visible = False
    frm.controls("ValidFrom").ColumnHidden = False
End Sub


