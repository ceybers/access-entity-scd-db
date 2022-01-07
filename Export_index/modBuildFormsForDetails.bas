Attribute VB_Name = "modBuildFormsForDetails"
'@Folder "Provisioning"
Option Compare Database
Option Explicit

Private Type TControlSet
    fieldName As String
    caption As String
    width As String
    lookupTable As String
    suffix As String
    format As String
    textalign As String
End Type

Public Function BuildFormForDetail(detailName As String)
    Dim tableName As String, formName As String
    tableName = "tblDetail" & detailName
    formName = "sfrmDetail" & detailName
    Dim controlSets() As TControlSet
    
    CloseFormInDesignMode formName
    DeleteExistingForm formName
    CreateBlankForm formName
    OpenFormInDesignMode formName
    RemoveAllControls formName
    SetFormProperties formName
    controlSets = GetFields(tableName)
    DrawFields formName, controlSets
    SetSCDFields formName
    CloseFormInDesignMode formName
    HideForm formName
End Function

Private Function CreateBlankForm(formName As String)
    Dim oldName As String
    Dim frm As Form
    Set frm = CreateForm()
    oldName = frm.name
    DoCmd.Close acForm, oldName, acSaveYes
    DoCmd.Rename formName, acForm, oldName
End Function

Private Function DeleteExistingForm(formName As String)
    Dim frm As Object
    For Each frm In CurrentProject.AllForms
        If frm.name = formName Then
            DoCmd.DeleteObject acForm, formName
            Exit Function
        End If
    Next frm
End Function

Private Function DrawFields(formName As String, fields() As TControlSet)
    Dim i As Long
    Dim x As Long
    Dim cs As TControlSet
    
    For i = 1 To UBound(fields)
        cs = fields(i)
        x = ((DEFAULT_HEIGHT + 60) * (i - 1)) + 120
        'CreateLabel formName, "lbl" & cs.fieldName, IIf(cs.caption = vbNullString, cs.fieldName, cs.caption), (0.25 * CM_TO_TWIP), x
        CreateLabel formName, "lblSuffix" & cs.fieldName, cs.suffix, (7.75 * CM_TO_TWIP), x
        
        ' TODO Refactor this exclusion list
        If cs.fieldName = TRACK_VALIDFROM_FIELDNAME Or cs.fieldName = TRACK_VALIDUNTIL_FIELDNAME Or cs.fieldName = "TrackFK" Or cs.fieldName = TRACK_COMMITFK_FIELDNAME Then
            CreateLabel formName, "lbl" & cs.fieldName, IIf(cs.caption = vbNullString, cs.fieldName, cs.caption), (0.25 * CM_TO_TWIP), x
            CreateTextBox formName, cs.fieldName, cs.fieldName, (3.5 * CM_TO_TWIP), x
        ElseIf cs.lookupTable = vbNullString Then
            CreateTextBox2 formName, "txtLHS", cs, (3.5 * CM_TO_TWIP), x
            CreateTextBox2 formName, "txtRHS", cs, (7.75 * CM_TO_TWIP), x
        Else
            CreateComboBox formName, "cmbLHS" & cs.fieldName, cs.fieldName, cs.lookupTable, (3.5 * CM_TO_TWIP), x
            CreateComboBox formName, "cmbRHS" & cs.fieldName, vbNullString, cs.lookupTable, (7.75 * CM_TO_TWIP), x
        End If
        'CreateLabel formName, "lblSuffix" & cs.FieldName, "", (12 * CM_TO_TWIP), x
        
    Next i
End Function

Private Function GetFields(tableName As String) As TControlSet()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim results() As TControlSet
    Dim i As Long
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM " & SCHEMA_TABLE & " WHERE TableName = '" & tableName & "';")
    i = 1
    
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            ReDim Preserve results(1 To i)
            results(i) = RecordToControlSet(rs)
            i = i + 1
            rs.MoveNext
        Loop
    End If
    
    If i = 1 Then
        Err.Raise 5, , "No entries in `metaSchema` table!"
    End If
    
    ' Add SCD common fields
    results = AppendToControlSet(results, CreateControlSet("TrackFK", "Track ID")) ' TODO Const this
    results = AppendToControlSet(results, CreateControlSet(TRACK_VALIDFROM_FIELDNAME, "Valid From"))
    results = AppendToControlSet(results, CreateControlSet(TRACK_VALIDUNTIL_FIELDNAME, "Valid Until"))
    results = AppendToControlSet(results, CreateControlSet(TRACK_COMMITFK_FIELDNAME, "Commit ID"))
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    GetFields = results
End Function

Private Function AppendToControlSet(ByRef coll() As TControlSet, tc As TControlSet) As TControlSet()
    Dim i As Long
    i = UBound(coll) + 1
    ReDim Preserve coll(1 To i)
    coll(i) = tc
    AppendToControlSet = coll
End Function

Private Function CreateControlSet(fieldName As String, caption As String) As TControlSet
    With CreateControlSet
        .fieldName = fieldName
        .caption = caption
        .width = 4
    End With
End Function

Private Function RecordToControlSet(ByRef rs As Recordset) As TControlSet
    With RecordToControlSet
        .fieldName = rs!fieldName
        .caption = Nz(rs!caption, vbNullString)
        .width = Nz(rs!width, vbNullString)
        .lookupTable = Nz(rs!lookupTable, vbNullString)
        .suffix = Nz(rs!suffix)
        .format = Nz(rs!format)
        .textalign = Nz(rs!textalign)
    End With
End Function

Private Function SetFormProperties(formName As String)
    Dim frm As Form
    Set frm = Forms(formName)
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.ScrollBars = 0 ' Neither
    frm.dataentry = False
    frm.AllowAdditions = True
    frm.AllowEdits = True
    frm.AllowDeletions = False
    frm.recordSource = GetSQL(Replace(formName, "sfrm", "tbl"))
End Function

Private Function RemoveAllControls(formName As String)
    Dim frm As Form
    Dim i As Long
    
    Set frm = Forms(formName)
    For i = frm.controls.count To 1 Step -1
        DeleteControl formName, frm.controls(i - 1).name
    Next i
    
End Function

Private Function CreateLabel(formName As String, controlName As String, caption As String, left As Long, top As Long)
    Dim lbl As Label
    Set lbl = CreateControl(formName:=formName, ControlType:=acLabel, left:=left, top:=top, width:=(3 * CM_TO_TWIP), height:=DEFAULT_HEIGHT)
    lbl.name = controlName
    lbl.caption = caption
    lbl.TopMargin = 31
End Function

Private Function CreateTextBox(formName As String, controlName As String, fieldName As String, left As Long, top As Long)
    Dim tb As textbox
    Set tb = CreateControl(formName:=formName, ControlType:=acTextBox, left:=left, top:=top, width:=(4 * CM_TO_TWIP), height:=DEFAULT_HEIGHT)
    tb.name = controlName
    tb.SpecialEffect = 2
    tb.TopMargin = 31
    tb.ControlSource = fieldName
    tb.textalign = 1 'Left
End Function

Private Function CreateTextBox2(formName As String, prefix As String, cs As TControlSet, left As Long, top As Long)
    Dim tb As textbox
    Set tb = CreateControl(formName:=formName, ControlType:=acTextBox, left:=left, top:=top, width:=(CDbl(cs.width) * CM_TO_TWIP), height:=DEFAULT_HEIGHT)
    tb.name = prefix & cs.fieldName
    tb.SpecialEffect = 2
    tb.TopMargin = 31
    If prefix = "txtLHS" Then
        tb.ControlSource = cs.fieldName
    End If
    If cs.textalign <> vbNullString Then
        tb.textalign = cs.textalign
    End If
    If cs.format <> vbNullString Then
        tb.format = cs.format
    End If
    If prefix = "txtLHS" Then
        With CreateControl(formName, acLabel, acDetail, tb.name, , 0.25 * CM_TO_TWIP, top, 3 * CM_TO_TWIP, DEFAULT_HEIGHT)
            .name = "lb" & cs.fieldName
            .caption = IIf(cs.caption = vbNullString, cs.fieldName, cs.caption)
        End With
    End If
End Function

Private Function CreateComboBox(formName As String, controlName As String, fieldName As String, lookup As String, left As Long, top As Long)
    Dim cb As ComboBox
    Set cb = CreateControl(formName:=formName, ControlType:=acComboBox, left:=left, top:=top, width:=(4 * CM_TO_TWIP), height:=DEFAULT_HEIGHT)
    cb.name = controlName
    cb.SpecialEffect = 2
    cb.TopMargin = 31
    cb.ControlSource = fieldName
    cb.RowSource = lookup
    cb.ColumnWidths = "0;2835" '2835 = 5cm
    cb.ColumnCount = 2
End Function

Private Function GetSQL(tableName As String)
     GetSQL = "SELECT * FROM ((" & tableName & " AS tblDetail LEFT JOIN " _
        & ENTITIES_TABLE & " ON tblDetail.EntityFK = " & ENTITIES_TABLE & ".ID) LEFT JOIN " _
        & TRACKS_TABLE & " ON tblDetail.TrackFK = " & TRACKS_TABLE & ".ID) LEFT JOIN " & _
        COMMITS_TABLE & " ON " & TRACKS_TABLE & "." & TRACK_COMMITFK_FIELDNAME & " = " & _
        COMMITS_TABLE & ".ID ORDER BY metaTrack.ValidUntil DESC;"
End Function

Private Sub SetSCDFields(formName As String)
    Dim frm As Form
    Set frm = Forms(formName)
    
    'frm!lblTrackFK.Visible = False
    'frm!TrackFK.Visible = False
End Sub

Private Sub HideForm(formName As String)
    Application.SetHiddenAttribute acForm, formName, True
End Sub


