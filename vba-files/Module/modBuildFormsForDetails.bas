Attribute VB_Name = "modBuildFormsForDetails"
Option Compare Database
Option Explicit

Const FILENAME As String = "C:\Users\User\Documents\xvba-access-test\schema.csv"
Dim FORM_NAME As String
Dim TABLE_NAME As String
Const CM_TO_TWIP As Integer = 567
Const DEFAULT_HEIGHT As Integer = 360

Private Type TControlSet
    fieldName As String
    caption As String
    width As String
    lookupTable As String
    suffix As String
End Type

Public Sub BuildFormsForDetails()
    If MsgBox("Build forms?", vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    
    'Call BuildFormForDetail("Dimensions")
    'Call BuildFormForDetail("MaintPlan")
    'Call BuildFormForDetail("Service")
    
    Call BuildFormForDetail("A")
    Call BuildFormForDetail("B")
    Call BuildFormForDetail("C")
End Sub

Private Function BuildFormForDetail(detailName As String)
    Dim tableName As String, formName As String
    tableName = "tblDetail" & detailName
    formName = "sfrmDetail" & detailName
    Dim controlSets() As TControlSet
    
    RemoveAllControls formName
    SetFormProperties formName
    controlSets = GetFields(tableName)
    Call DrawFields(formName, controlSets)
    SetSCDFields formName
End Function

Private Function DrawFields(formName As String, fields() As TControlSet)
    Dim i As Integer
    Dim x As Integer
    Dim cs As TControlSet
    DoCmd.OpenForm formName:=formName, View:=acDesign
    
    For i = 1 To UBound(fields)
        cs = fields(i)
        x = ((DEFAULT_HEIGHT + 60) * (i - 1)) + 120
        CreateLabel formName, "lbl" & cs.fieldName, cs.caption, (0.25 * CM_TO_TWIP), x
        CreateLabel formName, "lblSuffix" & cs.fieldName, cs.suffix, (7.75 * CM_TO_TWIP), x
        
        If cs.fieldName = "ValidFrom" Or cs.fieldName = "TrackFK" Or cs.fieldName = "CommitFK" Then
            CreateTextBox formName, cs.fieldName, cs.fieldName, (3.5 * CM_TO_TWIP), x
        ElseIf cs.lookupTable = "" Then
            CreateTextBox formName, "txtLHS" & cs.fieldName, cs.fieldName, (3.5 * CM_TO_TWIP), x
            CreateTextBox formName, "txtRHS" & cs.fieldName, "", (7.75 * CM_TO_TWIP), x
        Else
            CreateComboBox formName, "cmbLHS" & cs.fieldName, cs.fieldName, cs.lookupTable, (3.5 * CM_TO_TWIP), x
            CreateComboBox formName, "cmbRHS" & cs.fieldName, "", cs.lookupTable, (7.75 * CM_TO_TWIP), x
        End If
        'CreateLabel formName, "lblSuffix" & cs.FieldName, "", (12 * CM_TO_TWIP), x
        
    Next i
    DoCmd.Close acForm, formName, acSaveYes
End Function

Private Function GetFields(tableName As String) As TControlSet()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim results() As TControlSet
    Dim i As Integer
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM metaSchema WHERE TableName = '" & tableName & "';")
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
    results = AppendToControlSet(results, CreateControlSet("TrackFK", "Track ID"))
    results = AppendToControlSet(results, CreateControlSet("ValidFrom", "Valid From"))
    results = AppendToControlSet(results, CreateControlSet("CommitFK", "Commit ID"))
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    GetFields = results
End Function

Private Function AppendToControlSet(ByRef coll() As TControlSet, tc As TControlSet) As TControlSet()
    Dim i As Integer
    i = UBound(coll) + 1
    ReDim Preserve coll(1 To i)
    coll(i) = tc
    AppendToControlSet = coll
End Function

Private Function CreateControlSet(fieldName As String, caption As String) As TControlSet
    With CreateControlSet
        .fieldName = fieldName
        .caption = caption
    End With
End Function

Private Function RecordToControlSet(ByRef rs As Recordset) As TControlSet
    With RecordToControlSet
        .fieldName = rs!fieldName
        .caption = rs!caption
        .width = rs!caption
        .lookupTable = Nz(rs!lookupTable, "")
        .suffix = Nz(rs!suffix)
    End With
End Function

Private Function SetFormProperties(formName As String)
    Dim frm As Form
    DoCmd.OpenForm formName:=formName, View:=acDesign
    Set frm = Forms(formName)
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.ScrollBars = 0 ' Neither
    frm.dataentry = False
    frm.AllowAdditions = True
    frm.AllowEdits = True
    frm.AllowDeletions = False
    frm.RecordSource = ""
    'frm.RecordSource = Replace(formName, "sfrm", "tbl")
    frm.RecordSource = GetSQL(Replace(formName, "sfrm", "tbl"))
    DoCmd.Close acForm, formName, acSaveYes
End Function

Private Function RemoveAllControls(formName As String)
    Dim controlCount As Integer
    Dim frm As Form
    Dim i As Integer
    
    DoCmd.OpenForm formName:=formName, View:=acDesign
    
    Set frm = Forms(formName)
    For i = frm.controls.Count To 1 Step -1
        DeleteControl formName, frm.controls(i - 1).name
    Next i
    
    DoCmd.Close acForm, formName, acSaveYes
End Function

Private Function CreateLabel(formName As String, controlName As String, caption As String, left As Integer, top As Integer)
    Dim lbl As Label
    Set lbl = CreateControl(formName:=formName, ControlType:=acLabel, left:=left, top:=top, width:=(3 * CM_TO_TWIP), Height:=DEFAULT_HEIGHT)
    lbl.name = controlName
    lbl.caption = caption
    lbl.TopMargin = 31
End Function

Private Function CreateTextBox(formName As String, controlName As String, fieldName As String, left As Integer, top As Integer)
    Dim tb As textbox
    Set tb = CreateControl(formName:=formName, ControlType:=acTextBox, left:=left, top:=top, width:=(4 * CM_TO_TWIP), Height:=DEFAULT_HEIGHT)
    tb.name = controlName
    tb.SpecialEffect = 2
    tb.TopMargin = 31
    tb.ControlSource = fieldName
    tb.TextAlign = 1 'Left
End Function

Private Function CreateComboBox(formName As String, controlName As String, fieldName As String, lookup As String, left As Integer, top As Integer)
    Dim cb As ComboBox
    Set cb = CreateControl(formName:=formName, ControlType:=acComboBox, left:=left, top:=top, width:=(4 * CM_TO_TWIP), Height:=DEFAULT_HEIGHT)
    cb.name = controlName
    cb.SpecialEffect = 2
    cb.TopMargin = 31
    cb.ControlSource = fieldName
    cb.RowSource = lookup
    cb.ColumnWidths = "0;2835" '2835 = 5cm
    cb.ColumnCount = 2
End Function

Private Function GetSQL(tableName As String)
     GetSQL = "SELECT * FROM ((" & tableName & " AS tblDetail LEFT JOIN tblEntities ON tblDetail.EntityFK = tblEntities.ID) LEFT JOIN tblTrack ON tblDetail.TrackFK = tblTrack.ID) LEFT JOIN tblCommits ON tblTrack.CommitFK = tblCommits.ID;"
End Function

Private Sub SetSCDFields(formName As String)
    Dim frm As Form
    DoCmd.OpenForm formName:=formName, View:=acDesign
    Set frm = Forms(formName)
    
    frm!lblTrackFK.Visible = False
    frm!TrackFK.Visible = False
    
    DoCmd.Close acForm, formName, acSaveYes
End Sub

Private Function TEST_QueryControl()
    Dim frm As Form
    DoCmd.OpenForm formName:=FORM_NAME, View:=acDesign
    Set frm = Forms(FORM_NAME)
    Dim i As Integer
    Dim ctl As control
    Dim tb As textbox
    For i = frm.controls.Count To 1 Step -1
        Set ctl = frm.controls(i - 1)
        If ctl.ControlType = acTextBox Then
            Set tb = ctl
            Debug.Print "Layout: " & ctl.Layout
            Debug.Print "Top Margin: " & tb.TopMargin '31
            Debug.Print "Special Effect: " & tb.SpecialEffect '2
            Debug.Print "Width: " & tb.width '2268 = 4cm
            Debug.Print "Height: " & tb.Height ' 360 = 0.635cm
        End If
    Next i
End Function

