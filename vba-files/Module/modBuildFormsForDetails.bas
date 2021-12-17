Attribute VB_Name = "modBuildFormsForDetails"
Option Compare Database
Option Explicit

' SELECT * FROM ((tblDetailDimensions AS tblDetail LEFT JOIN tblEntities ON tblDetail.EntityFK = tblEntities.ID) LEFT JOIN tblTrack ON tblDetail.TrackFK = tblTrack.ID) LEFT JOIN tblCommits ON tblTrack.CommitFK = tblCommits.ID;

Const FILENAME As String = "C:\Users\User\Documents\xvba-access-test\schema.csv"
'Const FORM_NAME As String = "sfrmTestTable"
'Const TABLE_NAME As String = "tblTestTable"
Dim FORM_NAME As String
Dim TABLE_NAME As String
Const CM_TO_TWIP As Integer = 567
Const DEFAULT_HEIGHT As Integer = 360

Private Type TControlSet
    FieldName As String
    Caption As String
    Width As String
    LookupTable As String
    Suffix As String
End Type

Public Sub BuildFormsForDetails()
    If MsgBox("Build forms?", vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    
    Call BuildFormForDetail("Dimensions")
    Call BuildFormForDetail("MaintPlan")
    Call BuildFormForDetail("Service")
End Sub

Private Function BuildFormForDetail(detailName As String)
    Dim TableName As String, formName As String
    TableName = "tblDetail" & detailName
    formName = "sfrmDetail" & detailName
    Dim controlSets() As TControlSet
    
    RemoveAllControls formName
    SetFormProperties (formName)
    controlSets = GetFields(TableName)
    Call DrawFields(formName, controlSets)
End Function

Private Function DrawFields(formName As String, fields() As TControlSet)
    Dim i As Integer
    Dim x As Integer
    Dim cs As TControlSet
    DoCmd.OpenForm formName:=formName, View:=acDesign
    
    For i = 1 To UBound(fields)
        cs = fields(i)
        x = ((DEFAULT_HEIGHT + 60) * (i - 1)) + 120
        CreateLabel formName, "lbl" & cs.FieldName, cs.Caption, (0.25 * CM_TO_TWIP), x
        CreateLabel formName, "lblSuffix" & cs.FieldName, cs.Suffix, (7.75 * CM_TO_TWIP), x
        
        If cs.LookupTable = "" Then
            CreateTextBox formName, "txtLHS" & cs.FieldName, cs.FieldName, (3.5 * CM_TO_TWIP), x
            CreateTextBox formName, "txtRHS" & cs.FieldName, "", (7.75 * CM_TO_TWIP), x
        Else
            CreateComboBox formName, "cmbLHS" & cs.FieldName, cs.FieldName, cs.LookupTable, (3.5 * CM_TO_TWIP), x
            CreateComboBox formName, "cmbRHS" & cs.FieldName, cs.FieldName, cs.LookupTable, (7.75 * CM_TO_TWIP), x
        End If
        'CreateLabel formName, "lblSuffix" & cs.FieldName, "", (12 * CM_TO_TWIP), x
        
    Next i
    DoCmd.Close acForm, formName, acSaveYes
End Function

Private Function GetFields(TableName As String) As TControlSet()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim results() As TControlSet
    Dim i As Integer
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM metaSchema WHERE TableName = '" & TableName & "';")
    i = 1
    
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            ReDim Preserve results(1 To i)
            results(i) = RecordToControlSet(rs)
            i = i + 1
            rs.MoveNext
        Loop
    End If
    
    Set rs = Nothing
    Set db = Nothing
    
    GetFields = results
End Function

Private Function RecordToControlSet(ByRef rs As Recordset) As TControlSet
    With RecordToControlSet
        .FieldName = rs!FieldName
        .Caption = rs!Caption
        .Width = rs!Caption
        .LookupTable = Nz(rs!LookupTable, "")
        .Suffix = Nz(rs!Suffix)
    End With
End Function

Private Function SetFormProperties(formName As String)
    Dim frm As Form
    DoCmd.OpenForm formName:=formName, View:=acDesign
    Set frm = Forms(formName)
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.ScrollBars = 0 ' Neither
    frm.DataEntry = False
    frm.AllowAdditions = True
    frm.AllowEdits = True
    frm.AllowDeletions = False
    frm.RecordSource = ""
    frm.RecordSource = Replace(formName, "sfrm", "tbl")
    DoCmd.Close acForm, formName, acSaveYes
End Function

Private Function RemoveAllControls(formName As String)
    Dim controlCount As Integer
    Dim frm As Form
    Dim i As Integer
    
    DoCmd.OpenForm formName:=formName, View:=acDesign
    
    Set frm = Forms(formName)
    For i = frm.Controls.Count To 1 Step -1
        DeleteControl formName, frm.Controls(i - 1).name
    Next i
    
    DoCmd.Close acForm, formName, acSaveYes
End Function

Private Function CreateLabel(formName As String, controlName As String, Caption As String, left As Integer, top As Integer)
    Dim lbl As Label
    Set lbl = CreateControl(formName:=formName, ControlType:=acLabel, left:=left, top:=top, Width:=(3 * CM_TO_TWIP), Height:=DEFAULT_HEIGHT)
    lbl.name = controlName
    lbl.Caption = Caption
    lbl.TopMargin = 31
End Function

Private Function CreateTextBox(formName As String, controlName As String, FieldName As String, left As Integer, top As Integer)
    Dim tb As textbox
    Set tb = CreateControl(formName:=formName, ControlType:=acTextBox, left:=left, top:=top, Width:=(4 * CM_TO_TWIP), Height:=DEFAULT_HEIGHT)
    tb.name = controlName
    tb.SpecialEffect = 2
    tb.TopMargin = 31
    tb.ControlSource = FieldName
    tb.TextAlign = 1 'Left
End Function

Private Function CreateComboBox(formName As String, controlName As String, FieldName As String, lookup As String, left As Integer, top As Integer)
    Dim cb As ComboBox
    Set cb = CreateControl(formName:=formName, ControlType:=acComboBox, left:=left, top:=top, Width:=(4 * CM_TO_TWIP), Height:=DEFAULT_HEIGHT)
    cb.name = controlName
    cb.SpecialEffect = 2
    cb.TopMargin = 31
    cb.ControlSource = FieldName
    cb.RowSource = lookup
    cb.ColumnWidths = "0;2835" '2835 = 5cm
    cb.ColumnCount = 2
End Function

Private Function TEST_QueryControl()
    Dim frm As Form
    DoCmd.OpenForm formName:=FORM_NAME, View:=acDesign
    Set frm = Forms(FORM_NAME)
    Dim i As Integer
    Dim ctl As control
    Dim tb As textbox
    For i = frm.Controls.Count To 1 Step -1
        Set ctl = frm.Controls(i - 1)
        If ctl.ControlType = acTextBox Then
            Set tb = ctl
            Debug.Print "Layout: " & ctl.Layout
            Debug.Print "Top Margin: " & tb.TopMargin '31
            Debug.Print "Special Effect: " & tb.SpecialEffect '2
            Debug.Print "Width: " & tb.Width '2268 = 4cm
            Debug.Print "Height: " & tb.Height ' 360 = 0.635cm
        End If
    Next i
End Function
