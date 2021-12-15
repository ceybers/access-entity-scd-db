Attribute VB_Name = "modPopulateFormControls"
Option Compare Database
Option Explicit

Const FILENAME As String = "C:\Users\User\Documents\xvba-access-test\schema.csv"
Const FORM_NAME As String = "sfrmTestTable"
Const TABLE_NAME As String = "tblTestTable"
Const CM_TO_TWIP As Integer = 567
Const DEFAULT_HEIGHT As Integer = 360

Public Sub PopulateFormControls()
    'frm.RecordSource = "TableName"
    Dim frm As Form
    DoCmd.OpenForm formName:=FORM_NAME, View:=acDesign
    Set frm = Forms(FORM_NAME)
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.ScrollBars = 0 ' Neither
    
    frm.DataEntry = False
    ' The Data Entry property doesn't determine whether records can be added; it only determines whether existing records are displayed.
    ' DataEntry = True & AllowAdditions = True --> Form is only for adding new records. You cannot scroll through existing ones.
    
    frm.AllowAdditions = True
    frm.AllowEdits = True
    frm.AllowDeletions = False
    'DoCmd.Close acForm, FORM_NAME, acSaveYes
    'DoCmd.OpenForm formName:=FORM_NAME
    Exit Sub
    
    Call RemoveAllControls(FORM_NAME)
    DoCmd.OpenForm formName:=FORM_NAME, View:=acDesign
    Call PopulateFormWithControlSets(FORM_NAME, LoadControlSets())
    DoCmd.Close acForm, FORM_NAME, acSaveYes
End Sub

Private Function LoadControlSets() As Collection
    Dim controlSetCollection As Collection
    Dim controlSet As sfrmControlSet
    Dim TextLine As String
    Dim arr As Variant
    Dim idx As Integer
    
    Set controlSetCollection = New Collection
    
    Open FILENAME For Input As #1
    
    ' Check CSV schema
    Line Input #1, TextLine
    arr = split(TextLine, ",")
    Debug.Assert arr(0) = "table"
    Debug.Assert arr(1) = "field"
    Debug.Assert arr(2) = "caption"
    Debug.Assert arr(3) = "lookup"
    Debug.Assert arr(4) = "suffix"
    Debug.Assert arr(5) = "default"
    
    ' Load controlSets
    Do While Not Eof(1)
        Line Input #1, TextLine
        arr = split(TextLine, ",")
        If CStr(arr(0)) = TABLE_NAME Then
            Set controlSet = New sfrmControlSet
            With controlSet
                .index = idx
                .fieldName = CStr(arr(1))
                .caption = CStr(arr(2))
                .lookup = CStr(arr(3))
                .suffix = CStr(arr(4))
                .defaultValue = CStr(arr(5))
            End With
            idx = idx + 1
            controlSetCollection.Add controlSet
        End If
    Loop
    Close #1
    
    Set LoadControlSets = controlSetCollection
End Function

Private Function PopulateFormWithControlSets(formName As String, controlSetCollection As Collection)
    Dim controlSet As sfrmControlSet
    Dim x As Integer
    
    For Each controlSet In controlSetCollection
        x = ((DEFAULT_HEIGHT + 60) * controlSet.index) + 120
        CreateLabel FORM_NAME, "lbl" & controlSet.fieldName, controlSet.caption, (0.25 * CM_TO_TWIP), x
        If controlSet.lookup <> "" Then
            CreateComboBox FORM_NAME, "txt" & controlSet.fieldName, controlSet.fieldName, controlSet.lookup, (3.25 * CM_TO_TWIP), x
        Else
            CreateTextBox FORM_NAME, "txt" & controlSet.fieldName, controlSet.fieldName, (3.25 * CM_TO_TWIP), x
        End If
        
        ' Add second textbox for before/after X/Y
        CreateLabel FORM_NAME, "lblSuffix" & controlSet.fieldName, controlSet.suffix, (7.5 * CM_TO_TWIP), x
    Next controlSet
End Function

Private Function RemoveAllControls(formName As String)
    Dim controlCount As Integer
    Dim frm As Form
    Dim i As Integer
    
    DoCmd.OpenForm formName:=formName, View:=acDesign
    
    Set frm = Forms(formName)
    For i = frm.Controls.Count To 1 Step -1
        DeleteControl formName, frm.Controls(i - 1).Name
    Next i
    
    DoCmd.Close acForm, formName, acSaveYes
End Function

Private Function CreateLabel(formName As String, controlName As String, caption As String, left As Integer, top As Integer)
    Dim lbl As Label
    Set lbl = CreateControl(formName:=formName, ControlType:=acLabel, left:=left, top:=top, Width:=(3 * CM_TO_TWIP), Height:=DEFAULT_HEIGHT)
    lbl.Name = controlName
    lbl.caption = caption
    lbl.TopMargin = 31
End Function

Private Function CreateTextBox(formName As String, controlName As String, fieldName As String, left As Integer, top As Integer)
    Dim tb As TextBox
    Set tb = CreateControl(formName:=formName, ControlType:=acTextBox, left:=left, top:=top, Width:=(4 * CM_TO_TWIP), Height:=DEFAULT_HEIGHT)
    tb.Name = controlName
    tb.SpecialEffect = 2
    tb.TopMargin = 31
    tb.ControlSource = fieldName
End Function

Private Function CreateComboBox(formName As String, controlName As String, fieldName As String, lookup As String, left As Integer, top As Integer)
    Dim cb As ComboBox
    Set cb = CreateControl(formName:=formName, ControlType:=acComboBox, left:=left, top:=top, Width:=(4 * CM_TO_TWIP), Height:=DEFAULT_HEIGHT)
    cb.Name = controlName
    cb.SpecialEffect = 2
    cb.TopMargin = 31
    cb.ControlSource = fieldName
    cb.RowSource = lookup
    cb.ColumnWidths = "0;2835" '2835 = 5cm
    cb.ColumnCount = 2
End Function

Private Function TEST_QueryControl()
    Dim frm As Form
    DoCmd.OpenForm formName:=FORM_NAME, View:=acDesign
    Set frm = Forms(FORM_NAME)
    Dim i As Integer
    Dim ctl As Control
    Dim tb As TextBox
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
