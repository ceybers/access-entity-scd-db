Attribute VB_Name = "modCommitAutofill"
'@Folder "Explorer1"
Option Compare Database
Option Explicit

Private Enum CommitTypes
    Report = 2
    Drawing = 3
    Email = 6
End Enum

Public Function DoAutofill(Model As clsCommitViewModel, frm As Form)
    Select Case Model.AutofillType
        Case "Email"
            ProcessEmail Model, frm
        Case "Report"
            ProcessReport Model, frm
        Case "Drawing"
            ProcessDrawing Model, frm
    End Select
End Function

Private Function ProcessEmail(Model As clsCommitViewModel, frm As Form) As Boolean
    Dim s As Variant
    s = split(Model.AutofillData)
    If UBound(s) <> 3 Then Exit Function
    
    frm!Title = Model.AutofillData
    frm!RecvdFrom = CStr(s(0))
    frm!RecvdDate = CDate(s(2))
    frm!CommitType = CommitTypes.Email
    frm!Created = Now()
    
    ProcessEmail = True
End Function

Private Function ProcessReport(Model As clsCommitViewModel, frm As Form) As Boolean
    Dim repDate As Date
    Dim s As Variant
    s = split(Model.AutofillData, " - ")
    If UBound(s) <> 3 Then Exit Function
    
    repDate = CDate(s(3))
    
    frm!Title = CStr(s(1)) & " " & CStr(s(0)) & " " & format(repDate, "yyyymmdd")
    'frm!RecvdFrom = CStr(s(0))
    frm!RecvdDate = Now()
    frm!CommitType = CommitTypes.Report
    frm!DocReference = CStr(s(2))
    frm!DocDate = repDate
    frm!DocFilename = Model.AutofillData
    frm!Created = Now()
    
    ProcessReport = True
End Function

Private Function ProcessDrawing(Model As clsCommitViewModel, frm As Form) As Boolean
    Dim dwgDate As Date
    Dim s As Variant
    s = split(Model.AutofillData, " - ")
    If UBound(s) <> 3 Then Exit Function
    
    dwgDate = CDate(s(3))
    
    frm!Title = CStr(s(1)) & " " & CStr(s(0)) & " " & format(dwgDate, "yyyymmdd")
    'frm!RecvdFrom = CStr(s(0))
    frm!RecvdDate = Now()
    frm!CommitType = CommitTypes.Drawing
    frm!DocReference = CStr(s(2))
    frm!DocDate = dwgDate
    frm!DocFilename = Model.AutofillData
    frm!Created = Now()
    
    ProcessDrawing = True
End Function
