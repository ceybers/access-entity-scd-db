Attribute VB_Name = "TestSCDFSM"
'@Folder "SCDFSM"
Option Compare Database
Option Explicit

Public Sub Test()
    Dim vm As SCDFSMViewModel
    Dim view As IView
    
    Set vm = New SCDFSMViewModel
    Set view = Form_SCDFSMView
    
    view.ShowDialog vm
    
    'If view.ShowDialog(vm) Then
        'Debug.Print "ShowDialog true"
    'Else
        'Debug.Print "ShowDialog false"
    'End If
End Sub
