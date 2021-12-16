Attribute VB_Name = "modTest"
Option Compare Database
Option Explicit

Public Function test()
    Dim lst As ListBox
    Set lst = Forms!frmTest.lstEntities
    Debug.Print lst.OnClick
    lst.OnClick = "=Meow()"
End Function

Public Function Meow()
    MsgBox "woeM"
End Function
