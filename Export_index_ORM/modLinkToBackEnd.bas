Attribute VB_Name = "modLinkToBackEnd"
'@Folder "Main"
Option Compare Database
Option Explicit

Public Sub ZZZ_LinkToBackEnd()
    LinkTable ENTITYTYPES_TABLE
    LinkTable ENTITIES_TABLE
End Sub
