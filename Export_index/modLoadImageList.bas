Attribute VB_Name = "modLoadImageList"
'@Folder "Explorer2"
Option Compare Database
Option Explicit

Private Const PATH As String = "C:\Users\User\Documents\access-entity-scd-db\Resources\Icons"

Public Sub Clear(ByRef il As ImageList)
    Debug.Assert Not il Is Nothing
    'Debug.Assert il Is ImageList
    
    il.ListImages.Clear
End Sub

Public Sub Load(ByRef il As ImageList)
    Debug.Assert Not il Is Nothing
    'Debug.Assert il Is ImageList
    
    Dim filesToLoad As String
    'filesToLoad = "ic_fluent_building_factory_24_filled.ico;ic_fluent_building_factory_24_regular.ico;ic_fluent_database_24_filled.ico;ic_fluent_database_24_regular.ico;ic_fluent_document_24_regular.ico;ic_fluent_folder_24_filled.ico;ic_fluent_folder_24_regular.ico;ic_fluent_folder_open_24_filled.ico;ic_fluent_folder_open_24_regular.ico;ic_fluent_organization_24_filled.ico;ic_fluent_organization_24_regular.ico"
    'filesToLoad = "4.ico;5.ico;fontview_111.ico"
    filesToLoad = "directory_open.ico;directory_closed.ico;document.ico;template_empty-0.ico"
    Dim fileArr As Variant
    fileArr = split(filesToLoad, ";")
    
    If il.ListImages.count > 0 Then
        MsgBox "Cannot load - Please clear first", vbCritical
        Exit Sub
    End If
    
    Dim i As Integer
    For i = 0 To UBound(fileArr)
        il.ListImages.Add key:="K" & Replace(fileArr(i), ".ico", ""), Picture:=LoadPicture(PATH & "\" & fileArr(i))
    Next i
    
    'Debug.Print UBound(fileArr) & " icon(s) loaded OK" ', vbInformation + vbOKOnly
End Sub
