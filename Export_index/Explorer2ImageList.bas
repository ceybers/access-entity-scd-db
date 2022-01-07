Attribute VB_Name = "Explorer2ImageList"
'@Folder("Explorer2")
Option Compare Database
Option Explicit

Public Sub InitializeImageList(ByRef il As ImageList, ByVal Width As Long, ByVal height As Long)
    modLoadImageList.Clear il
        
    il.ImageWidth = Width
    il.ImageHeight = height
    
    If il.ListImages.count = 0 Then
        modLoadImageList.Load il
    End If
End Sub
