Attribute VB_Name = "Explorer2TreeView"
'@Folder("Explorer2")
Option Compare Database
Option Explicit

Public Sub ClearTreeview(ByRef treeview As treeview)
    Debug.Assert Not treeview Is Nothing

    If treeview.Nodes.count > 0 Then
        treeview.Nodes.Remove (1) 'Faster if we have one root node
        treeview.Nodes.Clear
    End If
End Sub

Public Function FindNodeInTreeview(ByRef tv As treeview, ByVal criteria As String, Optional exact As Boolean = False) As Node
    Dim localCriteria As String
    localCriteria = criteria
    
    Debug.Assert Not tv Is Nothing
    Debug.Assert Len(localCriteria) > 0
    
    Dim nde As Node
    
    If Not exact Then
        localCriteria = "*" & localCriteria & "*"
    End If
    
    For Each nde In tv.Nodes
        If nde.text Like localCriteria Then
            Set FindNodeInTreeview = nde
            Exit Function
        End If
    Next nde
End Function

Public Sub InitializeTreeview(ByRef treeview As treeview, ByRef ImageList As ImageList)
    Debug.Assert Not treeview Is Nothing
    
    With treeview
        .Appearance = ccFlat
        .Checkboxes = False
        .Indentation = 19
        .Style = tvwTreelinesPlusMinusPictureText
        .LabelEdit = tvwManual
        .LineStyle = tvwRootLines
        .HideSelection = False
        .SingleSel = False
        .FullRowSelect = True
        .ImageList = ImageList
    End With
End Sub

Public Sub PopulateTreeviewFromCollection(ByRef tv As treeview, ByRef coll As collection)
    Dim ent As Explorer2Entity
    Dim nde As Node
    
    For Each ent In coll
        If ent.Parent = "Entity#0" Then ' TODO Refactor
            Set nde = tv.Nodes.Add(, , ent.ID, ent.Entity)
        Else
            Set nde = tv.Nodes.Add(ent.Parent, tvwChild, ent.ID, ent.Entity)
        End If
        
        If ent.EntityType = "EntityType#4" Then
            nde.Image = "Kdocument"
        Else
            nde.Image = "Kdirectory_closed"
            nde.Expanded = True
        End If
    Next ent
End Sub
