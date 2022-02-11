VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Interface
'@Folder "Abstractions"
Option Compare Database
Option Explicit

'@Description "Displays the View modally and returns a value indicating confirmation (True) or cancellation (False)."
Public Function ShowDialog(ByVal ViewModel As IViewModel) As Boolean
Attribute ShowDialog.VB_Description = "Displays the View modally and returns a value indicating confirmation (True) or cancellation (False)."

End Function
