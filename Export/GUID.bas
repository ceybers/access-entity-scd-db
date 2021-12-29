Attribute VB_Name = "GUID"
Option Compare Database
Option Explicit

' https://stackoverflow.com/a/46474125
' Erik A
' 28 September 2017

Public Function CreateGUID() As String
    Do While Len(CreateGUID) < 32
        If Len(CreateGUID) = 16 Then
            '17th character holds version information
            CreateGUID = CreateGUID & Hex$(8 + CInt(Rnd * 3))
        End If
        CreateGUID = CreateGUID & Hex$(CInt(Rnd * 15))
    Loop
    CreateGUID = "{" & Mid$(CreateGUID, 1, 8) & "-" & Mid$(CreateGUID, 9, 4) & "-" & Mid$(CreateGUID, 13, 4) & "-" & Mid$(CreateGUID, 17, 4) & "-" & Mid$(CreateGUID, 21, 12) & "}"
End Function
