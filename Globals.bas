Attribute VB_Name = "Globals"
Option Explicit

Public gBooks As DOMDocument30 ' index for the saved books
Public gLoadedIndex As Long  ' which book is currently loaded
Public gLoadingID As String

Public Function GetUniqueKey() As String
'Generates a unique key for naming nodes in the tvwDirectory.
    On Error Resume Next
    Static keyCount As Long
    keyCount = keyCount + 1
    GetUniqueKey = "uk-" & keyCount
End Function
