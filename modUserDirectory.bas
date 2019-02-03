Attribute VB_Name = "modUserDirectory"
Option Explicit

'SHGetSpecialFolderLocation
'Returns the Folder ID of the user's My Documents folder (or another folder indicated
'by CSIDL) See "Chapter 4: Data and Settings Management".
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hWnd As Long, ByVal nFolder As Long, ppidl As Long) As Long
Private Const CSIDL_PERSONAL = &H5
Private Const CSIDL_LOCAL_APPDATA = &H1C
Private Const CSIDL_APPDATA As Long = &H1A

'SHGetPathFromIDList
'Returns the path (string) from the folder ID obtained by SHGetSpecialFolderLocation
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal Pidl As Long, ByVal pszPath As String) As Long

Public Function GetNonRoamingApplicationDirectory() As String
'uses the Windows API to get the path for the application directory,
'Application Data
    On Error Resume Next
    Dim path As String
    Dim result As Long
    Dim referenceID As Long
    
    'work out where to save them to:
    path = Space(260)
    result = SHGetSpecialFolderLocation(0, CSIDL_LOCAL_APPDATA, referenceID)
    result = SHGetPathFromIDList(referenceID, path)
    'assertion: path now contains path to My Documents
    path = Trim(path)
    'take off final null character which trim has left behind
    path = Replace(path, Chr(0), "")
    'return
    GetNonRoamingApplicationDirectory = path
End Function

Public Function GetMyDocuments() As String
'uses the Windows API to get the path for the user's home directory,
'My Documents.
    On Error Resume Next
    Dim path As String
    Dim result As Long
    Dim referenceID As Long
    
    'work out where to save them to:
    path = Space(260)
    result = modUserDirectory.SHGetSpecialFolderLocation(0, CSIDL_PERSONAL, referenceID)
    result = modUserDirectory.SHGetPathFromIDList(referenceID, path)
    'assertion: path now contains path to My Documents
    path = Trim(path)
    'take off final null character which trim has left behind
    path = Replace(path, Chr(0), "")
    'return
    GetMyDocuments = path
End Function

Public Function GetTempDirectory() As String
'uses the Windows API to get the path for the current temp directory
    On Error GoTo localHandler:
    GetTempDirectory = GetNonRoamingApplicationDirectory
    'This is now ....Local Settings\Application Data, but we want
    '....Local Settings\Temp
    'I'm assuming that these aren't localised!
    GetTempDirectory = Replace(GetNonRoamingApplicationDirectory, "Application Data", "Temp")
    Exit Function
localHandler:
    'MsgBox Err.Number & " " & Err.Description & vbNewLine & Err.Source, vbOKOnly, "Error: GetTempDirectory"
    Resume Next
End Function

Public Function GetRoamingApplicationDirectory() As String
'Uses the Windows API to get the path for the user's roaming profile
On Error Resume Next
    Dim path As String
    Dim result As Long
    Dim referenceID As Long
    
    'work out where to save them to:
    path = Space(260)
    result = SHGetSpecialFolderLocation(0, CSIDL_APPDATA, referenceID)
    result = SHGetPathFromIDList(referenceID, path)
    'assertion: path now contains path to My Documents
    path = Trim(path)
    'take off final null character which trim has left behind
    path = Replace(path, Chr(0), "")
    'return
    GetRoamingApplicationDirectory = path
End Function

Public Function GetCurrentDate() As String
'returns current date in RFC822 specification
    On Error Resume Next
    Dim datetime As String
    
    datetime = WeekdayName(Weekday(Date), True) & ", "
    datetime = datetime & Day(Date) & " "
    datetime = datetime & MonthName(Month(Date), True) & " "
    datetime = datetime & Right(Year(Date), 2) & " "
    datetime = datetime & DatePart("h", Now) & ":" & DatePart("n", Now)
    
    GetCurrentDate = datetime
End Function

Public Function BuildNonRoamingApplicationPath(companyName As String, productName As String, version As String) As String
'returns a Windows-compliant non-roaming application path where you can store
'programme data
    On Error Resume Next
    Dim path As String
    
    path = GetNonRoamingApplicationDirectory
    path = path & "\" & companyName & "\" & productName & "\" & version
    BuildNonRoamingApplicationPath = path
End Function

Public Function BuildRoamingApplicationPath(companyName As String, productName As String, version As String) As String
'returns a Windows-compliant nroaming application path where you can store
'programme data
    On Error Resume Next
    Dim path As String
    
    path = GetRoamingApplicationDirectory
    path = path & "\" & companyName & "\" & productName & "\" & version
    BuildRoamingApplicationPath = path
End Function
