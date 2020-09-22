Attribute VB_Name = "modShell_lnk_url"
Option Explicit
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" _
(ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, _
ByVal hIcon As Long) As Long

'---------------------------------------------------------------
'- Public API Declares...
'---------------------------------------------------------------
#If UNICODE Then
    Public Declare Function SHGetPathFromIDList Lib "Shell32" Alias "SHGetPathFromIDListW" (ByVal pidl As Long, ByVal szPath As Long) As Long
#Else
    Public Declare Function SHGetPathFromIDList Lib "Shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal szPath As String) As Long
#End If

Public Declare Function SHGetSpecialFolderLocation Lib "Shell32" (ByVal hwndOwner As Long, ByVal nFolder As Integer, ppidl As Long) As Long

'---------------------------------------------------------------
'- Public constants...
'---------------------------------------------------------------
Public Const MAX_PATH = 255
Public Const MAX_NAME = 40

'Moved to cShellLiink so it will be in the .dll
''Public Sub AboutBox(hwnd As Long)
'-----------------------------------------------------------------
    ' Show help about dialog...
''    ShellAbout hwnd, App.EXEName, _
''               vbCrLf & "BJ. E-Mail me: bryce3@bigpond.com", 0
'-----------------------------------------------------------------
''End Sub


