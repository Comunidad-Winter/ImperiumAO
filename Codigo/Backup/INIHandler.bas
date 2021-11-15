Attribute VB_Name = "INIH"

Declare Function WritePrivateProfileSection Lib _
"kernel32" Alias "WritePrivateProfileSectionA" _
(ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib _
"kernel32" Alias "WritePrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Declare Function GetPrivateProfileSection Lib _
"kernel32" Alias "GetPrivateProfileSectionA" _
(ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function GetPrivateProfileString Lib _
"kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Function ReadINI(Section As String, KeyName As String, FileName As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName$, "", sRet, Len(sRet), FileName))
End Function

Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim r
    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function

Function FileExist(file As String, FileType As VbFileAttribute) As Boolean
If Dir(file, FileType) = "" Then
    FileExist = False
Else
    FileExist = True
End If
End Function

