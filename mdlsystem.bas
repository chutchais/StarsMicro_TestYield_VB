Attribute VB_Name = "mdlsystem"
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias _
                "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                ByVal lpKeyName As Any, ByVal lpDefault As String, _
                ByVal lpReturnedString As String, ByVal nSize As Long, _
                ByVal lpFileName As String) As Long


Function FileToString(strFilename As String) As String
  IFile = FreeFile
  Open strFilename For Input As #IFile
    FileToString = StrConv(InputB(LOF(IFile), IFile), vbUnicode)
  Close #IFile
End Function


Function textContain(vString As String, vText As String) As Boolean
    Dim position As Integer
    position = InStr(1, vString, vText)
    
    If position > 0 Then
        textContain = True
    End If
End Function




'INI
Function getSectionString(sectionName As String, entryName As String) As String
    Dim vIniFile As String
    vIniFile = App.Path & "\system.ini"
    Dim test  As String
    sRetBuf$ = String$(256, 0)
    iLenBuf% = Len(sRetBuf$)
    X = GetPrivateProfileString(sectionName, entryName, _
                        "", sRetBuf$, iLenBuf%, vIniFile)
    getSectionString = Left$(sRetBuf$, X)
End Function



