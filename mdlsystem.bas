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


Public Function SortCollection(ByRef colInput As Collection, Optional bAsc = True) As Collection

    Dim varTemp                 As Object
    Dim lngCounter              As Long
    Dim lngCounter2             As Long

    For lngCounter = 1 To colInput.Count - 1
        For lngCounter2 = lngCounter + 1 To colInput.Count
            Select Case bAsc
            Case True:
                If colInput(lngCounter).seq > colInput(lngCounter2).seq Then
                    Set varTemp = colInput(lngCounter2)
                    colInput.Remove lngCounter2
                    colInput.Add varTemp
                End If

            Case False:
                If colInput(lngCounter) < colInput(lngCounter2) Then
                    varTemp = colInput(lngCounter2)
                    colInput.Remove lngCounter2
                    colInput.Add varTemp, varTemp, lngCounter
                End If
            End Select
        Next lngCounter2
    Next lngCounter

    Set fnVarBubbleSort = colInput

End Function
   


'INI
Function getSectionString(SectionName As String, EntryName As String) As String
    Dim vIniFile As String
    vIniFile = App.Path & "\system.ini"
    Dim test  As String
    sRetBuf$ = String$(256, 0)
    iLenBuf% = Len(sRetBuf$)
    x = GetPrivateProfileString(SectionName, EntryName, _
                        "", sRetBuf$, iLenBuf%, vIniFile)
    getSectionString = Left$(sRetBuf$, x)
End Function



