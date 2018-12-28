Attribute VB_Name = "mdlsystem"
Function FileToString(strFilename As String) As String
  iFile = FreeFile
  Open strFilename For Input As #iFile
    FileToString = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
End Function


Function textContain(vString As String, vText As String) As Boolean
    Dim position As Integer
    position = InStr(1, vString, vText)
    
    If position > 0 Then
        textContain = True
    End If
End Function
