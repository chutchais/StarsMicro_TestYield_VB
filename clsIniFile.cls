VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
'        clsIniFile.cls  a  VB Class for simplifying .INI file operations
' ***************************************************************************
'  Original written by Karl E. Peterson, February 1994, CIS 72302,3707.
'  Portions originally downloaded from CompuServe's MSBASIC forum as
'  MINIFILE.BAS, author unknown.  Comments and questions welcome!
'
'  Modified to "Class" form by Dave Keighan, June 1998, dkeigan@home.com
'
'  This module was constructed using code declared "Free to be used ..."
'  permission of the author in it's construction.
'
'                       ****************************
'
'   *** Revisions: ***
'
'           Version 1.0.001     1999/07/20
'  Added Exists routine (I use it everywhere might as well have it here) - it is needed for the next change.
'  Added GetLongPathName Routine to return a Proper case Long Path for program use.
'  Changed ExtractPath routine to return a Proper Case path and remove trailing back slash.
'  Removed the conversion to UCASE in the ExtractPath and ExtractName routines.
'
'           Version 1.0.000     1998/06/25
'  Updated 16 bit calls to 32 bit alternates.
'  Added new calls to GetPrivateProfileSection and GetPrivateProfileSection.
'  Removed routines dealing with "Device Drivers".
'  Added ByVal and ByRef keywords as required (although "Optional" I like to
'    use the "make no mistake..." approach).
'  Update routines to handle the new 32 bit API calls and parse results into
'    tables where the original routines required.
'  Rewrote ExtractName routine to make it a little more flexible (added BaseOnly
'     and ExtensionOnly as an optional parameter and a Public Enum for the two
'     constants.
'  Modified comments and variable notation to suit VB5 conventions.
'  Added comments to various routines.
'
'                       ****************************
'                            ******************
'   *** Notes: ***
'
'  I have converted the various Dim, Public and Private variable statements to
'     my particular brand of "Polish Notation" (very similar to VB) whenever possible.
'
'  WritePrivateProfileSection and WriteProfileSection are currently only used
'     to delete the entries in a section. It is my intention to complete the
'     class with PutSections, PutSectionsEx, PutSectEntries and PutSectEntriesEx
'     but my actions have yet to catch up with my plans.
'
'  The [Section] return buffer is set to 4k if you expect to be dealing with
'     larger [Sections] increase the Max_SectionBuffer below. I've had it up to
'     32767 before the API call choked.
'
'  ... Enjoy
'  ... Dave.
'
'                       ****************************
'
'            ** Windows API calls **
'
'            ** For Private INIs
  Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
  Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
  Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
  Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
  Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

'            ** For Windows INIs
  Private Declare Function GetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
  Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
  Private Declare Function GetProfileSection Lib "kernel32" Alias "GetProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
  Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
  Private Declare Function WriteProfileSection Lib "kernel32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long

'            ** Fow Windows Communications
  Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
  Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

'           ** Constants used to size buffers
'  Set the max size of the string to be returned in bytes - use factors of 1024 to a MAX of 32767.
  Private Const Max_SectionBuffer As Integer = 4096    'bytes
'  Set the max size of an entry in bytes.
  Private Const Max_EntryBuffer As Integer = 255   'bytes

'           ** Special values to alert other apps of Win.Ini changes
  Private Const HWND_BROADCAST = &HFFFF
  Private Const WM_WININICHANGE = &H1A

'           ** Enum variables for ExtractPath routine.
Public Enum ExtractFileName_Constants
    exnFullName = 1
    exnOnlyExtn = 2
End Enum

'  Module level variables to hold property values
Private mstrSectionName As String   'local copy
Private mstrFileName As String         'local copy
Private mstrFilePath As String           'local copy
Private mstrFullPath As String           'local copy
Private mstrPrivPath As String           'local copy
Private mstrFileExt As String              'local copy
Private mbooInitialized As Boolean     'local copy
Private mbooUseWinINI As Boolean   'local copy

Public Property Get PrivPath() As String
   'The private path is a read only storage area.
   'The private path is set during initialization
   '   but can be changed by setting the FullPath property.
   '   It exists to allow switching to/from win.ini.
    PrivPath = mstrPrivPath

End Property

Public Property Let UseWinINI(ByVal bChoice As Boolean)

Dim strBuff As String
Dim intRet As Integer
Dim intDot As Integer

    If mbooUseWinINI = bChoice Then Exit Property Else mbooUseWinINI = bChoice
    
    If bChoice Then       'Win INI is to be used.
      'Find Win.Ini
      strBuff = String$(Max_EntryBuffer, 0)
      intRet = GetWindowsDirectory(strBuff, Max_EntryBuffer)
      mstrFullPath = Left$(strBuff, intRet) + "\WIN.INI"
      mstrFilePath = ExtractPath(mstrFullPath)
      mstrFileName = ExtractName(mstrFullPath, exnFullName)
      intDot = InStr(mstrFileName, ".")
      If intDot Then mstrFileExt = Right$(mvarFilName, Len(mstrFileName) - intDot)
    Else
      FullPath = mstrPrivPath
    End If
    

End Property

Public Property Get UseWinINI() As Boolean

    UseWinINI = mbooUseWinINI
    
End Property

Public Property Let FullPath(ByVal NewPath As String)

  'Setting the FullPath property is used for using a file other than the default or win.ini files.
  'Assign the Private Path String. Adjust other properties to suit.
  'Don't set this property to the Win.INI file it's redundant - simply set the UseWinINI to true.
  
  'Make sure you're using a valid path BEFORE setting this property!
  'No internal path checking is done!
    
    mstrPrivPath = NewPath
    mstrFullPath = mstrPrivPath
    mstrFilePath = ExtractPath(mstrFullPath)
    mstrFileName = ExtractName(mstrFullPath, exnFullName)
    mstrFileExt = ExtractName(mstrFullPath, exnOnlyExtn)
    If Trim(mstrSectionName) = "" Then mbooInitialized = False Else mbooInitialized = True
    mbooUseWinINI = False

End Property

Public Property Get FullPath() As String

    'The full path property can be read to find out what
    '  file is currently in use.
    FullPath = mstrFullPath
    
End Property

Public Property Let SectionName(ByVal SetSection As String)

    'The section name is set prior to data being read or
    '  written to the selected INI file.
    mstrSectionName = SetSection

    If Trim$(SetSection) = "" Then
        mbooInitialized = False
    Else
        mbooInitialized = True
    End If

End Property

Public Property Get SectionName() As String

    'Retrieves the current section name.
    SectionName = mstrSectionName

End Property

Public Property Get FilePath() As String

    'The file path property is read only and contains only the path
    '  of the current INI file.
    FilePath = mstrFilePath

End Property

Public Property Get FileExt() As String

    'The file extension property is read only and contains only the extension
    '  of the current INI file.
    FileExt = mstrFileExt

End Property

Public Function ExtractName(ByVal PathIn As String, ByVal RetChoice As ExtractFileName_Constants) As String
  
  Dim intCount As Integer
  Dim intDot As Integer
  Dim strSpecOut As String

  On Local Error Resume Next

 'The Extract Name function is used internally by the class but I
 '  have left it as Public so that it is available for othe purposes.
 'If you use it for other purposes - use it properly.
  
  If Len(PathIn) = 0 Then PathIn = mstrFullPath
  
  'Extract and return the full file name.
  If InStr(PathIn, "\") Then
    For intCount = Len(PathIn) To 1 Step -1
      If Mid$(PathIn, intCount, 1) = "\" Then
        strSpecOut = Mid$(PathIn, intCount + 1)
        Exit For
      End If
    Next intCount
  ElseIf InStr(PathIn, ":") = 2 Then
    strSpecOut = Mid$(PathIn, 3)
  Else
    strSpecOut = PathIn
  End If
  
  intDot = InStr(strSpecOut, ".")
  
  'Returns only the base of the file name.
  If intDot And RetChoice = exnFullName Then strSpecOut = Left$(strSpecOut, intDot - 1)
  
  'Returns only the extension of the filename.
  If intDot And RetChoice = exnOnlyExtn Then strSpecOut = Right$(strSpecOut, Len(strSpecOut) - intDot)

  ExtractName = strSpecOut
  
End Function

Public Function ExtractPath(ByVal PathIn As String) As String

  Dim intCount As Integer
  Dim strSpecOut As String
  
  On Local Error Resume Next

 'The Extract Path function is used internally by the class but I
 '  have left it as Public so that it is available for othe purposes.
 'If you use it for other purposes - use it properly.
 
  If Len(PathIn) = 0 Then PathIn = mstrFullPath

  If InStr(PathIn, "\") Then
    For intCount = Len(PathIn) To 1 Step -1
      If Mid$(PathIn, intCount, 1) = "\" Then
        strSpecOut = Left$(PathIn, intCount - 1)        'Reduced length of strSpecOut by one - 99/07/17.
        Exit For
      End If
    Next intCount
  ElseIf InStr(PathIn, ":") = 2 Then
    strSpecOut = CurDir$(PathIn)
    If Len(strSpecOut) = 0 Then strSpecOut = CurDir$
  Else
    strSpecOut = CurDir$
  End If
    
  If Right$(strSpecOut, 1) = "\" Then
    strSpecOut = Left(strSpecOut, Len(strSpecOut) - 1)
  End If
  
  ExtractPath = strSpecOut

End Function

Public Sub DeleteStrEntry(ByVal EntryName As String)

  'Bail if not initialized
    If Not mbooInitialized Then
      WarnAuthor
      Exit Sub
    End If

  'Sets a specific entry to Nothing or Blank.
    Dim lngRetVal As Long
    If mbooUseWinINI Then
        lngRetVal = WriteProfileString(mstrSectionName, EntryName, "")
        WinIniChanged
    Else
        lngRetVal = WritePrivateProfileString(mstrSectionName, EntryName, "", mstrFullPath)
    End If

End Sub

Public Sub DeleteNumEntry(ByVal EntryName As String)

  'Bail if not initialized.
    If Not mbooInitialized Then
      WarnAuthor
      Exit Sub
    End If
  
  'Deletes a specific entry.
    Dim lngRetVal As Long
    If mbooUseWinINI Then
        lngRetVal = WriteProfileString(mstrSectionName, EntryName, 0&)
        WinIniChanged
    Else
        lngRetVal = WritePrivateProfileString(mstrSectionName, EntryName, 0&, mstrFullPath)
    End If
    
End Sub

Public Sub DeleteSection()

  'Bail if not initialized
    If Not mbooInitialized Then
      WarnAuthor
      Exit Sub
    End If

    'Deletes all of the current [Section]s Entries.
    Dim lngRetVal As Long
    If mbooUseWinINI Then
      lngRetVal = WriteProfileSection(mstrSectionName, "")
      WinIniChanged
    Else
      lngRetVal = WritePrivateProfileSection(mstrSectionName, "", mstrFullPath) 'WritePrivateProfileString(mstrSectionName, 0&, 0&, mstrFullPath)
    End If
      
    mstrSectionName = ""
    mbooInitialized = False


End Sub

Function Exists(ByVal filname As String) As Integer

' ************************************************************************
'  Name   :  Exists
'  Author :  Dave Keighan
'  Date   :  1999/07/07
'  Params :  ByVal filname As String
'  Returns:  As Integer
'  Purpose:
'
'  Code Chnages:
'  Date      By         What
'  --------  ---------- --------------------------------------
'
' ************************************************************************

'Set up default error handling
On Error GoTo Err_Exists:

'******** Coding starts here ********
' returns true if the file "filname" exist

If InStr(filname, ".") < 1 Then
    filname = filname + ".*"
End If

Exists = Dir$(filname) > ""

Exit Function

'********* Coding ends here *********
Exit Function

'Default error handler
Err_Exists:

Select Case Err
    Case Else
        'Call the default error handler
        myERROR.Handler "WINPARTS.BAS", "Exists", Err.Source, Err.Description, Err.Number
        'NB Default error handling ends procedure
  Resume End_Exists:
End Select

End_Exists:

End Function

Public Function GetInt(ByVal EntryName As String, ByVal DefaultInt As Integer) As Integer

  'Bail if not initialized
    If Not mbooInitialized Then
      WarnAuthor
      Exit Function
    End If

  'Retrieves an Integer value, range: 0-32767
    If mbooUseWinINI Then
        GetInt = GetProfileInt(mstrSectionName, EntryName, DefaultInt)
    Else
        GetInt = GetPrivateProfileInt(mstrSectionName, EntryName, DefaultInt, mstrFullPath)
    End If

End Function

Public Function GetSectEntries() As String

  'Bail if not initialized.
    If Not mbooInitialized Then
      WarnAuthor
      Exit Function
    End If

    'Retrieves all Entries in a [Section]
    'Returnes a string of Null delineated entries "EntryName=Value&vbNull&...." with the
    '  last entry double-terminated.
    
    Dim strTemp As String * Max_SectionBuffer
    Dim lngRetVal As Long
    If mbooUseWinINI Then
        lngRetVal = GetProfileSection(mstrSectionName, strTemp, Len(strTemp))
    Else
        lngRetVal = GetPrivateProfileSection(mstrSectionName, strTemp, Len(strTemp), mstrFullPath)
    End If
    
    GetSectEntries = Left$(strTemp, lngRetVal + 1)
    
End Function

Public Function GetSectEntriesEx(ByRef DataArry() As String) As Integer

  'Bail if not initialized.
    If Not mbooInitialized Then
      WarnAuthor
      Exit Function
    End If

  'GetSectionEntriesEx retrieves all of the entries in the current [Section] and
  '  returns tha values in the two dimensional DataArry(1, Entries) array.
  '  The 0 Row holding the Entry Name and the 1 holding the Values.
  '  There will be as many Columns as there are Entries.

  On Error Resume Next

  'Get "normal" null terminated string of all [Section] Entries.
    Dim strTemp As String
    strTemp = GetSectEntries()
    If Len(strTemp) = 1 Then
      GetSectionsEntriesEx = 0
      Exit Function
    End If
  
  'Parse null terminated string of [Section] Entries into table.
    Dim intEntries As Integer
    Dim intNull As Integer
    Dim intLoc As Integer
    Do While Asc(strTemp)
      ReDim Preserve DataArry(0 To 1, 0 To intEntries)
      intNull = InStr(strTemp, Chr$(0))
      intLoc = InStr(Left$(strTemp, intNull - 1), "=")
      DataArry(0, intEntries) = Left$(strTemp, intLoc - 1)
      DataArry(1, intEntries) = Mid$(strTemp, intLoc + 1, intNull - intLoc - 1)
      strTemp = Mid$(strTemp, intNull + 1)
      intEntries = intEntries + 1
      If strTemp = "" Then Exit Do
    Loop

  'Make function assignment
    GetSectEntriesEx = intEntries

End Function

Public Function GetSections() As String

  'Bail if not initialized
    If Not mbooInitialized Then
      WarnAuthor
      Exit Function
    End If

 'The GetSections function is used to return a Null delineated string of all
 '  the [Section]s in current file.
  
  'Setup some variables
    Dim strRet As String
    Dim strBuff As String
    Dim intFileHandle As Integer

  'Extract all [Section] lines
    intFileHandle = FreeFile
    Open mstrFullPath For Input As intFileHandle
    Do While Not EOF(intFileHandle)
      Line Input #intFileHandle, strBuff
      strBuff = StripComment(strBuff)
      If InStr(strBuff, "[") = 1 And InStr(strBuff, "]") = Len(strBuff) Then
        strRet = strRet + Mid$(strBuff, 2, Len(strBuff) - 2) + vbNullChar
      End If
    Loop
    Close intFileHandle

  'Assign return value
    If Len(strRet) Then
      GetSections = strRet + Chr$(0)
    Else
      GetSections = String$(2, 0)
    End If

End Function

Public Function GetSectionsEx(ByRef DataArry() As String) As Integer

 'The GetSectionsEx function is used to return an array of all the [Section]s
 '  in current file.
  

  'Get "normal" list of all [Section]'s
    Dim strSect As String
    strSect = GetSections()
    If Len(strSect) = 0 Then
      GetSectionsEx = 0
      Exit Function
    End If

  'Parse [Section]'s into table
    Dim intEntries As Integer
    Dim intNull As Integer
    Do While Asc(strSect)
      ReDim Preserve DataArry(0 To intEntries)
      intNull = InStr(strSect, vbNullChar)
      DataArry(intEntries) = Left$(strSect, intNull - 1)
      strSect = Mid$(strSect, intNull + 1)
      intEntries = intEntries + 1
    Loop

  'Make function assignment (number of [Sections]s added to table).
    GetSectionsEx = intEntries
  
End Function

Public Function GetString(ByVal EntryName As String, ByVal DefaultStr As String) As String

  'Bail if not initialized
    If Not mbooInitialized Then
      WarnAuthor
      Exit Function
    End If

  'Retrieves Specific string Entry from INI.
    Dim strTemp As String * Max_EntryBuffer
    Dim lngRetVal As Long
    If mbooUseWinINI Then
        lngRetVal = GetProfileString(mstrSectionName, EntryName, DefaultStr, strTemp, Len(strTemp))
    Else
        lngRetVal = GetPrivateProfileString(mstrSectionName, EntryName, DefaultStr, strTemp, Len(strTemp), mstrFullPath)
    End If
    
    If lngRetVal Then
      GetString = Left$(strTemp, lngRetVal)
    End If

End Function

Public Function GetTF(ByVal EntryName As String, ByVal DefaultInt As Integer) As Boolean
  
  'Retrieves Specific Entry as either True/False.
  'local vars
    Dim strTF As String
    Dim strDefault As String

  'get string value from INI
    If DefaultInt Then
      strDefault = "True"
    Else
      strDefault = "False"
    End If
    
    strTF = GetString(EntryName, strDefault)
    
  'interpret return string and translate to T/F.
    Select Case Trim$(UCase$(strTF))
      Case "YES", "Y", "TRUE", "T", "ON", "1", "-1"
        GetTF = True
      Case "NO", "N", "FALSE", "F", "OFF", "0"
        GetTF = False
      Case Else
        GetTF = False
    End Select

End Function

Public Sub FlushCache()

  'Bail if not initialized
    If Not mbooInitialized Then
      WarnAuthor
      Exit Sub
    End If

  'To improve performance, Windows keeps a cached version of the most-recently
  'accessed initialization file. If that filename is specified and the other
  'three parameters are NULL, Windows flushes the cache
    Dim lngRetVal As Long
    If mbooUseWinINI Then
        lngRetVal = WriteProfileString(0&, 0&, 0&)
    Else
        lngRetVal = WritePrivateProfileString(0&, 0&, 0&, mstrFullPath)
    End If

End Sub

Private Sub WarnAuthor()
  
  'Warn *PROGRAMMER* that there's a logic error!
    MsgBox "[Section] and FileName Not Registered in Private.Ini!", vbInformation + vbOKOnly, "IniFile Logic Error"

End Sub

Public Sub IniRead(ByVal SectionName As String, ByVal EntryName As String, ByVal DefaultInt As Integer, ByVal DefaultStr As String, ByVal ReturnStr As String, ByVal ReadNumber As Boolean, ByVal FullPath As String)

  'One-shot read from Ini, more *work* than it's worth
  'It does not use any of the claas properties (except the UseWinIni value) - only what you pass to it!
  
    Dim lngRetVal As Long
    Dim RetStr As String * Max_EntryBuffer 'Create an empty string to be filled

    If ReadNumber Then     'we are looking for integer input
        If mbooUseWinINI Then
            ReadNumber = GetProfileInt(SectionName, EntryName, DefaultInt)
        Else
            ReadNumber = GetPrivateProfileInt(SectionName, EntryName, DefaultInt, FullPath)
        End If
    Else
        If mbooUseWinINI Then
            lngRetVal = GetProfileString(SectionName, EntryName, DefaultStr, strTemp, Len(strTemp))
        Else
            lngRetVal = GetPrivateProfileString(SectionName, EntryName, DefaultStr, strTemp, Len(strTemp), FullPath)
        End If
        If lngRetVal Then
          ReturnStr = Left$(RetStr, lngRetVal)
        End If
    End If

End Sub
Private Sub WinIniChanged()
  
  'Notify all other applications that Win.Ini has been changed
    Dim rtn&
    'Rtn = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0&, ByVal mstrSectionName)

End Sub

Public Sub IniWrite(ByVal SectionName As String, ByVal EntryName As String, ByVal NewVal As String, ByVal FullPath As String)
    
    Dim lngRetVal As Long
    
    'One-shot write to Private.Ini, more *work* than it's worth.
    'It does not use any of the claas properties (except the UseWinIni value) - only what you pass to it!
        
        If mbooUseWinINI Then
            lngRetVal = WriteProfileString(SectionName, EntryName, NewVal)
        Else
            lngRetVal = WritePrivateProfileString(SectionName, EntryName, NewVal, FullPath)
        End If

    
End Sub

Public Function PutInt(ByVal EntryName As String, ByVal IntValue As Integer) As Integer

  'Bail if not initialized
    If Not mbooInitialized Then
      WarnAuthor
      Exit Function
    End If
  
  'Write an integer to INI
    If mbooUseWinINI Then
        PutInt = WriteProfileString(mstrSectionName, EntryName, CStr(IntValue))
        WinIniChanged
    Else
        PutInt = WritePrivateProfileString(mstrSectionName, EntryName, CStr(IntValue), mstrFullPath)
    End If
    
End Function


Public Function PutString(ByVal EntryName As String, ByVal StrValue As String) As Integer

  'Bail if not initialized
    If Not mbooInitialized Then
      WarnAuthor
      Exit Function
    End If

  'Write a string to INI
    If mbooUseWinINI Then
        PutString = WriteProfileString(mstrSectionName, EntryName, StrValue)
        WinIniChanged
    Else
        PutString = WritePrivateProfileString(mstrSectionName, EntryName, StrValue, mstrFullPath)
    End If
    
End Function

Public Function PutTF(ByVal EntryName As String, ByVal IntValue As Integer) As Boolean

  'Set an entry in .Ini to True/False
  'local vars
    Dim strTF As String

  'translate the value  to a string.
    If IntValue Then
      strTF = "True"
    Else
      strTF = "False"
    End If

    'enter the value in the INI
    PutTF = PutString(EntryName, strTF)
    
    If mbooUseWinINI Then WinIniChanged
    
End Function

Public Function SectExist(ByVal SectionName As String) As Boolean

  'Retrieve list of all [Section]'s
    Dim strSect As String
    strSect = GetSections()
    If Len(strSect) = 0 Then
      SectExist = False
      Exit Function
    End If

  'Check for existence registered [Section]
    strSect = Chr$(0) + UCase$(strSect)
    If InStr(strSect, Chr$(0) + UCase$(SectionName) + Chr$(0)) Then
      SectExist = True
    Else
      SectExist = False
    End If

End Function

Private Function StripComment(ByVal StrIn As String) As String
  Dim intRet As Integer
  'Check for comment
    intRet = InStr(StrIn, ";")

  'Remove it if present
    If intRet = 1 Then
      'Whole string is a comment
        StripComment = ""
        Exit Function
    ElseIf intRet > 1 Then
      'Strip comment
        StrIn = Left$(StrIn, intRet - 1)
    End If
  
  'Trim any trailing space
    StripComment = Trim$(StrIn)

End Function

Public Property Get FileName() As String

    FileName = mstrFileName

End Property

Private Sub Class_Initialize()

    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & mlClassDebugID & " created"
    #End If
    
    If Trim(App.EXEName) = "" Then Exit Sub
    
    'Written to use long path names instead of the short path name returned by the App Object.
    'If the class is initialised early in the application the long path name will be supplied.
        FullPath = GetLongPathName(App.Path) & "\" & StrConv(App.EXEName, vbProperCase) & ".ini"
    
    mstrSectionName = "General"

End Sub

Private Sub Class_Terminate()
    
    'the class is being destroyed
    #If DebugMode Then
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If

End Sub

Private Function GetLongPathName(ByVal strShortName As String) As String

Dim strLongName As String
Dim strTemp As String
Dim intSlashPos As Integer

If Len(strShortName) < 1 Then Exit Function

'Check to see if the root directory has been passed.
If Right$(strShortName, 2) = ":\" Then
    GetLongPathName = strShortName
    Exit Function
End If

'Clip off the trailing back-slash
If Right$(strShortName, 1) <> "\" Then
    strShortName = strShortName & "\" 'Left$(strShortName, Len(strShortName) - 1)
End If

intSlashPos = InStr(4, strShortName, "\")

While intSlashPos
    strTemp = Dir(Left$(strShortName, intSlashPos - 1), vbNormal + vbHidden + vbSystem + vbDirectory)
    If strTemp = "" Then
        GetLongPathName = ""
        Exit Function
    End If
    strLongName = strLongName & "\" & strTemp
    intSlashPos = InStr(intSlashPos + 1, strShortName, "\")
Wend

GetLongPathName = Left$(strShortName, 2) & strLongName

End Function
