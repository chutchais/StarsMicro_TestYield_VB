VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTMT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim vFileContains As String
Dim vLogfileName As String
Dim vTestProgram As String
Dim vTested As Double
Dim vPassed As Double
Dim vFailed As Double

Dim vAssy As String
Dim vLot As String
Dim vSeq As String

Dim vOperator As String
Dim vMostFailBin As String
Dim vComputer As String
Dim vBinNumber As String
Dim vHandler As String
Dim vYield As String
Dim vAutocorrelation As String
Dim vNextSerial As String
Dim vStartDate As String
Dim vStopDate As String
Dim vTemperature As String

Dim objHwBins As New Collection
Dim objSwBins As New Collection

'COntrol
Dim vSwBinLineStart As Boolean
Dim vSwSiteLineStart As Boolean
Dim vSwSiteTotal As Integer

'Status
Dim vCompleted As Boolean
Dim vDescription As String


'Start Property
Public Property Get Completed() As String
    Completed = vCompleted
End Property
Public Property Get Description() As String
    Description = vDescription
End Property

Public Property Get LogFileName() As String
    LogFileName = vLogfileName
End Property

Public Property Get TestProgram() As String
    TestProgram = vTestProgram
End Property

Public Property Get Tested() As Double
    Tested = vTested
End Property

Public Property Get Passed() As Double
    Passed = vPassed
End Property
Public Property Get Failed() As Double
    Failed = vFailed
End Property

Public Property Get AssyNumber() As String
    AssyNumber = vAssy
End Property

Public Property Get Lot() As String
    Lot = vLot
End Property

Public Property Get Seq() As String
    Seq = vSeq
End Property

Public Property Get Operator() As String
    Operator = vOperator
End Property

Public Property Get MostFailBin() As String
    MostFailBin = vMostFailBin
End Property

Public Property Get Computer() As String
    Computer = vComputer
End Property

Public Property Get BinNumber() As String
    BinNumber = vBinNumber
End Property

Public Property Get Handler() As String
    Handler = vHandler
End Property

Public Property Get Yield() As String
    Yield = vYield
End Property

Public Property Get Autocorrelation() As String
    Autocorrelation = vAutocorrelation
End Property

Public Property Get NextSerial() As String
    NextSerial = vNextSerial
End Property

Public Property Get StartDate() As String
    StartDate = vStartDate
End Property

Public Property Get StopDate() As String
    StopDate = vStopDate
End Property

Public Property Get HardwareBins() As Collection
  
    Set HardwareBins = objHwBins
    
End Property

Public Property Get SoftwareBins() As Collection
    Set SoftwareBins = objSwBins
End Property
Public Property Get Temperature() As String
    Temperature = vTemperature
End Property


'End Property

Friend Sub Init(ByVal file_name As String)
  If file_name = "" Then Err.Raise 123, , "File name not set"
  'Init object state here
  'vLogfileName = file_name
  vFileContains = FileToString(file_name)
  
  'Get Seq from file Name
  Dim vOnlyFileName As String
  vOnlyFileName = get_only_fileName(file_name)
  vLogfileName = vOnlyFileName
  
  Dim vFileSplited() As String
  vFileSplited = Split(vOnlyFileName, "_")
  For i = 2 To UBound(vFileSplited)
    vSeq = vSeq + "_" + vFileSplited(i)
  Next
  vSeq = Replace(vSeq, ".lsr", "")
  vSeq = Mid(vSeq, 2)
  '----------------------------------
  process_file
End Sub

Function get_only_fileName(vFullPath As String) As String
    Dim vFileArray() As String
    vFileArray = Split(vFullPath, "\")
    If UBound(vFileArray) > 0 Then
        get_only_fileName = vFileArray(UBound(vFileArray))
    Else
        get_only_fileName = ""
    End If
End Function

Private Sub process_file()
On Error GoTo HasError
    Dim vLineStr As String
    Dim vKey As String
    Dim FileArray() As String
    FileArray = Split(vFileContains, vbCrLf)
    
    For i = 0 To UBound(FileArray())
        vLineStr = FileArray(i)
        If vLineStr = "" Then
            GoTo NextLine
        End If
        
        If textContain(vLineStr, "SW Bins") Or textContain(vLineStr, "HW Bins") Then vSwBinLineStart = True
        If textContain(vLineStr, "SW Site 1") Then vSwBinLineStart = False: vSwSiteLineStart = True
        
        'Hardware Bin and Software Bin
        If vSwBinLineStart Then
            getSwAndHwBin vLineStr
            GoTo NextLine
        End If
        
        'Software Bin Site information
        If vSwSiteLineStart Then
            getSwSite vLineStr
            GoTo NextLine
        End If
        
        
        'Test Program and Tested
        vKey = "Test Program    :"
        If textContain(vLineStr, vKey) Then getTestProgram vLineStr
        
        'Passed
        vKey = "Total Pass   :"
        If textContain(vLineStr, vKey) Then getPassed vLineStr
        
        'Assy,Lot and Seq , Failed
        vKey = "Lot ID          :"
        If textContain(vLineStr, vKey) Then getLot vLineStr
        
        'Operator , Most failed Bin
        vKey = "Operator        :"
        If textContain(vLineStr, vKey) Then getOperator vLineStr
        
        'Computer , Bin#
        vKey = "Computer        :"
        If textContain(vLineStr, vKey) Then getComputer vLineStr
        
        'Handler , Yield
        vKey = "Handler         :"
        If textContain(vLineStr, vKey) Then getHandler vLineStr
        
        'Autocorrelation , Next Serial
        vKey = "Autocorrelation :"
        If textContain(vLineStr, vKey) Then getAutocorrelation vLineStr
        
        'Start/Stop Date
        vKey = "/"
        If textContain(vLineStr, vKey) Then getTestDate vLineStr
        

NextLine:
        
    Next
    
    vCompleted = True
    Exit Sub
HasError:
    vDescription = "Error on line " & i & " : " & Err.Description
End Sub

Private Sub getTestProgram(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, " ")
    
    'Get Test Program name
    Dim ixColon As Integer
    ixColon = getColonPosition(vLineData)
    vTestProgram = vData(ixColon + 1)
    
    'Get Tested
    ixColon = getColonPosition(vLineData, ixColon + 1)
    vTested = vData(ixColon + 1)
End Sub

Private Sub getPassed(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, " ")
    Dim ixColon As Integer
    ixColon = getColonPosition(vLineData)
    ixColon = getColonPosition(vLineData, ixColon + 1)
    vPassed = vData(ixColon + 1)
End Sub

Private Sub getLot(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, " ")
    
    'get Assy,Lot and Seq
    Dim ixColon As Integer
    ixColon = getColonPosition(vLineData)
    Dim vLotArry() As String
    vLotArry = Split(vData(ixColon + 1), "_")
    vAssy = vLotArry(0)
    vLot = vLotArry(1)
    'vSeq = Mid(vLotArry(2), 1, 2)
    
    'get Failed
    ixColon = getColonPosition(vLineData, ixColon + 1)
    vFailed = vData(ixColon + 1)

End Sub

Private Sub getOperator(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, " ")
    
    'get Operator
    Dim ixColon As Integer
    ixColon = getColonPosition(vLineData)
    vOperator = vData(ixColon + 1)
    'get MostBin
    ixColon = getColonPosition(vLineData, ixColon + 1)
    vMostFailBin = vData(ixColon + 1)
End Sub

Private Sub getComputer(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, " ")
    
    'get Computer
    Dim ixColon As Integer
    ixColon = getColonPosition(vLineData)
    vComputer = vData(ixColon + 1)
    'get Bin number
    ixColon = getColonPosition(vLineData, ixColon + 1)
    vBinNumber = vData(ixColon + 1)
End Sub

Private Sub getHandler(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, " ")
    
    'get Handler
    Dim ixColon As Integer
    ixColon = getColonPosition(vLineData)
    vHandler = vData(ixColon + 1)
    'get Yield
    ixColon = getColonPosition(vLineData, ixColon + 1)
    vYield = vData(ixColon + 1)
End Sub

Private Sub getAutocorrelation(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, " ")
    
    'get Autocorrelation
    Dim ixColon As Integer
    ixColon = getColonPosition(vLineData)
    vAutocorrelation = vData(ixColon + 1)
    'get Next Serial
    ixColon = getColonPosition(vLineData, ixColon + 1)
    vNextSerial = vData(ixColon + 1)
End Sub

Private Sub getTestDate(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "/")
    
    If IsDate(vData(0)) Then
        Exit Sub
    End If
    
    'get Autocorrelation
    vStartDate = vData(0)
    vStopDate = vData(1)
End Sub


Private Sub getSwAndHwBin(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, " ")
    If vData(0) = "" Then Exit Sub
    
    'Software bin
    Dim vSwBinNumber As String
    Dim vSwDescription As String
    Dim vSwTotal As Double
    Dim vSwPercent As String
    Dim ixPercent As Integer
    'Get Bin number
    vSwBinNumber = Replace(Replace(vData(0), "[", ""), "]", "")
    'get Bin Description
    ixPercent = getPercentPosition(vLineData, 1)
    For i = 1 To 10
        If vData(i) <> "" Then vSwDescription = vSwDescription & " " & vData(i)
    Next
    vSwDescription = Trim(vSwDescription)
    
    'Get Total
    For i = 11 To ixPercent - 2
        If vData(i) <> "" Then
            If Val(vData(i)) >= 0 Then vSwTotal = Val(vData(i)): Exit For
        End If
    Next
    'Get Percent
    vSwPercent = vData(ixPercent - 1)
    'Add SW Bin in to collection
    Dim objBin As New clsBin
    With objBin
        .Name = vSwBinNumber
        .Description = vSwDescription
        .Total = vSwTotal
        .Yield = vSwPercent
    End With
    objSwBins.Add objBin
    '---------------------------
    
    'Hardware Bin
    Dim vHwBinNumber As String
    Dim vHwTotal As Double
    Dim vHwYield As Double
    If InstrCount(vLineData, "%") = 2 Then
        'MsgBox ""
        For i = ixPercent + 1 To UBound(vData)
            If vData(i) <> "" Then
                If Val(vData(i)) >= 0 Then vHwBinNumber = vData(i): ixPercent = i: Exit For
            End If
        Next
        
        For i = ixPercent + 1 To UBound(vData)
            If vData(i) <> "" Then
                If Val(vData(i)) >= 0 Then vHwTotal = Val(vData(i)):  Exit For
            End If
        Next
        
        ixPercent = getPercentPosition(vLineData, ixPercent)
        vHwYield = vData(ixPercent - 1)
        
        'Add HW Bin in to collection
        Dim objHWBin As New clsBin
        With objHWBin
            .Name = vHwBinNumber
            .Description = ""
            .Total = vHwTotal
            .Yield = vHwYield
        End With
        objHwBins.Add objHWBin
    End If
    
End Sub

Private Sub getSwSite(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, " ")
    If vLineData Like "*Site*" Then
        vSwSiteTotal = InstrCount(vLineData, "Site")
    End If
    If vData(0) = "" Then Exit Sub
    
    Dim vSwSiteNumber As String
    Dim vSwBinNumber As String
    Dim vSiteTotal As Double
    Dim vSiteYield As Double
    Dim ixPercent As Integer
    Dim ixStart As Integer
    ixStart = 1
    vSwBinNumber = Replace(Replace(vData(0), "[", ""), "]", "")
    For i = 1 To vSwSiteTotal
        vSwSiteNumber = i
        ixPercent = getPercentPosition(vLineData, ixPercent + 1)
        'get Total
        For x = ixStart To ixPercent - 2
            If vData(x) <> "" Then
                If Val(vData(x)) >= 0 Then vSiteTotal = Val(vData(x)):  Exit For
            End If
        Next
        ixStart = ixPercent + 1
        'get Yield
        vSiteYield = vData(ixPercent - 1)
        'Add SW Bin in to collection
        Dim objSite As New clsSite
        With objSite
            .Name = vSwSiteNumber
            .Description = ""
            .Total = vSiteTotal
            .Yield = vSiteYield
        End With
        'Find SW bin in collection and add site to BIN
        Dim objCurrentBin As New clsBin
        Set objCurrentBin = TryGet(vSwBinNumber, objSwBins)
        objCurrentBin.Sites.Add objSite
        Set objSite = Nothing
        Set objCurrentBin = Nothing
        '---------------------------
    Next
    
    
End Sub


'Common function for Class
Public Function getBin(key As String, col As Collection) As Object
  On Error GoTo errhandler
  For Each c In col
        If c.Name = key Then
            Set getBin = c
            Exit For
        End If
  Next
  Exit Function
errhandler:
  Set getBin = Nothing
End Function

Private Function TryGet(key As String, col As Collection) As clsBin
  On Error GoTo errhandler
  For Each c In col
        If c.Name = key Then
            Set TryGet = c
            Exit For
        End If
  Next
  Exit Function
errhandler:
  Set TryGet = Nothing
End Function

Private Function InstrCount(StringToSearch As String, _
           StringToFind As String) As Long

    If Len(StringToFind) Then
        InstrCount = UBound(Split(StringToSearch, StringToFind))
    End If
End Function

Private Function getColonPosition(vDataString As String, Optional start As Integer = 0) As Integer
    Dim vData() As String
    vData = Split(vDataString, " ")
    For i = start To UBound(vData)
       If vData(i) Like "*:*" Then
            getColonPosition = i
            Exit Function
       End If
    Next
End Function

Private Function getPercentPosition(vDataString As String, Optional start As Integer = 0) As Integer
    Dim vData() As String
    vData = Split(vDataString, " ")
    For i = start To UBound(vData)
       If vData(i) Like "*%*" Then
            getPercentPosition = i
            Exit Function
       End If
    Next
End Function





