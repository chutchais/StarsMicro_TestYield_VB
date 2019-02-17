VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsETS"
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
Dim vYield As Double

Dim vAssy As String
Dim vLot As String
Dim vSeq As String
'
Dim vOperator As String
Dim vTemperature As String
Dim vTester As String
'Dim vMostFailBin As String
'Dim vComputer As String
'Dim vBinNumber As String
Dim vHandler As String
Dim vStartDate As String
Dim vStartTime As String
'Dim vStopDate As String
Dim vSystemId As String
Dim vTestName As String
'
Dim objHwBins As New Collection
Dim objSwBins As New Collection
'
'COntrol
Dim vTestedLineStart As Boolean
Dim vSwBinLineStart As Boolean
Dim vHwBinLineStart As Boolean
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
    Failed = vTested - vPassed
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
'
Public Property Get Operator() As String
    Operator = vOperator
End Property

Public Property Get Temperature() As String
    Temperature = vTemperature
End Property

Public Property Get Tester() As String
    Tester = vTester
End Property

Public Property Get TestName() As String
    TestName = vTestName
End Property
'
'Public Property Get MostFailBin() As String
'    MostFailBin = vMostFailBin
'End Property
'
'Public Property Get Computer() As String
'    Computer = vComputer
'End Property
'
'Public Property Get BinNumber() As String
'    BinNumber = vBinNumber
'End Property
'
Public Property Get Handler() As String
    Handler = vHandler
End Property
'
Public Property Get Yield() As Double
    Yield = vYield
End Property

Public Property Get StartDate() As String
    StartDate = vStartDate
End Property

Public Property Get StartTime() As String
    StartTime = vStartTime
End Property

Public Property Get SystemId() As String
    SystemId = vSystemId
End Property
'
Public Property Get HardwareBins() As Collection
    Set HardwareBins = objHwBins
End Property
'
Public Property Get SoftwareBins() As Collection
    Set SoftwareBins = objSwBins
End Property


'End Property

Friend Sub Init(ByVal file_name As String)
  If file_name = "" Then Err.Raise 123, , "File name not set"
  'Init object state here
  'vLogfileName = file_name
  
  'Check file exist on version 1.0.7
    If Dir(file_name) = "" Then
        vDescription = "File Not found!!!"
        Exit Sub
    End If
    '-----------------------------
    
  vFileContains = FileToString(file_name)
  
    'Get Seq from file Name
    Dim vOnlyFileName As String
    
    
    
    
    
    vOnlyFileName = get_only_fileName(file_name)
    vLogfileName = vOnlyFileName
    
  Dim vFileSplited() As String
  vFileSplited = Split(vOnlyFileName, "_")
  For i = 1 To IIf(UBound(vFileSplited) = 1, 1, UBound(vFileSplited) - 1)
    vSeq = vSeq + "_" + vFileSplited(i)
  Next
  vSeq = Replace(vSeq, ".txt", "")
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
        
        If textContain(vLineStr, "DUTs Tested") Or textContain(vLineStr, "DUTs Passed") Then vTestedLineStart = True
        If textContain(vLineStr, "Sfwr") Then vTestedLineStart = False: vSwBinLineStart = True
        If textContain(vLineStr, "Hdwr") Then vSwBinLineStart = False: vHwBinLineStart = True
'
        'Tested,Passed,Failed and Yield
        If vTestedLineStart Then
            getTestedPassedFailed vLineStr
            GoTo NextLine
        End If

        'Software Bin Site information
        If vSwBinLineStart Then
            getSwBin vLineStr
            GoTo NextLine
        End If
        
        'Hardware Bin Site information
        If vHwBinLineStart Then
            getHwBin vLineStr
            GoTo NextLine
        End If
        
        
        'Test Program and Tested
        vKey = "Program Name:"
        If textContain(vLineStr, vKey) Then getTestProgram vLineStr
        
        'Tested
        vKey = "UNITS TESTED"
        If textContain(vLineStr, vKey) Then getTested vLineStr
        
        'Tested
        vKey = "UNITS PASSED"
        If textContain(vLineStr, vKey) Then getPassed vLineStr
        
        'Assy,Lot and Seq
        vKey = "Report for Lot:"
        If textContain(vLineStr, vKey) Then getLot vLineStr

        'Operator
        vKey = "Operator:"
        If textContain(vLineStr, vKey) Then getOperator vLineStr
        
        'Temperature
        vKey = "Temperature:"
        If textContain(vLineStr, vKey) Then getTemperature vLineStr
        
        'Tester
        vKey = "Data Collected by Station:"
        If textContain(vLineStr, vKey) Then getTester vLineStr
'
'
'        'Computer , Bin#
'        vKey = "Computer        :"
'        If textContain(vLineStr, vKey) Then getComputer vLineStr
'
        'Handler , Yield
        vKey = "Handler ID:"
        If textContain(vLineStr, vKey) Then getHandler vLineStr
'
'        'Autocorrelation , Next Serial
'        vKey = "Autocorrelation :"
'        If textContain(vLineStr, vKey) Then getAutocorrelation vLineStr
'
        'Start/Stop Date
        vKey = "Data Collection Start Date:"
        If textContain(vLineStr, vKey) Then getStartDate vLineStr
        
        vKey = "Data Collection Stop Date:"
        If textContain(vLineStr, vKey) Then getStopDate vLineStr
        
        vKey = "SYSTEM ID"
        If textContain(vLineStr, vKey) Then getSystemId vLineStr
        
        vKey = "Test Name:"
        If textContain(vLineStr, vKey) Then getTestName vLineStr
        

NextLine:
        
    Next
    
     vCompleted = True
    Exit Sub
HasError:
    vDescription = "Error on line " & i & " : " & Err.Description
    
End Sub

Private Sub getTestProgram(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Program Name:")
    'Get Test Program name
    vTestProgram = vData(1)
End Sub


Private Sub getTested(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, " ")
    
    For i = UBound(vData) - 1 To 5 Step -1
        'MsgBox (i)
        If vData(i) <> "" Then
            If Val(vData(i)) >= 0 Then vTested = Val(vData(i)): Exit For
            
        End If
    Next

End Sub

Private Sub getPassed(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, " ")
    
    vYield = Val(Replace(vData(UBound(vData)), "%", ""))
    For i = UBound(vData) - 1 To 5 Step -1
        'MsgBox (i)
        If vData(i) <> "" Then
            If Val(vData(i)) >= 0 Then vPassed = Val(vData(i)): Exit For
        End If
    Next
   'vPassed = 0
End Sub
'
Private Sub getLot(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Report for Lot:")

    'get Assy,Lot and Seq
    Dim vLotArry() As String
    vLotArry = Split(vData(1), "_")
    'vAssy = Trim(vLotArry(0))
    vLot = Trim(vLotArry(0))
    'vSeq = Trim(Mid(vLotArry(1), 1, 2))
End Sub

Private Sub getOperator(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Operator:")

    'get Operator
    vOperator = Trim(vData(1))
End Sub

Private Sub getTestName(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Test Name:")

    'get Operator
    vTestName = Trim(vData(1))
End Sub


Private Sub getTemperature(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Temperature:")

    'get Temperature
    vTemperature = Trim(vData(1))
End Sub

Private Sub getTester(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Data Collected by Station:")

    'get Tester
    vTester = Trim(vData(1))
End Sub



'
'Private Sub getComputer(vLineData As String)
'    Dim vData() As String
'    vData = Split(vLineData, " ")
'
'    'get Computer
'    Dim ixColon As Integer
'    ixColon = getColonPosition(vLineData)
'    vComputer = vData(ixColon + 1)
'    'get Bin number
'    ixColon = getColonPosition(vLineData, ixColon + 1)
'    vBinNumber = vData(ixColon + 1)
'End Sub
'
Private Sub getHandler(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Handler ID:")
    vHandler = vData(1)
End Sub

Private Sub getSystemId(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "SYSTEM ID")
    vSystemId = Trim(vData(1))
End Sub

Private Sub getTestDate(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Date:")

    If Not IsDate(Trim(vData(1))) Then
        Exit Sub
    End If

    'get Date
    vStartDate = Trim(vData(1))

End Sub

Private Sub getStartDate(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Data Collection Start Date:")

    If Not IsDate(Trim(vData(1))) Then
        Exit Sub
    End If

    'get Date
    vStartDate = Trim(vData(1))

End Sub

Private Sub getStopDate(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Data Collection Stop Date:")

    If Not IsDate(Trim(vData(1))) Then
        Exit Sub
    End If

    'get Date
    vStopDate = Trim(vData(1))

End Sub


Private Sub getTestTime(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Time:")

    If Not IsDate(Trim(vData(1))) Then
        Exit Sub
    End If

    'get Time
    vStartTime = Trim(vData(1))

End Sub


Private Sub getTestedPassedFailed(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "|")
    If vData(0) = "" Or Val(vData(0)) <= 0 Then Exit Sub
    vTested = Val(vData(0))
    vPassed = Val(vData(1))
    vFailed = Val(vData(2))
    vYield = Val(Replace(vData(3), "%", ""))
    
End Sub
'
'
'Private Sub getSwAndHwBin(vLineData As String)
'    Dim vData() As String
'    vData = Split(vLineData, " ")
'    If vData(0) = "" Then Exit Sub
'
'    'Software bin
'    Dim vSwBinNumber As String
'    Dim vSwDescription As String
'    Dim vSwTotal As Double
'    Dim vSwPercent As String
'    Dim ixPercent As Integer
'    'Get Bin number
'    vSwBinNumber = Replace(Replace(vData(0), "[", ""), "]", "")
'    'get Bin Description
'    ixPercent = getPercentPosition(vLineData, 1)
'    For i = 1 To 10
'        If vData(i) <> "" Then vSwDescription = vSwDescription & " " & vData(i)
'    Next
'    vSwDescription = Trim(vSwDescription)
'
'    'Get Total
'    For i = 11 To ixPercent - 2
'        If vData(i) <> "" Then
'            If Val(vData(i)) >= 0 Then vSwTotal = Val(vData(i)): Exit For
'        End If
'    Next
'    'Get Percent
'    vSwPercent = vData(ixPercent - 1)
'    'Add SW Bin in to collection
'    Dim objBin As New clsBin
'    With objBin
'        .Name = vSwBinNumber
'        .Description = vSwDescription
'        .Total = vSwTotal
'        .Yield = vSwPercent
'    End With
'    objSwBins.Add objBin
'    '---------------------------
'
'    'Hardware Bin
'    Dim vHwBinNumber As String
'    Dim vHwTotal As Double
'    Dim vHwYield As Double
'    If InstrCount(vLineData, "%") = 2 Then
'        'MsgBox ""
'        For i = ixPercent + 1 To UBound(vData)
'            If vData(i) <> "" Then
'                If Val(vData(i)) >= 0 Then vHwBinNumber = vData(i): ixPercent = i: Exit For
'            End If
'        Next
'
'        For i = ixPercent + 1 To UBound(vData)
'            If vData(i) <> "" Then
'                If Val(vData(i)) >= 0 Then vHwTotal = Val(vData(i)):  Exit For
'            End If
'        Next
'
'        ixPercent = getPercentPosition(vLineData, ixPercent)
'        vHwYield = vData(ixPercent - 1)
'
'        'Add HW Bin in to collection
'        Dim objHWBin As New clsBin
'        With objHWBin
'            .Name = vHwBinNumber
'            .Description = ""
'            .Total = vHwTotal
'            .Yield = vHwYield
'        End With
'        objHwBins.Add objHWBin
'    End If
'
'End Sub
'
Private Sub getSwBin(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, " ")
    
    If UBound(vData) < 6 Then Exit Sub
    
    Dim vBinNumber As String
    vBinNumber = IIf(vData(4) <> "", vData(4), vData(5))
    
    Dim ixDescStart As Integer
    ixDescStart = IIf(Len(vBinNumber) = 1, 12, 11)
    
    
    If Not IsNumeric(vBinNumber) Then Exit Sub
    
    Dim vBinDescription As String
    Dim vTotal As Double
    Dim vYield As Double
    
    For i = ixDescStart To UBound(vData)
        If IsNumeric(vData(i)) Then
            vTotal = Val(vData(i))
            Exit For
        End If
        vBinDescription = vBinDescription & " " & vData(i)
    Next
    vBinDescription = Trim(vBinDescription)
    vYield = Replace(vData(UBound(vData)), "%", "")
'        'Add SW Bin in to collection
        Dim objBin As New clsBin
        With objBin
            .Name = vBinNumber
            .Description = vBinDescription
            .Total = vTotal
            .Yield = vYield
        End With
        objSwBins.Add objBin
End Sub

Private Sub getHwBin(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, " ")
    
    If UBound(vData) < 6 Then Exit Sub
    
    Dim vBinNumber As String
    vBinNumber = IIf(vData(4) <> "", vData(4), vData(5))
    
    Dim ixDescStart As Integer
    ixDescStart = IIf(Len(vBinNumber) = 1, 12, 11)
    
    
    If Not IsNumeric(vBinNumber) Then Exit Sub
    
    Dim vBinDescription As String
    Dim ixTotal As Integer
    For i = ixDescStart To UBound(vData)
        If IsNumeric(vData(i)) Then
            ixTotal = i
            Exit For
        End If
        vBinDescription = vBinDescription & " " & vData(i)
    Next
    vBinDescription = Trim(vBinDescription)
    
'        'Add SW Bin in to collection
        Dim objBin As New clsBin
        With objBin
            .Name = vBinNumber
            .Description = vBinDescription
            .Total = vData(ixTotal)
            .Yield = Replace(vData(UBound(vData)), "%", "")
        End With
        objHwBins.Add objBin
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









