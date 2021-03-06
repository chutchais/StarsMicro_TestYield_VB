VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMAV"
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
Dim vHandler As String
Dim vStartDate As String
Dim vStartTime As String
Dim vStopDate As String
Dim vSystemId As String
Dim vTestFlow As String
Dim vTestCount As String
Dim vDeviceName As String
Dim vProgramName As String
Dim vProgramRev As String
Dim vLotSize As String
Dim vTestType As String

'
Dim objUnitBins As New Collection
Dim objFailureBins As New Collection
Dim objIBBins As New Collection
Dim objDBBins As New Collection

Dim oIBins As New Collection 'Interface bin (%INTERFACE BINS)
Dim oDBins As New Collection 'Databin bin (%DATA BINS)
'
'COntrol
Dim vSiteTotal As Integer
'Dim vSwBinLineStart As Boolean
'Dim vSwSiteLineStart As Boolean
'Dim vSwSiteTotal As Integer

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

Public Property Get TestCount() As String
    TestCount = vTestCount
End Property

Public Property Get TestFlow() As String
    TestFlow = vTestFlow
End Property

Public Property Get DeviceName() As String
    DeviceName = vDeviceName
End Property

Public Property Get ProgramName() As String
    ProgramName = vProgramName
End Property

Public Property Get LotSize() As String
    LotSize = vLotSize
End Property

Public Property Get ProgramRev() As String
    ProgramRev = vProgramRev
End Property

Public Property Get Temperature() As String
    Temperature = vTemperature
End Property

Public Property Get Tester() As String
    Tester = vTester
End Property

Public Property Get TestType() As String
    TestType = vTestType
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
On Error GoTo HasError
    Yield = Format((vPassed / vTested) * 100, "#0.00")
    Exit Property
HasError:
    Yield = 0
End Property

Public Property Get StartDate() As String
    StartDate = vStartDate
End Property

Public Property Get StopDate() As String
    StopDate = vStopDate
End Property

Public Property Get StartTime() As String
    StartTime = vStartTime
End Property

Public Property Get SystemId() As String
    SystemId = vSystemId
End Property
'
Public Property Get HardwareBins() As Collection
    Set HardwareBins = objIBBins
End Property

Public Property Get SoftwareBins() As Collection
    Set SoftwareBins = objDBBins
End Property

Public Property Get UnitBins() As Collection
    Set UnitBins = objUnitBins
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
  
  'get Assy and Seq
  Dim vArray() As String
  vArray = Split(file_name, "\")
  vAssy = Split(vArray(UBound(vArray)), "_")(0)

  
  'Get Seq from file Name
  
  'Get only fileName
  
  Dim vOnlyFileName As String
  vOnlyFileName = get_only_fileName(file_name)
  vLogfileName = vOnlyFileName
  
  Dim vFileSplited() As String
  vFileSplited = Split(vOnlyFileName, "_")
  For i = 2 To UBound(vFileSplited)
    vSeq = vSeq + "_" + vFileSplited(i)
  Next
  vSeq = Replace(vSeq, ".txt", "")
  vSeq = Mid(vSeq, 2)
  '----------------------------------
  vTestType = Mid(vOnlyFileName, 1, 2)
  
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
        
        If textContain(vLineStr, "%INTERFACE BINS") Then
                vInterfaceBinLineStart = True
                GoTo NextLine
        End If
        If textContain(vLineStr, "%DATA BINS") Then
                vInterfaceBinLineStart = False: vSwSiteLineStart = True
                GoTo NextLine
        End If
'
        'Hardware Bin and Software Bin
        If vInterfaceBinLineStart Then
            getInterfaceBin vLineStr
            GoTo NextLine
        End If
'
'        'Software Bin Site information
        If vSwSiteLineStart Then
            getSoftwareBin vLineStr
            GoTo NextLine
        End If
        
        'Check How many SITE
        vKey = "%LOT RESULTS"
        If textContain(vLineStr, vKey) Then
            vSiteTotal = InstrCount(vLineStr, "SITE")
            GoTo NextLine
        End If
        
        
        'Failure Bins
        vKey = " FAIL "
        If textContain(vLineStr, vKey) Then getFailure vLineStr
        
        'IB Bins
        vKey = "IB "
        If textContain(vLineStr, vKey) Then getIB vLineStr
        
        'Hardware Bin
        vKey = "Bin  "
        If textContain(vLineStr, vKey) Then getIB vLineStr
        
        
        
        
        'DB Bins
        vKey = "DB "
        If textContain(vLineStr, vKey) Then getDB vLineStr
        
        'Test Program and Tested
        vKey = "Program Name:"
        If textContain(vLineStr, vKey) Then getTestProgram vLineStr
        
        'Tested
        vKey = "UNITS TESTED"
        If textContain(vLineStr, vKey) Then getTested vLineStr
        
        'Passed
        vKey = "UNITS PASSED"
        If textContain(vLineStr, vKey) Then getPassed vLineStr
        
        'Failed
        vKey = "UNITS FAILED"
        If textContain(vLineStr, vKey) Then getFailed vLineStr
        
        'Assy,Lot and Seq
        vKey = "Lot Number:"
        If textContain(vLineStr, vKey) Then getLot vLineStr

        'Operator
        vKey = "Operator ID:"
        If textContain(vLineStr, vKey) Then getOperator vLineStr
        
        'Temperature
        vKey = "Temperature:"
        If textContain(vLineStr, vKey) Then getTemperature vLineStr
        
        'Tester
        vKey = "Tester ID:"
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
        vKey = "End Time:"
        If textContain(vLineStr, vKey) Then getStopDate vLineStr
        
        vKey = "Start Time:"
        If textContain(vLineStr, vKey) Then getStartDate vLineStr
        
        vKey = "SYSTEM ID"
        If textContain(vLineStr, vKey) Then getSystemId vLineStr
        
        vKey = "Test Flow:"
        If textContain(vLineStr, vKey) Then getTestFlow vLineStr
        
        vKey = "Test Count:"
        If textContain(vLineStr, vKey) Then getTestCount vLineStr
        
        vKey = "Device Name:"
        If textContain(vLineStr, vKey) Then getDeviceName vLineStr
        
        vKey = "Program Name:"
        If textContain(vLineStr, vKey) Then getProgramName vLineStr
        
        vKey = "Format Rev:"
        If textContain(vLineStr, vKey) Then getProgramRev vLineStr
        
        'Unit Bins
        vKey = "UNITS "
        If textContain(vLineStr, vKey) Then getUnit vLineStr
        
        vKey = "Lot Size:"
        If textContain(vLineStr, vKey) Then getLotSize vLineStr
        
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
    
    For i = 2 To UBound(vData) - 1
        'MsgBox (i)
        If vData(i) <> "" Then
            If Val(vData(i)) >= 0 Then vTested = Val(vData(i)): Exit For
            
        End If
    Next

End Sub

Private Sub getFailed(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, " ")
    
    For i = 2 To UBound(vData) - 1
        'MsgBox (i)
        If vData(i) <> "" Then
            If Val(vData(i)) >= 0 Then vFailed = Val(vData(i)): Exit For
            
        End If
    Next

End Sub

Private Sub getPassed(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, " ")
    
    
    For i = 2 To UBound(vData) - 1
        'MsgBox (i)
        If vData(i) <> "" Then
            If Val(vData(i)) >= 0 Then vPassed = Val(vData(i)): Exit For
        End If
    Next
    'vYield = Format((vPassed / vTested) * 100, "#0.00")
   'vPassed = 0
End Sub
'
Private Sub getLot(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Lot Number:")
    vLot = Trim(vData(1))

End Sub

Private Sub getOperator(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Operator ID:")

    'get Operator
    vOperator = Trim(vData(1))
End Sub

Private Sub getTestFlow(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Test Flow:")

    'get Operator
    vTestFlow = Trim(vData(1))
End Sub

Private Sub getTestCount(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Test Count:")

    'get Operator
    vTestCount = Trim(vData(1))
End Sub


Private Sub getTemperature(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Temperature:")

    'get Temperature
    vTemperature = Trim(vData(1))
End Sub

Private Sub getTester(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Tester ID:")

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
    vHandler = Trim(vData(1))
End Sub

Private Sub getSystemId(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "SYSTEM ID")
    vSystemId = Trim(vData(1))
End Sub

Private Sub getDeviceName(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Device Name:")
    vDeviceName = Trim(vData(1))
End Sub

Private Sub getProgramName(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Program Name:")
    vProgramName = Trim(vData(1))
End Sub

Private Sub getLotSize(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Lot Size:")
    vLotSize = Trim(vData(1))
End Sub


Private Sub getProgramRev(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Format Rev:")
    vProgramRev = Trim(vData(1))
End Sub

Private Sub getStartDate(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Start Time:")

'    If Not IsDate(Trim(vData(1))) Then
'        Exit Sub
'    End If

    'get Date
    vStartDate = Trim(vData(1))

End Sub

Private Sub getStopDate(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "End Time:")

'    If Not IsDate(Trim(vData(1))) Then
'        Exit Sub
'    End If

    'get Date
    vStopDate = Trim(vData(1))

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

Private Sub getTestTime(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Time:")

    If Not IsDate(Trim(vData(1))) Then
        Exit Sub
    End If

    'get Time
    vStartTime = Trim(vData(1))

End Sub

Private Sub getUnit(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, " ")
    
    Dim vFaiureName As String
    Dim vTotal As Double
    Dim vYield As Double
    Dim ixYield As Integer
    Dim k As Integer
    ixYield = getPercentPosition(vLineData)
    For i = 0 To IIf(ixYield = 0, 4, ixYield - 1)
        vFaiureName = vFaiureName & " " & vData(i)
        If vData(i) = "" Then k = i + 1: Exit For
    Next
    vFaiureName = Trim(vFaiureName)
    For i = k To IIf(ixYield = 0, 10, ixYield - 1)
        If vData(i) <> "" Then
            If IsNumeric(vData(i)) Then vTotal = Val(vData(i)): k = i: Exit For
        End If
    Next
    vYield = Val(Replace(vData(ixYield), "%", ""))
    
    'Create Fialure Bin
    Dim objBin As New clsBin
    With objBin
        .Name = vFaiureName
        .Description = vFaiureName
        .Total = vTotal
        .Yield = vYield
    End With
    
    'Site data
    Dim vSiteNumber As Integer
    vSiteNumber = 1
    For i = IIf(ixYield = 0, k + 1, ixYield + 1) To UBound(vData)
        If vData(i) <> "" Then
            If IsNumeric(vData(i)) Then
                vTotal = Val(vData(i))
                'Crate Site object
                Dim objSite As New clsSite
                With objSite
                    .Name = Trim(Str(vSiteNumber))
                    .Description = ""
                    .Total = vTotal
                    .Yield = 0
                End With
                'Add Site to failure bin
                objBin.Sites.Add objSite
                Set objSite = Nothing
                vSiteNumber = vSiteNumber + 1
            End If
        End If
    Next
    objUnitBins.Add objBin
 End Sub

Private Sub getFailure(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, " ")
    
    Dim vFaiureName As String
    Dim vTotal As Double
    Dim vYield As Double
    Dim ixYield As Integer
    Dim k As Integer
    ixYield = getPercentPosition(vLineData)
    For i = 0 To ixYield - 1
        vFaiureName = vFaiureName & " " & vData(i)
        If vData(i) = "" Then k = i + 1: Exit For
    Next
    vFaiureName = Trim(vFaiureName)
    For i = k To ixYield - 1
        If vData(i) <> "" Then
            If IsNumeric(vData(i)) Then vTotal = Val(vData(i)): k = i: Exit For
        End If
    Next
    vYield = Val(Replace(vData(ixYield), "%", ""))
    
    'Create Fialure Bin
    Dim objBin As New clsBin
    With objBin
        .Name = vFaiureName
        .Description = ""
        .Total = vTotal
        .Yield = vYield
    End With
    
    'Site data
    Dim vSiteNumber As Integer
    vSiteNumber = 1
    For i = ixYield + 1 To UBound(vData)
        If vData(i) <> "" Then
            If IsNumeric(vData(i)) Then
                vTotal = Val(vData(i))
                'Crate Site object
                Dim objSite As New clsSite
                With objSite
                    .Name = Str(vSiteNumber)
                    .Description = ""
                    .Total = vTotal
                    .Yield = 0
                End With
                'Add Site to failure bin
                objBin.Sites.Add objSite
                Set objSite = Nothing
                vSiteNumber = vSiteNumber + 1
            End If
        End If
    Next
    objFailureBins.Add objBin
 End Sub
 
 Private Sub getIB(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, " ")
    
    Dim vName As String
    Dim vTotal As Double
    Dim vYield As Double
    Dim ixYield As Integer
    Dim k As Integer
    ixYield = getPercentPosition(vLineData)
    
  'comment on Jan 18,2019
  'found many space
'    vName = IIf(vData(1) <> "", vData(1), vData(2))
'    vName = Trim(vName)
'k = IIf(Len(vName) = 1, 3, 2)
    For i = 1 To ixYield - 1
        If vData(i) <> "" Then
            vName = vData(i)
            k = i + 1
            Exit For
        End If
    Next
    
    
    
    
    For i = k To ixYield - 1
        If vData(i) <> "" Then
            If IsNumeric(vData(i)) Then vTotal = Val(vData(i)): k = i: Exit For
        End If
    Next
    vYield = Val(Replace(vData(ixYield), "%", ""))
    
    'Create Fialure Bin
    Dim objBin As New clsBin
    With objBin
        .Name = vName
        .Description = ""
        .Total = vTotal
        .Yield = vYield
    End With
    
    'Site data
    Dim vSiteNumber As Integer
    vSiteNumber = 1
    For i = ixYield + 1 To UBound(vData)
        If vData(i) <> "" Then
            If IsNumeric(vData(i)) Then
                vTotal = Val(vData(i))
                'Crate Site object
                Dim objSite As New clsSite
                With objSite
                    .Name = Str(vSiteNumber)
                    .Description = ""
                    .Total = vTotal
                    .Yield = 0
                End With
                'Add Site to failure bin
                objBin.Sites.Add objSite
                Set objSite = Nothing
                vSiteNumber = vSiteNumber + 1
            End If
        End If
    Next
    objIBBins.Add objBin
 End Sub


 Private Sub getDB(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, " ")
    
    Dim vName As String
    Dim vTotal As Double
    Dim vYield As Double
    Dim ixYield As Integer
    Dim k As Integer
    ixYield = getPercentPosition(vLineData)
    
'    vName = IIf(vData(1) <> "", vData(1), vData(2))
'    vName = Trim(vName)
'    k = IIf(Len(vName) = 1, 3, 2)
    For i = 1 To ixYield - 1
            If vData(i) <> "" Then
                vName = vData(i)
                k = i + 1
                Exit For
            End If
        Next
    
    
    For i = k To ixYield - 1
        If vData(i) <> "" Then
            If IsNumeric(vData(i)) Then vTotal = Val(vData(i)): k = i: Exit For
        End If
    Next
    vYield = Val(Replace(vData(ixYield), "%", ""))
    
    'Create Fialure Bin
    Dim objBin As New clsBin
    With objBin
        .Name = vName
        .Description = ""
        .Total = vTotal
        .Yield = vYield
    End With
    
    'Site data
    Dim vSiteNumber As Integer
    vSiteNumber = 1
    For i = ixYield + 1 To UBound(vData)
        If vData(i) <> "" Then
            If IsNumeric(vData(i)) Then
                vTotal = Val(vData(i))
                'Crate Site object
                Dim objSite As New clsSite
                With objSite
                    .Name = Str(vSiteNumber)
                    .Description = ""
                    .Total = vTotal
                    .Yield = 0
                End With
                'Add Site to failure bin
                objBin.Sites.Add objSite
                vSiteNumber = vSiteNumber + 1
            End If
        End If
    Next
    objDBBins.Add objBin
 End Sub
 
 Public Function get_report_header(Optional delimiter = ",") As String
 Dim vColName As String
    vColName = "FileName,Lot no,Seq,test Flow,Test Type,Device Name,Lot Size,test Count,Temperature,Program Name,Program Rev,Tester ID,Handler ID," & _
    "UNITS Tested,UNITS Passed,UNITS Failed,test Yield,"
    
    For i = 1 To 8
        vColName = vColName & "Test_SITE" & Trim(Str(i)) & "," & "Pass_SITE" & Trim(Str(i)) & "," & "Fail_SITE" & Trim(Str(i)) & ","
    Next
    For i = 1 To 8
        vColName = vColName & "BIN" & Trim(Str(i)) & ","
    Next
    For i = 1 To 32
        vColName = vColName & "DB" & Trim(Str(i)) & ","
    Next
    
    For i = 1 To 8
        vColName = vColName & "IBIN" & Trim(Str(i)) & ","
    Next
    
    For i = 1 To 40
        vColName = vColName & "DBIN" & Trim(Str(i)) & ","
    Next
    
    vColName = vColName & "DialogTempCorrection,TempCorrection,start Time,End Time"

    get_report_header = vColName
 
 End Function
 
 Public Function get_all_detail(Optional delimiter = ",") As String
    Dim vOutput As New Collection
    With vOutput
        .Add vLogfileName
        .Add vLot
        .Add vSeq
        .Add vTestFlow
        .Add vTestType
        .Add vDeviceName
        .Add vLotSize
        .Add vTestCount
        .Add vTemperature
        .Add vProgramName
        .Add vProgramRev
        .Add vTester
        .Add vHandler
        .Add vTested
        .Add vPassed
        .Add vFailed
        .Add Yield
        'Loop for 8 sites
        Dim oUnitTest As clsBin
        Dim oUnitPass As clsBin
        Dim oUnitFail As clsBin
        Dim oSite As Object
        For i = 1 To 8
            'UNITS TESTED
            Set oUnitTest = getCollection("UNITS TESTED", objUnitBins)
            Set oSite = getSite(Str(i), oUnitTest)
            If oSite Is Nothing Then
                .Add 0
            Else
                .Add oSite.Total
            End If
            'UNITS PASSED
            Set oUnitPass = getCollection("UNITS PASSED", objUnitBins)
            Set oSite = getSite(Str(i), oUnitPass)
            If oSite Is Nothing Then
                .Add 0
            Else
                .Add oSite.Total
            End If
            'UNITS FAILED
            Set oUnitFail = getCollection("UNITS FAILED", objUnitBins)
            Set oSite = getSite(Str(i), oUnitFail)
            If oSite Is Nothing Then
                .Add 0
            Else
                .Add oSite.Total
            End If
        Next
        
        'Interface Bin (8 items) -- BINXX
        For i = 1 To 8
            'UNITS TESTED
            Set oUnitTest = getCollection(Trim(Str(i)), objIBBins)
            If oUnitTest Is Nothing Then
                .Add ""
            Else
                .Add oUnitTest.Total
            End If
        Next
        
        'Data Bin (40 items) (DBXX)
        For i = 1 To 32
            'UNITS TESTED
            Set oUnitTest = getCollection(Trim(Str(i)), objDBBins)
            If oUnitTest Is Nothing Then
                .Add ""
            Else
                .Add oUnitTest.Total
            End If
        Next
        
        
        'Interface Bin (8 items) (IBINXX)
        For i = 1 To 8
            'UNITS TESTED
            Set oUnitTest = getCollection(Format(Str(i), "00"), oIBins)
            If oUnitTest Is Nothing Then
                .Add ""
            Else
                .Add oUnitTest.Description
            End If
        Next
        
        
        'Data Bin (32 items) (DBINXX)
        For i = 1 To 40
            'UNITS TESTED
            Set oUnitTest = getCollection(Trim(Str(i)), oDBins)
            If oUnitTest Is Nothing Then
                .Add ""
            Else
                .Add oUnitTest.Description
            End If
        Next
        
        'DialogTempCorrection
        .Add ""
        'TempCorrection
        .Add ""
        'Start Time
        .Add vStartDate
        'End Time
        .Add vStopDate
    End With
    get_all_detail = merge_collection(vOutput, delimiter)
 End Function
 
 Function merge_collection(col As Collection, Optional delimeter = ",") As String
    Dim vOut As String
    For Each c In col
        vOut = vOut & delimeter & Replace(c, delimeter, " ")
    Next
    vOut = Mid(vOut, 2)
    merge_collection = vOut
 End Function
 
 Private Sub getInterfaceBin(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, ":")
    If vData(0) = "" Then Exit Sub
    
    '    'Add SW Bin in to collection
    Dim objBin As New clsBin
    With objBin
        .Name = Split(vData(0), " ")(1)
        .Description = Trim(vData(1))
        .Total = 0
        .Yield = 0
    End With
    oIBins.Add objBin
'    '---------------------------

 End Sub
 
  Private Sub getSoftwareBin(vLineData As String)
    Dim vData() As String
    
    If Not (textContain(vLineData, "BIN") Or textContain(vLineData, "Bin")) Then
        Exit Sub
    End If
    
    If Not textContain(vLineData, ":") Then
        Exit Sub
    End If
    
    
    
    vData = Split(vLineData, ":")
    If vData(0) = "" Then Exit Sub
    Dim vName As String
    vName = Val(Split(vData(0), " ")(UBound(Split(vData(0), " "))))
    '    'Add SW Bin in to collection
    Dim objBin As New clsBin
    With objBin
        .Name = vName
        .Description = Trim(vData(1))
        .Total = 0
        .Yield = 0
    End With
    oDBins.Add objBin
'    '---------------------------

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


'Private Sub getSwSite(vLineData As String)
'    Dim vData() As String
'    vData = Split(vLineData, " ")
'    If vLineData Like "*Site*" Then
'        vSwSiteTotal = InstrCount(vLineData, "Site")
'    End If
'    If vData(0) = "" Then Exit Sub
'
'    Dim vSwSiteNumber As String
'    Dim vSwBinNumber As String
'    Dim vSiteTotal As Double
'    Dim vSiteYield As Double
'    Dim ixPercent As Integer
'    Dim ixStart As Integer
'    ixStart = 1
'    vSwBinNumber = Replace(Replace(vData(0), "[", ""), "]", "")
'    For i = 1 To vSwSiteTotal
'        vSwSiteNumber = i
'        ixPercent = getPercentPosition(vLineData, ixPercent + 1)
'        'get Total
'        For x = ixStart To ixPercent - 2
'            If vData(x) <> "" Then
'                If Val(vData(x)) >= 0 Then vSiteTotal = Val(vData(x)):  Exit For
'            End If
'        Next
'        ixStart = ixPercent + 1
'        'get Yield
'        vSiteYield = vData(ixPercent - 1)
'        'Add SW Bin in to collection
'        Dim objSite As New clsSite
'        With objSite
'            .Name = vSwSiteNumber
'            .Description = ""
'            .Total = vSiteTotal
'            .Yield = vSiteYield
'        End With
'        'Find SW bin in collection and add site to BIN
'        Dim objCurrentBin As New clsBin
'        Set objCurrentBin = TryGet(vSwBinNumber, objSwBins)
'        objCurrentBin.Sites.Add objSite
'        Set objSite = Nothing
'        Set objCurrentBin = Nothing
'        '---------------------------
'    Next
'
'
'End Sub


'Common function for Class

Private Function getCollection(key As String, col As Collection) As clsBin
  On Error GoTo errhandler
  For Each c In col
        If c.Name = key Then
            Set getCollection = c
            Exit For
        End If
  Next
  Exit Function
errhandler:
  Set getCollection = Nothing
End Function

Private Function getSite(siteName As String, col As clsBin) As Object
  On Error GoTo errhandler
  For Each c In col.Sites
        If c.Name = Trim(siteName) Then
            Set getSite = c
            Exit For
        End If
  Next
  Exit Function
errhandler:
  Set getSite = Nothing
End Function

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

