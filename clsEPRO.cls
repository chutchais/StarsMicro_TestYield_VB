VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEPRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim vFileContains As String
Dim vLogfileName As String
Dim vSummaryName As String
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
Dim vSystemId As String
'
Dim objIBBins As New Collection
Dim objDBBins As New Collection
Dim objDataBins As New Collection

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

Public Property Get IBs() As Collection
    Set IBs = objIBBins
End Property

Public Property Get DBs() As Collection
    Set DBs = objDBBins
End Property

Public Property Get HardwareBins() As Collection
    Set HardwareBins = objIBBins
End Property

Public Property Get SoftwareBins() As Collection
    Set SoftwareBins = objDBBins
End Property

Public Property Get DataBins() As Collection
    Set DataBins = objDataBins
End Property

Public Property Get SummaryName() As String
    SummaryName = vSummaryName
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

'End Property

Friend Sub Init(ByVal file_name As String)
  If file_name = "" Then Err.Raise 123, , "File name not set"
  'Check file exist on version 1.0.7
    If Dir(file_name) = "" Then
        vDescription = "File Not found!!!"
        Exit Sub
    End If
    '-----------------------------
  'Init object state here
  vFileContains = FileToString(file_name)
  vLogfileName = get_only_fileName(file_name)
  
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
        
        'Summary Name
        vKey = "Summary Name:"
        If textContain(vLineStr, vKey) Then getSummaryName vLineStr
        
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
        vKey = "Lot Number:"
        If textContain(vLineStr, vKey) Then getLot vLineStr

        'Operator
        vKey = "Operator:"
        If textContain(vLineStr, vKey) Then getOperator vLineStr
        
        'Temperature
        vKey = "Temperature:"
        If textContain(vLineStr, vKey) Then getTemperature vLineStr
        
        'Tester
        vKey = "Tester #:"
        If textContain(vLineStr, vKey) Then getTester vLineStr

        'Handler , Yield
        vKey = "Handler ID:"
        If textContain(vLineStr, vKey) Then getHandler vLineStr

        

        'Start/Stop Date
        vKey = "Date:"
        If textContain(vLineStr, vKey) Then getTestDate vLineStr
        
        vKey = "Time:"
        If textContain(vLineStr, vKey) Then getTestTime vLineStr
        
        vKey = "SYSTEM ID"
        If textContain(vLineStr, vKey) Then getSystemId vLineStr
        
        'IB
        vKey = "IB "
        If textContain(vLineStr, vKey) Then getIB vLineStr
        
        'DB
        vKey = "DB "
        If textContain(vLineStr, vKey) Then getDB vLineStr
        
        'Data Bins
        vKey = "Data Bin"
        If textContain(vLineStr, vKey) Then getDataBin vLineStr
        
        'Data Bins
        vKey = "Bin "
        If textContain(vLineStr, vKey) Then getDataBin vLineStr
        

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

Private Sub getSummaryName(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Summary Name:")
    'Get SummaryName
    vSummaryName = vData(1)
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
    vData = Split(vLineData, "Lot Number:")

    'get Assy,Lot and Seq
    Dim vLotArry() As String
    vLotArry = Split(vData(1), "_")
    vAssy = Trim(vLotArry(0))
    If UBound(vLotArry) = 0 Then Exit Sub
    vLot = Trim(vLotArry(1))
    vSeq = Trim(Mid(vLotArry(2), 1))
End Sub

Private Sub getOperator(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Operator:")

    'get Operator
    vOperator = Trim(vData(1))
End Sub

Private Sub getTemperature(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Temperature:")

    'get Temperature
    vTemperature = Trim(vData(1))
End Sub

Private Sub getTester(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Tester #:")

    'get Tester
    vTester = Trim(vData(1))
End Sub



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

Private Sub getTestTime(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, "Time:")

    If Not IsDate(Trim(vData(1))) Then
        Exit Sub
    End If

    'get Time
    vStartTime = Trim(vData(1))

End Sub

Private Sub getIB(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, " ")
    If vData(0) = "" Then Exit Sub
    
    Dim vIBNumber As String
    Dim vIBTotal As Double
    Dim vIBYield As Double
    vIBNumber = vData(2)
    For i = 3 To UBound(vData)
        If vData(i) <> "" And Val(vData(i)) >= 0 Then
            vIBTotal = Val(vData(i))
            Exit For
        End If
    Next
    
    vIBYield = Val(Replace(vData(UBound(vData)), "%", ""))
    
    'Add SW Bin in to collection
    Dim objBin As New clsBin
    With objBin
        .Name = vIBNumber
        .Description = ""
        .Total = vIBTotal
        .Yield = vIBYield
    End With
    objIBBins.Add objBin
    '---------------------------
End Sub


Private Sub getDB(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, " ")
    If vData(0) = "" Then Exit Sub
    
    Dim vIBNumber As String
    Dim vIBTotal As Double
    Dim vIBYield As Double
    vIBNumber = vData(2)
    For i = 3 To UBound(vData)
        If vData(i) <> "" And Val(vData(i)) >= 0 Then
            vIBTotal = Val(vData(i))
            Exit For
        End If
    Next
    
    vIBYield = Val(Replace(vData(UBound(vData)), "%", ""))
    
    'Add SW Bin in to collection
    Dim objBin As New clsBin
    With objBin
        .Name = vIBNumber
        .Description = ""
        .Total = vIBTotal
        .Yield = vIBYield
    End With
    objDBBins.Add objBin
    '---------------------------
End Sub

Private Sub getDataBin(vLineData As String)
    Dim vData() As String
    vData = Split(vLineData, " ")
    
    Dim vIBNumber As String
    Dim vDescription As String
    Dim vIBTotal As Double
    Dim vIBYield As Double
    
    
    If vData(0) = "Bin" Then
        vIBNumber = Replace(vData(2), ":", "")
        For i = 3 To UBound(vData)
            vDescription = vDescription & " " & vData(i)
        Next
        vDescription = Trim(vDescription)
        
    Else
    
        If vData(4) = "" Then Exit Sub
        
        
        vIBNumber = Replace(vData(4), ":", "")
        For i = 5 To UBound(vData)
            vDescription = vDescription & " " & vData(i)
        Next
        vDescription = Trim(vDescription)
    End If
    
    'Add SW Bin in to collection
    Dim objBin As New clsBin
    With objBin
        .Name = vIBNumber
        .Description = vDescription
        .Total = 0
        .Yield = 0
    End With
    objDataBins.Add objBin
    '---------------------------
End Sub


Public Function get_report_header(Optional delimiter = ",") As String
 Dim vColName As String
    vColName = "FileName,LOT NO,Lot ID,Seq,Operator,Temperature,Tester,Handler,Summary Address,Date,Time,SYSTEM ID,PROGRAM NAME,Total,Total Pass,Total Fail,Yield,"
    
    For i = 1 To 6
        vColName = vColName & "IB" & Trim(Str(i)) & ","
    Next
    For i = 1 To 32
        vColName = vColName & "DB" & Trim(Str(i)) & ","
    Next
    
    For i = 1 To 32
        vColName = vColName & "BIN" & Trim(Str(i)) & ","
    Next
    get_report_header = vColName
 
 End Function
 
 Public Function get_all_detail(Optional delimiter = ",") As String
    Dim vOutput As New Collection
    Dim oUnitTest As clsBin
    With vOutput
        .Add vLogfileName
        .Add vLot
        .Add vLot
        .Add vSeq
        .Add vOperator
        .Add vTemperature
        .Add vTester
        .Add vHandler
        .Add vSummaryName
        .Add vStartDate
        .Add vStartTime
        .Add vSystemId
        .Add vTestProgram
        .Add vTested
        .Add vPassed
        .Add (vTested - vPassed)
        .Add Yield
        
        
        'Interface Bin (6 items) -- IBX
        For i = 1 To 6
            'UNITS TESTED
            Set oUnitTest = getCollection(Trim(Str(i)), objIBBins)
            If oUnitTest Is Nothing Then
                .Add ""
            Else
                .Add oUnitTest.Total
            End If
        Next
        
        'D Bin (32 items) -- DBX
        For i = 1 To 32
            'UNITS TESTED
            Set oUnitTest = getCollection(Trim(Str(i)), objDBBins)
            If oUnitTest Is Nothing Then
                .Add ""
            Else
                .Add oUnitTest.Total
            End If
        Next
        
        'D Bin (32 items) -- BINX
        For i = 1 To 32
            'UNITS TESTED
            Set oUnitTest = getCollection(Trim(Str(i)), objDataBins)
            If oUnitTest Is Nothing Then
                .Add ""
            Else
                .Add oUnitTest.Description
                
            End If
        Next
    
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







