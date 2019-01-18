VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Tester Yield Report"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17190
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   17190
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7170
      Left            =   90
      TabIndex        =   10
      Top             =   1350
      Width           =   17010
      _ExtentX        =   30004
      _ExtentY        =   12647
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Summary Report"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "File list"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstFile"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "QA Summarize"
         Height          =   1275
         Left            =   135
         TabIndex        =   15
         Top             =   1845
         Width           =   16620
         Begin MSFlexGridLib.MSFlexGrid fGridQa 
            Height          =   735
            Left            =   135
            TabIndex        =   17
            Top             =   360
            Width           =   16440
            _ExtentX        =   28998
            _ExtentY        =   1296
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "FT Summarize"
         Height          =   1275
         Left            =   135
         TabIndex        =   14
         Top             =   495
         Width           =   16620
         Begin MSFlexGridLib.MSFlexGrid fGridFT 
            Height          =   735
            Left            =   135
            TabIndex        =   16
            Top             =   360
            Width           =   16440
            _ExtentX        =   28998
            _ExtentY        =   1296
            _Version        =   393216
         End
      End
      Begin VB.ListBox lstFile 
         Height          =   6495
         Left            =   -74775
         TabIndex        =   12
         Top             =   495
         Width           =   5460
      End
      Begin VB.Frame Frame1 
         Caption         =   "Summary Table"
         Height          =   3885
         Left            =   135
         TabIndex        =   11
         Top             =   3195
         Width           =   16665
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   3525
            Left            =   135
            TabIndex        =   13
            Top             =   270
            Width           =   16440
            _ExtentX        =   28998
            _ExtentY        =   6218
            _Version        =   393216
         End
      End
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate Report"
      Height          =   405
      Left            =   6975
      TabIndex        =   8
      Top             =   90
      Width           =   1455
   End
   Begin VB.TextBox txtFolder 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1170
      TabIndex        =   7
      Top             =   540
      Width           =   9780
   End
   Begin VB.TextBox txtLotNumber 
      Height          =   375
      Left            =   4365
      TabIndex        =   6
      Top             =   90
      Width           =   2580
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export Excel"
      Enabled         =   0   'False
      Height          =   495
      Left            =   14400
      TabIndex        =   4
      Top             =   495
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cbTester 
      Height          =   315
      ItemData        =   "frmMain.frx":0038
      Left            =   1185
      List            =   "frmMain.frx":0048
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13905
      Top             =   4590
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Select Folder"
      Height          =   495
      Left            =   11025
      TabIndex        =   0
      Top             =   450
      Width           =   1455
   End
   Begin VB.Label lblFilesCount 
      Caption         =   "0 File(s)"
      Height          =   255
      Left            =   1170
      TabIndex        =   9
      Top             =   945
      Width           =   3195
   End
   Begin VB.Label Label3 
      Caption         =   "STARs Lot no :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      TabIndex        =   5
      Top             =   180
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Tester :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Folder name :"
      Height          =   255
      Left            =   90
      TabIndex        =   1
      Top             =   630
      Width           =   1035
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim objIni As clsIniFile
Dim objEPRO As clsEPRO
Dim objETS As clsETS
Dim objMAV As clsMAV
Dim objTMT As clsTMT

Dim objFiles As New Collection
Dim colFiles As New Collection
Dim colLots As New Collection

'Tester Log file and Ext
Dim vCurrentFolder As String
Dim vCurrentFileExt As String

Private Sub cbTester_Change()
    cmdExport.Enabled = False
    
End Sub

Private Sub cbTester_Click()

    Me.MousePointer = 11
    clearContains
    
    
    vCurrentFolder = getSectionString(cbTester.Text, "path")
    vCurrentFileExt = getSectionString(cbTester.Text, "file ext")
    txtFolder.Text = vCurrentFolder
    showFileCountInFolder vCurrentFolder, vCurrentFileExt
    Me.MousePointer = 0
End Sub

Sub showFileCountInFolder(folderName As String, Optional extFile As String = "*.*")
    Dim oFs As New FileSystemObject
    If oFs.FolderExists(folderName) Then
        lblFilesCount.Caption = "Found " & CountFiles(folderName, extFile) & " File(s)"
    Else
        lblFilesCount.Caption = "Folder or File doesn't exist"
    End If
    
End Sub


Sub clearContains()
       
    Set colFiles = New Collection
    lstFile.Clear
    initial_Grid_Summary
    txtLotNumber.Text = ""
    cmdExport.Enabled = False
    'lblFileName.Caption = ""
    txtFolder.Text = ""
'    lblTested.Caption = "0"
'    lblPassed.Caption = "0"
'    lblFailed.Caption = "0"
'    lblYield.Caption = "0"
End Sub

Private Sub cmdBrowse_Click()
    If vCurrentFileExt = "" Then
        MsgBox "Please select Tester", vbCritical, "Select tester"
        Exit Sub
    End If
     
     Me.MousePointer = 11
    clearContains
    txtFolder.Text = BrowseForFolder(hwnd, "Please select a Server folder.")
    showFileCountInFolder txtFolder.Text, vCurrentFileExt
    Me.MousePointer = 0
End Sub

Private Sub cmdExport_Click()
    
Select Case cbTester.Text
        Case "EPRO": EPRO_Export_Excel objEPRO
        Case "ETS": ETS_Export_Excel objETS
        Case "MAV": MAV_Export_Excel objMAV
        Case "TMT": TMT_Export_Excel objTMT
    End Select
    


End Sub


Sub process_files(LogfilePath As String, _
                    FileExt As String, _
                    Optional lotNumber As String, _
                    Optional filterByFileName As Boolean = True)
Dim oFs As New FileSystemObject
Dim sAns() As String
Dim oFolder As Folder
Dim oFile As File
Dim lElement As Long

ReDim sAns(0) As String
If oFs.FolderExists(LogfilePath) Then
    Set oFolder = oFs.GetFolder(LogfilePath)
 
    For Each oFile In oFolder.Files
'      lElement = IIf(sAns(0) = "", 0, lElement + 1)
'      ReDim Preserve sAns(lElement) As String
'      sAns(lElement) = oFile.Name
        If textContain(oFile.Name, FileExt) Then
            If filterByFileName Then
                'Filter by File naming
                If textContain(oFile.Name, lotNumber) Then
                    objFiles.Add (oFile)
                End If
            Else
                'Filter by file Contain
                '1)File to Object
                '2)Check Obj.Lot is match
            End If
        End If
    Next
    Debug.Print objFiles.Count & " files"
End If

errhandler:
    Set oFs = Nothing
    Set oFolder = Nothing
    Set oFile = Nothing
End Sub

'File Count
Sub filterFileByLot(lotNumber As String)
    lstFile.Clear
    Set colLots = New Collection
    For Each f In colFiles
        If f Like "*" & lotNumber & "*" Then
            lstFile.AddItem f
            colLots.Add f
        End If
        
    Next
    
End Sub

Sub filterFileByLotForEPRO(lotNumber As String)
    lstFile.Clear
    
    Dim vObjEPRO As clsEPRO
    Set colLots = New Collection
    
    For Each f In colFiles
        'load file to EPRO object
        Set vObjEPRO = New clsEPRO
            vObjEPRO.Init vCurrentFolder & "\" & f
        If vObjEPRO.Lot Like "*" & lotNumber & "*" Then
            lstFile.AddItem f
            colLots.Add f
        End If
        
    Next
    
End Sub

Function CountFiles(strDirectory As String, Optional strExt As String = "*.*") As Double
'   then count only files of that type, otherwise return a count of all files.
    Dim objFso As Object
    Dim objFiles As Object
    Dim objFile As Object

    'Set Error Handling
    On Error GoTo EarlyExit

    'Create objects to get a count of files in the directory
    Set objFso = CreateObject("Scripting.FileSystemObject")
    Set objFiles = objFso.GetFolder(strDirectory).Files

    'Count files (that match the extension if provided)
    If strExt = "*.*" Then
        CountFiles = objFiles.Count
    Else
        For Each objFile In objFiles
            If UCase(Right(objFile.Path, (Len(objFile.Path) - InStrRev(objFile.Path, ".")))) = UCase(strExt) Then
                colFiles.Add (objFile.Name)
                CountFiles = CountFiles + 1
            End If
        Next objFile
    End If

EarlyExit:
    'Clean up
    On Error Resume Next
    Set objFile = Nothing
    Set objFiles = Nothing
    Set objFso = Nothing
    On Error GoTo 0
End Function

'            If LotNumber <> "" Then
'                'Filter Lot number
'                If textContain(oFile.Name, LotNumber) Then
'                    objFiles.Add (oFile)
'                End If
'            Else
'                'Not filter
'                objFiles.Add (oFile)
'            End If



'Private Sub cmdTest_Click()
'
'    process_files "C:\Users\Chutchai\Documents\8oclock\StartsMicro\projects\Tester Yield\Tester logs\ETS", "txt", _
'                    "TS00182350", True
'
'    Dim vFilter As String
'    Dim vExt As String
'    Select Case cbTester.Text
'        Case "EPRO":
'                vFilter = "Apps (*.sum)|*.sum|All files (*.*)|*.sum"
'                vExt = "sum"
'        Case "ETS":
'                vFilter = "Apps (*.txt)|*.txt|All files (*.*)|*.txt"
'                vExt = "txt"
'        Case "MAV":
'                vFilter = "Apps (*.txt)|*.txt|All files (*.*)|*.txt"
'                vExt = "txt"
'        Case "TMT":
'                vFilter = "Apps (*.lsr)|*.lsr|All files (*.*)|*.lsr"
'                vExt = "lsr"
'    End Select
'
'    With CommonDialog1
'        .Filter = vFilter
'        .DefaultExt = vExt
'        .DialogTitle = "Select File"
'        .ShowOpen
'    End With
'
'    If CommonDialog1.FileName = "" Then Exit Sub
'
'    lblFileName.Caption = CommonDialog1.FileName
'
'    cmdExport.Enabled = False
'
'    Dim objFile As Object
'    Select Case cbTester.Text
'        Case "EPRO":
'            Set objEPRO = New clsEPRO
'            objEPRO.Init CommonDialog1.FileName
'
'            If objEPRO.Completed Then
'                showDetail objEPRO
'                cmdExport.Enabled = True
'            Else
'                MsgBox objEPRO.Description, vbCritical, "Profile Error"
'            End If
'
'        Case "ETS":
'            Set objETS = New clsETS
'            objETS.Init CommonDialog1.FileName
'            showDetail objETS
'
'            If objETS.Completed Then
'                showDetail objETS
'                cmdExport.Enabled = True
'            Else
'                MsgBox objETS.Description, vbCritical, "Profile Error"
'            End If
'
'        Case "MAV":
'
'            Set objMAV = New clsMAV
'            objMAV.Init CommonDialog1.FileName
'            showDetail objMAV
'            If objMAV.Completed Then
'                showDetail objMAV
'                cmdExport.Enabled = True
'            Else
'                MsgBox objMAV.Description, vbCritical, "Profile Error"
'            End If
'
'        Case "TMT":
'
'            Set objTMT = New clsTMT
'            objTMT.Init CommonDialog1.FileName
'            showDetail objTMT
'            If objTMT.Completed Then
'                showDetail objTMT
'                cmdExport.Enabled = True
'            Else
'                MsgBox objTMT.Description, vbCritical, "Profile Error"
'            End If
'    End Select
'
'End Sub

Sub showDetail(vObject As Object)
    With vObject
                lblTested.Caption = Format(.Tested, "###,##0")
                lblPassed.Caption = Format(.Passed, "###,##0")
                lblFailed.Caption = Format(.Failed, "###,##0")
                lblYield.Caption = .Yield
            End With
End Sub

Private Sub cmdGenerate_Click()
Me.MousePointer = 11
If txtLotNumber.Text = "" Then
    Exit Sub
End If

initial_Grid_Summary
initial_Grid_FT
initial_Grid_QA

Dim objFileReport As New Collection

'Only EPRO ,can not using file name to filter Lot number (must read in file content)
    If cbTester.Text <> "EPRO" Then
        filterFileByLot txtLotNumber.Text
    Else
        filterFileByLotForEPRO txtLotNumber.Text
    End If


Select Case cbTester.Text
        Case "EPRO":
        Dim objEPRO As New clsEPRO
            For i = 0 To lstFile.ListCount - 1
                Set objEPRO = New clsEPRO
                objEPRO.Init vCurrentFolder & "\" & lstFile.List(i)
                If objEPRO.Completed Then
                    objFileReport.Add objEPRO
                    'each file
                    add_data_to_Grid_Summary objEPRO, lstFile.List(i)
                End If
            Next
            
        Case "ETS":
        Dim objETS As New clsETS
            For i = 0 To lstFile.ListCount - 1
                Set objETS = New clsETS
                objETS.Init vCurrentFolder & "\" & lstFile.List(i)
                If objETS.Completed Then
                    objFileReport.Add objETS
                    'each file
                    add_data_to_Grid_Summary objETS, lstFile.List(i)
                End If
            Next

        Case "MAV":
            For i = 0 To lstFile.ListCount - 1
                Set objMAV = New clsMAV
                objMAV.Init vCurrentFolder & "\" & lstFile.List(i)
                If objMAV.Completed Then
                    objFileReport.Add objMAV
                    add_data_to_Grid_Summary objMAV, lstFile.List(i)
                End If
            Next
        Case "TMT":
            For i = 0 To lstFile.ListCount - 1
                Set objTMT = New clsTMT
                objTMT.Init vCurrentFolder & "\" & lstFile.List(i)
                If objTMT.Completed Then
                    objFileReport.Add objTMT
                    add_data_to_Grid_Summary objTMT, lstFile.List(i)
                End If
            Next
End Select
 'summary file
            add_data_to_FT_Grid_Summary objFileReport
            add_data_to_QA_Grid_Summary objFileReport
    
 Me.MousePointer = 0
End Sub



'Private Sub cmdRefresh_Click()
'     Me.MousePointer = 11
'         Set colFiles = New Collection
'        lstFile.Clear
'        txtLotNumber.Text = ""
'        vCurrentFileExt = txtFolder.Text
'        showFileCountInFolder txtFolder.Text, vCurrentFileExt
'    Me.MousePointer = 0
'End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " version : " & App.Major & "." & App.Minor & "." & App.Revision
    '
    initial_Grid_Summary
    initial_Grid_FT
    initial_Grid_QA
End Sub
Sub initial_Grid_Summary()
        With MSFlexGrid1
    
        .Rows = 2
        .Row = 1
        .Clear
        
    .Cols = 12 + 32
    .ColWidth(0) = 3200
    For i = 1 To .Cols - 1
        .ColWidth(i) = 700
    Next
    
    .Row = 0
    .Text = "Lot"
    .col = 1
    .Text = "Step"
    .col = 2
    .Text = "Temp"
    .col = 3
    .Text = "Total"
    .col = 4
    .Text = "Pass"
    .col = 5
    .Text = "Fail"
    'Hardware Bin
    .col = 6
    .Text = "HBin2"
    .col = 7
    .Text = "HBin3"
    .col = 8
    .Text = "HBin4"
    .col = 9
    .Text = "HBin5"
    .col = 10
    .Text = "HBin6"
    .col = 11
    .Text = "HBin7"
    .col = 12
    .Text = "HBin8"
    
    For i = 1 To 31
        .col = 12 + i
        .Text = "DB" & Trim(i)
    Next
'    .col = 13
'    .Text = "DB1"
'    .col = 14
'    .Text = "DB2"
'    .col = 15
'    .Text = "DB3"
'    .col = 16
'    .Text = "DB4"
'    .col = 17
'    .Text = "DB5"
'    .col = 18
'    .Text = "DB6"
    
    End With
End Sub

Sub initial_Grid_FT()
        With fGridFT
    
        .Rows = 2
        .Row = 1
        .Clear
        
    .Cols = 14
    .ColWidth(0) = 3200
    For i = 1 To .Cols - 1
        .ColWidth(i) = 1000
    Next
    
    .Row = 0
    .Text = "Lot"
    .col = 1
    .Text = "Step"
    .col = 2
    .Text = "Temp"
    .col = 3
    .Text = "Total"
    .col = 4
    .Text = "Pass"
    .col = 5
    .Text = "Fail"
    .col = 6
    .Text = "Yield"
    'Hardware Bin
    .col = 7
    .Text = "HBin2"
    .col = 8
    .Text = "HBin3"
    .col = 9
    .Text = "HBin4"
    .col = 10
    .Text = "HBin5"
    .col = 11
    .Text = "HBin6"
    .col = 12
    .Text = "HBin7"
    .col = 13
    .Text = "HBin8"
    
    
    End With
End Sub

Sub initial_Grid_QA()
        With fGridQa
    
        .Rows = 2
        .Row = 1
        .Clear
        
    .Cols = 14
    .ColWidth(0) = 3200
    For i = 1 To .Cols - 1
        .ColWidth(i) = 1000
    Next
    
    .Row = 0
    .Text = "Lot"
    .col = 1
    .Text = "Step"
    .col = 2
    .Text = "Temp"
    .col = 3
    .Text = "Total"
    .col = 4
    .Text = "Pass"
    .col = 5
    .Text = "Fail"
    
    .col = 6
    .Text = "Yield"
    'Hardware Bin
    .col = 7
    .Text = "HBin2"
    .col = 8
    .Text = "HBin3"
    .col = 9
    .Text = "HBin4"
    .col = 10
    .Text = "HBin5"
    .col = 11
    .Text = "HBin6"
    .col = 12
    .Text = "HBin7"
    .col = 13
    .Text = "HBin8"
    
    
    End With
End Sub

Sub add_data_to_Grid_Summary(obj As Object, vFileName As String)
    Dim i As Integer
    
    With MSFlexGrid1
    If .Row = 0 Then
        .Row = 1
        .col = 0
        .Text = vFileName 'obj.Lot
        .col = 1
        .Text = obj.Seq
        .col = 2
        .Text = obj.Temperature
        .col = 3
        .Text = obj.Tested
        .col = 4
        .Text = obj.Passed
        .col = 5
        .Text = obj.Failed
        
'        .col = 6
'        .Text = objEts.Failed
    Else
        .AddItem vFileName 'obj.Lot
        .Row = .Row + 1
        .col = 0
        .Text = vFileName 'obj.Lot
        .col = 1
        .Text = obj.Seq
        .col = 2
        .Text = obj.Temperature
        .col = 3
        .Text = obj.Tested
        .col = 4
        .Text = obj.Passed
        .col = 5
        .Text = obj.Failed
    End If
    
    Dim ixCol As Integer
    Dim objFind As Object
    'Hardware Bin
    ixCol = 2
    For i = 6 To 12
        Dim vSeqArry() As String
        Dim vBinNumber As Integer
        vSeqArry = Split(obj.Seq, "_")
        If UBound(vSeqArry) > 0 Then
            If vSeqArry(0) Like "R*" Then
                If vSeqArry(1) Like "B*" Then
                    vBinNumber = Val(Replace(vSeqArry(1), "B", ""))
                End If
            End If
            'MsgBox ("ssds")
        End If
    
        .col = i
        Set objFind = obj.getBin(Trim(Str(ixCol)), obj.HardwareBins)
        If Not objFind Is Nothing Then
            .Text = objFind.Total
            
            
        End If
        If vBinNumber = i - 4 Then
                .CellBackColor = vbRed
        End If
        ixCol = ixCol + 1
    Next
    
    'Software Bin
    ixCol = 1
    For i = 13 To 13 + 30
        .col = i
        Set objFind = obj.getBin(Trim(Str(ixCol)), obj.SoftwareBins)
        If Not objFind Is Nothing Then
            .Text = objFind.Total
        End If
        ixCol = ixCol + 1
    Next
    
    
    End With
End Sub


Sub add_data_to_FT_Grid_Summary(objs As Collection)
    Dim i As Integer
    Dim vLot As String
    Dim vTested As Long
    Dim vPassed As Long
    Dim vFailed As Long
    Dim vRetestPassed As Long
    
    Dim vHWBin1 As Long
    Dim vHWBin2 As Long
    Dim vHWBin3 As Long
    Dim vHWBin4 As Long
    Dim vHWBin5 As Long
    Dim vHWBin6 As Long
    Dim vHWBin7 As Long
    
    Dim vTemp As String

    
    
    For Each obj In objs
        Dim vSeq As String
        vSeq = obj.Seq
        vLot = obj.Lot
        
        Dim vSeqArry() As String
        Dim vBinNumber As Integer
        vSeqArry = Split(obj.Seq, "_")
        If UBound(vSeqArry) > 0 Then
            If vSeqArry(0) Like "R*" Then
                If vSeqArry(1) Like "B*" Then
                    vBinNumber = Val(Replace(vSeqArry(1), "B", ""))
                End If
            End If
        End If
        
        'Functional
        Dim vFunctionTest As Boolean
        vFunctionTest = IIf(Mid(vSeq, 1, 1) = "F", True, False)
        If (Len(vSeq) = 2 And Mid(vSeq, 1, 1) = "F") Or Mid(vSeq, 1, 1) = "R" Then
            vFailed = vFailed + obj.Failed
            vTested = vTested + obj.Tested
            vPassed = vPassed + obj.Passed
            Dim objHwBins As Object
            Set objHWBin = obj.getBin("1", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
'                If vFunctionTest Then
'                    vHWBin1 = vHWBin1 + objHWBin.Total
'                End If
                vHWBin1 = vHWBin1 + objHWBin.Total
            End If
            
            Set objHWBin = obj.getBin("2", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
'                If vFunctionTest Then
'                    vHWBin2 = vHWBin2 + objHWBin.Total
'                End If
                vHWBin2 = vHWBin2 + objHWBin.Total
                If vBinNumber = 2 Then
                    vHWBin2 = objHWBin.Total
                End If
            End If
            
            Set objHWBin = obj.getBin("3", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
'                If vFunctionTest Then
'                    vHWBin3 = vHWBin3 + objHWBin.Total
'                End If
                vHWBin3 = vHWBin3 + objHWBin.Total
                If vBinNumber = 3 Then
                    vHWBin3 = objHWBin.Total
                End If
            End If
            
            Set objHWBin = obj.getBin("4", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
'                If vFunctionTest Then
'                    vHWBin4 = vHWBin4 + objHWBin.Total
'                End If
                vHWBin4 = vHWBin4 + objHWBin.Total
                If vBinNumber = 4 Then
                    vHWBin4 = objHWBin.Total
                End If
            End If
            
            Set objHWBin = obj.getBin("5", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
'                If vFunctionTest Then
'                    vHWBin5 = vHWBin5 + objHWBin.Total
'                End If
                vHWBin5 = vHWBin5 + objHWBin.Total
                If vBinNumber = 5 Then
                    vHWBin5 = objHWBin.Total
                End If
            End If
            
            Set objHWBin = obj.getBin("6", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
'                If vFunctionTest Then
'                    vHWBin6 = vHWBin6 + objHWBin.Total
'                End If
                vHWBin6 = vHWBin6 + objHWBin.Total
                If vBinNumber = 6 Then
                    vHWBin6 = objHWBin.Total
                End If
            End If
            
            Set objHWBin = obj.getBin("7", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
'                If vFunctionTest Then
'                    vHWBin7 = vHWBin7 + objHWBin.Total
'                End If
                vHWBin7 = vHWBin7 + objHWBin.Total
                If vBinNumber = 7 Then
                    vHWBin7 = objHWBin.Total
                End If
            End If
        End If
        
        
        'Retest
        If Len(vSeq) >= 2 And Mid(vSeq, 1, 1) = "R" Then
            vRetestPassed = obj.Passed
        End If
        
        vTemp = obj.Temperature
                
    Next
    
    vPassed = vPassed + vRetestPassed
    
                   
    
    
    With fGridFT
    If .Row = 0 Then
        .Row = 1
        .col = 0
        .Text = vLot
        .col = 1
        .Text = "F" 'obj.Seq
        .col = 2
        .Text = vTemp 'obj.Temperature
        .col = 3
        .Text = vTested 'obj.Tested
        .col = 4
        .Text = vPassed 'obj.Passed
        .col = 5
        .Text = (vTested - vPassed) 'obj.Failed
        
        .col = 6
        If vPassed <> 0 And vTested <> 0 Then
            .Text = Format((vPassed / vTested) * 100, "#0.00") 'obj.Failed
        Else
            .Text = "0"
        End If
        'Hardware Bin (start bin2)
        .col = 7
        .Text = vHWBin2
            If vBinNumber = 2 Then
                .CellBackColor = vbRed
            End If
        .col = 8
        .Text = vHWBin3
            If vBinNumber = 3 Then
                .CellBackColor = vbRed
            End If
        .col = 9
        .Text = vHWBin4
            If vBinNumber = 4 Then
                .CellBackColor = vbRed
            End If
        .col = 10
        .Text = vHWBin5
            If vBinNumber = 5 Then
                .CellBackColor = vbRed
            End If
        .col = 11
        .Text = vHWBin6
            If vBinNumber = 6 Then
                .CellBackColor = vbRed
            End If
        .col = 12
        .Text = vHWBin7
            If vBinNumber = 7 Then
                .CellBackColor = vbRed
            End If
    End If
    
'    Dim ixCol As Integer
'    Dim objFind As Object
'    'Hardware Bin
'    ixCol = 2
'    For i = 6 To 12
'        .col = i
'        Set objFind = obj.getBin(Trim(Str(ixCol)), obj.HardwareBins)
'        If Not objFind Is Nothing Then
'            .Text = objFind.Total
'        End If
'        ixCol = ixCol + 1
'    Next
'
'    'Software Bin
'    ixCol = 1
'    For i = 13 To 18
'        .col = i
'        Set objFind = obj.getBin(Trim(Str(ixCol)), obj.SoftwareBins)
'        If Not objFind Is Nothing Then
'            .Text = objFind.Total
'        End If
'        ixCol = ixCol + 1
'    Next
    
    
    End With
End Sub

Sub add_data_to_QA_Grid_Summary(objs As Collection)
    Dim i As Integer
    Dim vLot As String
    Dim vTested As Long
    Dim vPassed As Long
    Dim vFailed As Long
    
    Dim vHWBin1 As Long
    Dim vHWBin2 As Long
    Dim vHWBin3 As Long
    Dim vHWBin4 As Long
    Dim vHWBin5 As Long
    Dim vHWBin6 As Long
    Dim vHWBin7 As Long

    Dim vTemp As String
    Dim vFirstTested As Boolean
    vFirstTested = True
    
    For Each obj In objs
        Dim vSeq As String
        vSeq = obj.Seq
        vLot = obj.Lot
        'Functional
        If Len(vSeq) = 2 And Mid(vSeq, 1, 1) = "Q" Then
            If vFirstTested Then
                vTested = vTested + obj.Tested
                vFirstTested = False
            End If
            
            vPassed = vPassed + obj.Passed
            
            
            Dim objHwBins As Object
            Set objHWBin = obj.getBin("1", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
                vHWBin1 = vHWBin1 + objHWBin.Total
            End If
            
            Set objHWBin = obj.getBin("2", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
                vHWBin2 = vHWBin2 + objHWBin.Total
            End If
            
            Set objHWBin = obj.getBin("3", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
                vHWBin3 = vHWBin3 + objHWBin.Total
            End If
            
            Set objHWBin = obj.getBin("4", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
                vHWBin4 = vHWBin4 + objHWBin.Total
            End If
            
            Set objHWBin = obj.getBin("5", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
                vHWBin5 = vHWBin5 + objHWBin.Total
            End If
            
            Set objHWBin = obj.getBin("6", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
                vHWBin6 = vHWBin6 + objHWBin.Total
            End If
            
            Set objHWBin = obj.getBin("7", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
                vHWBin7 = vHWBin7 + objHWBin.Total
            End If
        End If
        vTemp = obj.Temperature
    Next
    
    vFailed = vTested - vPassed
    
    
    With fGridQa
    If .Row = 0 Then
        .Row = 1
        .col = 0
        .Text = vLot
        .col = 1
        .Text = "Q" 'obj.Seq
        .col = 2
        .Text = vTemp 'obj.Temperature
        .col = 3
        .Text = vTested 'obj.Tested
        .col = 4
        .Text = vPassed 'obj.Passed
        .col = 5
        .Text = vFailed 'obj.Failed
        
        .col = 6
        If vPassed <> 0 And vTested <> 0 Then
            .Text = Format((vPassed / vTested) * 100, "#0.00") 'obj.Failed
        Else
            .Text = "0"
        End If
        'Hardware Bin (start bin2)
        .col = 7
        .Text = vHWBin2
        .col = 8
        .Text = vHWBin3
        .col = 9
        .Text = vHWBin4
        .col = 10
        .Text = vHWBin5
        .col = 11
        .Text = vHWBin6
        .col = 12
        .Text = vHWBin7
    End If
    
'    Dim ixCol As Integer
'    Dim objFind As Object
'    'Hardware Bin
'    ixCol = 2
'    For i = 6 To 12
'        .col = i
'        Set objFind = obj.getBin(Trim(Str(ixCol)), obj.HardwareBins)
'        If Not objFind Is Nothing Then
'            .Text = objFind.Total
'        End If
'        ixCol = ixCol + 1
'    Next
'
'    'Software Bin
'    ixCol = 1
'    For i = 13 To 18
'        .col = i
'        Set objFind = obj.getBin(Trim(Str(ixCol)), obj.SoftwareBins)
'        If Not objFind Is Nothing Then
'            .Text = objFind.Total
'        End If
'        ixCol = ixCol + 1
'    Next
    
    
    End With
End Sub

Private Sub txtLotNumber_Change()
    lstFile.Clear
    initial_Grid_Summary
    initial_Grid_FT
    initial_Grid_QA
End Sub

Private Sub txtLotNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        filterFileByLot txtLotNumber.Text
        cmdGenerate_Click
    End If
End Sub
