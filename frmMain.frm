VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Tester Yield Report"
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16890
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   16890
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbTestStep 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   5265
      List            =   "frmMain.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   990
      Width           =   810
   End
   Begin VB.TextBox txtPartId 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1170
      TabIndex        =   20
      Top             =   945
      Width           =   3210
   End
   Begin VB.CheckBox chkUpdateIndex 
      Caption         =   "update index"
      Height          =   285
      Left            =   8460
      TabIndex        =   18
      Top             =   135
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7755
      Left            =   90
      TabIndex        =   9
      Top             =   1350
      Width           =   16740
      _ExtentX        =   29528
      _ExtentY        =   13679
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Summary Report"
      TabPicture(0)   =   "frmMain.frx":0020
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "File list"
      TabPicture(1)   =   "frmMain.frx":003C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstFile"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "QA Summarize"
         Height          =   1635
         Left            =   135
         TabIndex        =   14
         Top             =   2205
         Width           =   16620
         Begin MSFlexGridLib.MSFlexGrid fGridQa 
            Height          =   1185
            Left            =   135
            TabIndex        =   16
            Top             =   360
            Width           =   16440
            _ExtentX        =   28998
            _ExtentY        =   2090
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "FT Summarize"
         Height          =   1635
         Left            =   135
         TabIndex        =   13
         Top             =   495
         Width           =   16620
         Begin MSFlexGridLib.MSFlexGrid fGridFT 
            Height          =   1185
            Left            =   135
            TabIndex        =   15
            Top             =   360
            Width           =   16440
            _ExtentX        =   28998
            _ExtentY        =   2090
            _Version        =   393216
         End
      End
      Begin VB.ListBox lstFile 
         Height          =   6495
         Left            =   -74775
         TabIndex        =   11
         Top             =   495
         Width           =   16305
      End
      Begin VB.Frame Frame1 
         Caption         =   "Summary Table"
         Height          =   3660
         Left            =   90
         TabIndex        =   10
         Top             =   3825
         Width           =   16665
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   3300
            Left            =   135
            TabIndex        =   12
            Top             =   270
            Width           =   16440
            _ExtentX        =   28998
            _ExtentY        =   5821
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
      Left            =   12510
      TabIndex        =   4
      Top             =   450
      Width           =   1455
   End
   Begin VB.ComboBox cbTester 
      Height          =   315
      ItemData        =   "frmMain.frx":0058
      Left            =   1185
      List            =   "frmMain.frx":0068
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
   Begin VB.Label lblLeakActualYield 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   16470
      TabIndex        =   50
      Top             =   1080
      Width           =   360
   End
   Begin VB.Label lblkelvinActualYield 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   16470
      TabIndex        =   49
      Top             =   765
      Width           =   360
   End
   Begin VB.Label lblOSActualYield 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   16470
      TabIndex        =   48
      Top             =   450
      Width           =   360
   End
   Begin VB.Label lblTriggerActualYield 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   16470
      TabIndex        =   47
      Top             =   135
      Width           =   360
   End
   Begin VB.Label lblLeakActualYield 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   16155
      TabIndex        =   46
      Top             =   1080
      Width           =   360
   End
   Begin VB.Label lblkelvinActualYield 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   16155
      TabIndex        =   45
      Top             =   765
      Width           =   360
   End
   Begin VB.Label lblOSActualYield 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   16155
      TabIndex        =   44
      Top             =   450
      Width           =   360
   End
   Begin VB.Label lblTriggerActualYield 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   16155
      TabIndex        =   43
      Top             =   135
      Width           =   360
   End
   Begin VB.Label lblLeakActualYield 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   15795
      TabIndex        =   42
      Top             =   1080
      Width           =   360
   End
   Begin VB.Label lblkelvinActualYield 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   15795
      TabIndex        =   41
      Top             =   765
      Width           =   360
   End
   Begin VB.Label lblOSActualYield 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   15795
      TabIndex        =   40
      Top             =   450
      Width           =   360
   End
   Begin VB.Label lblTriggerActualYield 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   15795
      TabIndex        =   39
      Top             =   135
      Width           =   360
   End
   Begin VB.Label lblLeakActualYield 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   15435
      TabIndex        =   38
      Top             =   1080
      Width           =   360
   End
   Begin VB.Label lblkelvinActualYield 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   15435
      TabIndex        =   37
      Top             =   765
      Width           =   360
   End
   Begin VB.Label lblOSActualYield 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   15435
      TabIndex        =   36
      Top             =   450
      Width           =   360
   End
   Begin VB.Label lblTriggerActualYield 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   15435
      TabIndex        =   35
      Top             =   135
      Width           =   360
   End
   Begin VB.Label lblLeakYield 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15075
      TabIndex        =   34
      Top             =   1080
      Width           =   360
   End
   Begin VB.Label lblkelvinYield 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15075
      TabIndex        =   33
      Top             =   765
      Width           =   360
   End
   Begin VB.Label lblOSYield 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15075
      TabIndex        =   32
      Top             =   450
      Width           =   360
   End
   Begin VB.Label lblTriggerYield 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15075
      TabIndex        =   31
      Top             =   135
      Width           =   360
   End
   Begin VB.Label Label7 
      Caption         =   "Leak Yield :"
      Height          =   255
      Left            =   14175
      TabIndex        =   30
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label6 
      Caption         =   "Kelvin Yield :"
      Height          =   255
      Left            =   14085
      TabIndex        =   29
      Top             =   765
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "O/S Yield :"
      Height          =   255
      Left            =   14220
      TabIndex        =   28
      Top             =   450
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Trigger Yield :"
      Height          =   255
      Left            =   14040
      TabIndex        =   27
      Top             =   135
      Width           =   1035
   End
   Begin VB.Label lblTempStep4 
      Caption         =   "Temp4"
      Height          =   240
      Left            =   11295
      TabIndex        =   26
      Top             =   990
      Width           =   1860
   End
   Begin VB.Label lblTempStep3 
      Caption         =   "Temp3"
      Height          =   240
      Left            =   9405
      TabIndex        =   25
      Top             =   990
      Width           =   1770
   End
   Begin VB.Label lblTempStep2 
      Caption         =   "Temp2"
      Height          =   240
      Left            =   7650
      TabIndex        =   24
      Top             =   990
      Width           =   1680
   End
   Begin VB.Label lblTempStep1 
      Caption         =   "Temp1"
      Height          =   240
      Left            =   6165
      TabIndex        =   23
      Top             =   990
      Width           =   1410
   End
   Begin VB.Label lblTestStep 
      Caption         =   "Test Step :"
      Height          =   255
      Left            =   4455
      TabIndex        =   22
      Top             =   1035
      Width           =   810
   End
   Begin VB.Label lblPartId 
      Caption         =   "PART ID :"
      Height          =   255
      Left            =   180
      TabIndex        =   19
      Top             =   990
      Width           =   810
   End
   Begin VB.Label lblSearchBy 
      Caption         =   "..."
      Height          =   195
      Left            =   9990
      TabIndex        =   17
      Top             =   180
      Width           =   3885
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
Private Declare Function GetTickCount Lib "kernel32" () As Long

'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
 

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

Dim vReportFolder As String

Dim objFileReport As New Collection

'Added on version 1.0.13
Dim vFoundMatrixFile As Boolean
Dim vMatrixFileName As String
Dim iRowCount As Long
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object

Dim vTempStep1 As String
Dim vTempStep2 As String
Dim vTempStep3 As String
Dim vTempStep4 As String

Dim vTempStep1_Result As Boolean
Dim vTempStep2_Result As Boolean
Dim vTempStep3_Result As Boolean
Dim vTempStep4_Result As Boolean

'Added by Chutchai on March 25,2019 (version 1.0.15)
'To validate Trigger yield,O/S,Kelvin and Leak Yield
Dim vTriggerYieldValue As Double
Dim vTriggerOpenShortYieldValue As Double
Dim vTriggerKelvinYieldValue As Double
Dim vTriggerLeakYieldValue As Double

Dim vTriggerOpenShortBin As String
Dim vTriggerKelvinBin As String
Dim vTriggerLeakBin As String

Dim Matrix_Array() As String
Dim vFTSummaryReport As New Collection






Private Sub cbTester_Change()
    cmdExport.Enabled = False
    txtPartId.Text = ""
End Sub

Private Sub cbTester_Click()

    Me.MousePointer = 11
    clearContains
    txtPartId.Text = ""
    
    vCurrentFolder = getSectionString(cbTester.Text, "path")
    vCurrentFileExt = getSectionString(cbTester.Text, "file ext")
    txtFolder.Text = vCurrentFolder
'    showFileCountInFolder vCurrentFolder, vCurrentFileExt
    Me.MousePointer = 0
    
    lblTempStep1.Caption = "Temp1(" & vTempStep1 & ")": lblTempStep1.ForeColor = vbDefault
    lblTempStep2.Caption = "Temp2(" & vTempStep2 & ")": lblTempStep2.ForeColor = vbDefault
    lblTempStep3.Caption = "Temp3(" & vTempStep3 & ")": lblTempStep3.ForeColor = vbDefault
    lblTempStep4.Caption = "Temp4(" & vTempStep4 & ")": lblTempStep4.ForeColor = vbDefault
    
End Sub

Sub showFileCountInFolder(folderName As String, Optional extFile As String = "*.*")
'    Dim oFs As New FileSystemObject
'    If oFs.FolderExists(folderName) Then
'        lblFilesCount.Caption = "Found " & CountFiles(folderName, extFile) & " File(s)"
'    Else
'        lblFilesCount.Caption = "Folder or File doesn't exist"
'    End If
'
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

Private Sub cbTestStep_Click()
    get_PartId_Matrix_Configuration Trim(UCase(txtPartId.Text)), Trim(cbTestStep.Text)
    
    If objFileReport.Count > 0 Then
        checkTempStep objFileReport
    End If
    
End Sub

Private Sub cmdBrowse_Click()
    If vCurrentFileExt = "" Then
        MsgBox "Please select Tester", vbCritical, "Select tester"
        Exit Sub
    End If
     
     Me.MousePointer = 11
    clearContains
    txtFolder.Text = BrowseForFolder(hwnd, "Please select a Server folder.")
    vCurrentFolder = txtFolder.Text
'    showFileCountInFolder txtFolder.Text, vCurrentFileExt
    Me.MousePointer = 0
End Sub

Private Sub cmdExport_Click()
    
'Select Case cbTester.Text
'        Case "EPRO": EPRO_Export_Excel objEPRO
'        Case "ETS": ETS_Export_Excel objETS
'        Case "MAV": MAV_Export_Excel objMAV
'        Case "TMT": TMT_Export_Excel objTMT
'    End Select
'
export_excel

End Sub


Sub export_excel()
Dim vLotDetailReport As String
Dim obj As Object
    'Open file for write
    '[lot].csv file format
    Dim iFileNumber As Integer
    Dim vReportFileName As String
    vReportFileName = vReportFolder & "\" & cbTester.Text & "\" & txtLotNumber & ".csv"
    iFileNumber = FreeFile
    Open vReportFileName For Output As iFileNumber
    
    
    Dim bFirstLine As Boolean
    bFirstLine = True
    
    For Each obj In objFileReport
        If bFirstLine Then
'            Write #iFileNumber, "sep=;"
            Print #iFileNumber, obj.get_report_header
            bFirstLine = False
        End If
        vLotDetailReport = obj.get_all_detail
        Print #iFileNumber, vLotDetailReport
    Next
    
    Close #iFileNumber
    
    MsgBox "Save report finished", vbInformation, "Save file finished"
    
    ShellExecute Me.hwnd, "open", vReportFileName, vbNullString, vbNullString, vbNormalFocus
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
            'vObjEPRO.Init vCurrentFolder & "\" & f
            vObjEPRO.Init f
        If vObjEPRO.Lot Like "*" & lotNumber & "*" Then
            lstFile.AddItem f
            colLots.Add f
        End If
        
    Next
    
End Sub

Sub searchFolder(path As String, strLot As String, strFileExe As String, intFile As Variant, Optional appendIndex As Boolean)
        
        


        Dim fso, curFolder, x, lotFolder, fileList As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        
       
        'Proceed if folder exists
        Dim strFolderName As String
        If fso.FolderExists(path) Then
            Set curFolder = fso.GetFolder(path)
            For Each x In curFolder.SubFolders
            'Label1.Caption = "Working on :" & x: DoEvents
                iFolderNumber = iFolderNumber + 1
                strFolderName = x.path
                    'Check folder that naming correct
                    Dim arrFolder() As String
                    arrFolder = Split(strFolderName, "\")
                    
                    'create new index
                    If Not appendIndex Then
                        If UBound(Split(arrFolder(UBound(arrFolder)), "_")) = 2 Then
                            Set fileList = Nothing
                            Set lotFolder = fso.GetFolder(strFolderName)
                            'Set fileList = lotFolder.Files
                            Dim f As Object
                                For Each f In lotFolder.Files
                                    '--add to index file--
                                    If f.Name Like "*." & strFileExe Then
                                        Print #intFile, f
                                        colFiles.Add (f)
                                    End If
                                    '--------------------
                                    
                                Next
                            iFoundNumber = iFoundNumber + 1
                            GoTo next_loop:
                        End If
                    End If
                    '--------end create new index-----
                    
                    
                    If strFolderName Like "*" & strLot & "*" And UBound(Split(arrFolder(UBound(arrFolder)), "_")) = 2 Then
                        'Set lotFolder = fso.GetFolder(aPath)
                        'k = UBound(Split(arrFolder(UBound(arrFolder)), "_"))
                        Set fileList = Nothing
                        
                        Set lotFolder = fso.GetFolder(strFolderName)
                        'Set fileList = lotFolder.Files
                        Dim f1 As Object
                            For Each f1 In lotFolder.Files
                                    'Append to index file
                                    If f.Name Like "*." & strFileExe Then
                                        Print #intFile, f
                                    End If
                                    '--------------------
                                    colFiles.Add (f1)
  
                            Next
                        
                        iFoundNumber = iFoundNumber + 1
                        GoTo next_loop:
                    End If
                    
                    'check sub folder
                    searchFolder x.path, strLot, strFileExe, intFile
next_loop:
            Next
        End If
        
    End Sub

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

Function searchFiles(strDirectory As String, srtLot As String, _
                    Optional strExt As String = "*") As Double
    
    
    Dim objFso As Object
    Dim objFiles As Object
    Dim objFile As Object
    
    Dim bReIndex As Boolean
    
        'version 1.0.6 Check in master file first.
    'if FOund lot existing then add to collection and exit
    
    Dim notFoundLot As Boolean
    Dim vMasterFile As String
    vMasterFile = App.path & "\index_" & cbTester.Text & ".txt"
    
    
    'version 1.0.8 -- if Re-index checked , delete index file and recreate file again.
    If chkUpdateIndex.Value Or Dir(vMasterFile) = "" Then
        bReIndex = True
        'Delete file
        Dim FileSystemObject As Object
         Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
        If FileSystemObject.FileExists(vMasterFile) Then
            FileSystemObject.DeleteFile vMasterFile
        End If
     
        '-----------------
    End If
    
    

    
    
    Dim bIndexFileExist As Boolean
    
    'Comment on 1.0.8 , regenerate Index file everythome when Re-Index checked.
        If Not bReIndex Then
        bIndexFileExist = True

        Dim vFileContains As String
        vFileContains = FileToString(vMasterFile)
        Dim lineArray() As String
        lineArray = Split(vFileContains, vbCrLf)
        For i = 0 To UBound(lineArray())
            If textContain(lineArray(i), srtLot) Then
                colFiles.Add (lineArray(i))
            End If
        Next

        If colFiles.Count > 0 Then
            lblSearchBy.Caption = "Found in index file.." & colFiles.Count & " file(s)"
            Exit Function
        End If
        notFoundLot = True
       End If
    
    '1.0.7 -- to update index file based on modified date
    'Re-index master file by checking new file most update than index file
    
    
    
    lblSearchBy.Caption = "Re-scan folder..."
    '-----------------------------------------


'   then count only files of that type, otherwise return a count of all files.
ReIndex:

    'Set Error Handling
    On Error GoTo EarlyExit

    'Create objects to get a count of files in the directory
    Set objFso = CreateObject("Scripting.FileSystemObject")
    Set objFiles = objFso.GetFolder(strDirectory).Files
    
    
    
    'version 1.0.6 -- Save file name to master file
    FileNum = FreeFile
    If bReIndex Then
        Open vMasterFile For Output As FileNum
        
    Else
        Open vMasterFile For Append As FileNum
    End If

    '--------------
            
    'Count files (that match the extension if provided)
    Dim iNewFile As Integer
        For Each objFile In objFiles
            
            If objFile.Name Like "*." & strExt Then
                If UBound(Split(objFile.Name, "_")) >= 2 Then
                    Print #FileNum, objFile
                End If
            End If
            
            
            If objFile Like "*" & srtLot & "*." & strExt Then
                colFiles.Add (objFile)
            End If
next_loop:
        Next objFile
                   
    Close #FileNum

EarlyExit:
    'Clean up
    On Error Resume Next
    Close FileNum
    Set objFile = Nothing
    Set objFiles = Nothing
    Set objFso = Nothing
    On Error GoTo 0
End Function

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
            If UCase(Right(objFile.path, (Len(objFile.path) - InStrRev(objFile.path, ".")))) = UCase(strExt) Then
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
    Me.MousePointer = 0
    Exit Sub
End If


Dim lngTime As Long
  Dim lngIndex As Long
  'record start
lngTime = GetTickCount

initial_Grid_Summary
initial_Grid_FT
initial_Grid_QA
clear_all_yield

Set objFileReport = New Collection

Set colFiles = New Collection

'Only EPRO ,can not using file name to filter Lot number (must read in file content)
    If cbTester.Text <> "EPRO" Then
        'filterFileByLot txtLotNumber.Text
        
        searchFiles vCurrentFolder, txtLotNumber.Text, vCurrentFileExt
        filterFileByLot txtLotNumber.Text
    Else
    
        'version 1.0.6 Check in master file first.
        'if FOund lot existing then add to collection and exit
        Dim bReIndex As Boolean
        Dim notFoundLot As Boolean
        Dim vMasterFile As String
        vMasterFile = App.path & "\index_" & cbTester.Text & ".txt"
        
        'version 1.0.8 -- if Re-index checked , delete index file and recreate file again.
        If chkUpdateIndex.Value Or Dir(vMasterFile) = "" Then
            bReIndex = True
            'Delete file
            Dim FileSystemObject As Object
             Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
            If FileSystemObject.FileExists(vMasterFile) Then
                FileSystemObject.DeleteFile vMasterFile
            End If
         
            '-----------------
        End If
        
        
        If Not bReIndex Then
            Dim vFileContains As String
            vFileContains = FileToString(vMasterFile)
            Dim lineArray() As String
            lineArray = Split(vFileContains, vbCrLf)
            For i = 0 To UBound(lineArray())
                If textContain(lineArray(i), txtLotNumber.Text) Then
                    colFiles.Add (lineArray(i))
                End If
            Next
            
            If colFiles.Count > 0 Then
                lblSearchBy.Caption = "Found in index file.."
                GoTo process_file
            End If
            notFoundLot = True
        End If
        lblSearchBy.Caption = "Re-scan folder..."
        '-----------------------------------------
            
        'version 1.0.6 -- Save file name to master file
                Dim FileNameEPRO As Integer
                FileNameEPRO = FreeFile
                If bReIndex Then
                    Open vMasterFile For Output As FileNameEPRO
                Else
                    Open vMasterFile For Append As FileNameEPRO
                End If
                
                '--------------
            
                
                searchFolder vCurrentFolder, txtLotNumber.Text, vCurrentFileExt, FileNameEPRO, notFoundLot
                Close #FileNameEPRO
process_file:
                filterFileByLot txtLotNumber.Text
            End If
lngTime = GetTickCount - lngTime
lblSearchBy.Caption = lblSearchBy.Caption & ". searching time: " & CStr(lngTime) & " ms"

Select Case cbTester.Text
        Case "EPRO":
        Dim objEPRO As New clsEPRO
            For i = 0 To lstFile.ListCount - 1
                Set objEPRO = New clsEPRO
                objEPRO.Init lstFile.List(i)
                If objEPRO.Completed Then
                    objFileReport.Add objEPRO
                    'each file
                    add_data_to_Grid_Summary objEPRO, Replace(lstFile.List(i), vCurrentFolder, "")
                End If
            Next
            
        Case "ETS":
        Dim objETS As New clsETS
            For i = 0 To lstFile.ListCount - 1
                Set objETS = New clsETS
                'objETS.Init vCurrentFolder & "\" & lstFile.List(i)
                objETS.Init lstFile.List(i)
                If objETS.Completed Then
                    objFileReport.Add objETS
                    'each file
                    add_data_to_Grid_Summary objETS, get_only_fileName(lstFile.List(i))
                End If
            Next

        Case "MAV":
            For i = 0 To lstFile.ListCount - 1
                Set objMAV = New clsMAV
                'objMAV.Init vCurrentFolder & "\" & lstFile.List(i)
                objMAV.Init lstFile.List(i)
                If objMAV.Completed Then
                    objFileReport.Add objMAV
                    add_data_to_Grid_Summary objMAV, get_only_fileName(lstFile.List(i))
                End If
            Next
        Case "TMT":
            For i = 0 To lstFile.ListCount - 1
                Set objTMT = New clsTMT
                'objTMT.Init vCurrentFolder & "\" & lstFile.List(i)
                objTMT.Init lstFile.List(i)
                If objTMT.Completed Then
                    objFileReport.Add objTMT
                    add_data_to_Grid_Summary objTMT, get_only_fileName(lstFile.List(i))
                End If
            Next
End Select
 'summary file
            Set vFTSummaryReport = New Collection
            add_data_to_FT_Grid_Summary objFileReport
            add_data_to_QA_Grid_Summary objFileReport
    cmdExport.Enabled = True
 Me.MousePointer = 0
    
    
    'Check temperture
    checkTempStep objFileReport
End Sub

'Added on version 1.0.13 , March 18,2019
Sub checkTempStep(objs As Collection)
    
    
    Dim iFileNumber As Integer
    Dim vReportFileName As String
    vReportFileName = vReportFolder & "\" & cbTester.Text & "\judgment_" & txtLotNumber & ".txt"
    iFileNumber = FreeFile
    Open vReportFileName For Output As iFileNumber
    
    
    'Validate with Matrix
    Dim objTmpSwBin As Object
    Dim vActualYied As Double
    
    clear_all_yield
    
    Dim ixTemperature As Integer
    Dim vResult As Boolean
    Dim vLotJudgmentResult As Boolean
    Dim txt As String
    ixTemperature = 0
    
    For Each b In vFTSummaryReport
        vLotJudgmentResult = True
        Print #iFileNumber, "Yield Trigger of temperature : " & IIf(b.Temperature = "", "25", b.Temperature)
        
        lblTriggerActualYield(ixTemperature).Visible = True
        lblOSActualYield(ixTemperature).Visible = True
        lblkelvinActualYield(ixTemperature).Visible = True
        lblLeakActualYield(ixTemperature).Visible = True
        
        lblTriggerActualYield(ixTemperature).Caption = Format(b.Yield, "##.0")
        If b.Yield < vTriggerYieldValue Then
            'Low yield
            lblTriggerActualYield(ixTemperature).ForeColor = vbRed
            vResult = False
            vLotJudgmentResult = False
        Else
            lblTriggerActualYield(ixTemperature).ForeColor = vbBlue
            vResult = True
        End If
        txt = "1.Trigger Yield " & vTriggerYieldValue & " , Lot  Yield = " & Format(b.Yield, "##.0") & _
            " --> Yield = " & IIf(vResult, "G", "NG") & " (LOT Yield > Trigger Yield = " & IIf(vResult, "G", "NG") & " )"
        Print #iFileNumber, txt
        
        'Open short Bin
        If vTriggerOpenShortBin <> "" And vTriggerOpenShortBin <> "NA" Then
            Set objTmpSwBin = b.getBin(vTriggerOpenShortBin, b.SwBinSums)
            If objTmpSwBin Is Nothing Then
                GoTo bypass_Os
            End If
            vActualYied = (objTmpSwBin.Total / b.Total) * 100
            lblOSActualYield(ixTemperature).Caption = Format(vActualYied, ".000")
            If vActualYied > vTriggerOpenShortYieldValue Then
                lblOSActualYield(ixTemperature).ForeColor = vbRed
                vResult = False
                vLotJudgmentResult = False
            Else
                lblOSActualYield(ixTemperature).ForeColor = vbBlue
                vResult = True
            End If
            txt = "2.Trigger O/S Bin " & vTriggerOpenShortBin & " , Trigger  value = " & vTriggerOpenShortYieldValue & _
                " , Lot Bin " & vTriggerOpenShortBin & " = " & Format(vActualYied, "0.000") & " --> " & IIf(vResult, "G", "NG") & _
                "  (O/S Lot > Trigger value = " & IIf(vResult, "G", "NG") & " )"
            Print #iFileNumber, txt
        End If
bypass_Os:
        'KelVin
        If vTriggerKelvinBin <> "" And vTriggerKelvinBin <> "NA" Then
            Set objTmpSwBin = b.getBin(vTriggerKelvinBin, b.SwBinSums)
            If objTmpSwBin Is Nothing Then
                GoTo bypass_kelvin
            End If
            vActualYied = (objTmpSwBin.Total / b.Total) * 100
            lblkelvinActualYield(ixTemperature).Caption = Format(vActualYied, "#.00")
            If vActualYied > vTriggerKelvinYieldValue Then
                lblkelvinActualYield(ixTemperature).ForeColor = vbRed
                vResult = False
                vLotJudgmentResult = False
            Else
                lblkelvinActualYield(ixTemperature).ForeColor = vbBlue
                vResult = True
            End If
            txt = "3.Trigger Kelvin Bin " & vTriggerKelvinBin & " , Trigger  value = " & vTriggerKelvinYieldValue & _
                " , Lot Bin " & vTriggerKelvinBin & " = " & Format(vActualYied, "0.000") & " --> " & IIf(vResult, "G", "NG") & _
                "  (Kelvin Lot > Trigger value = " & IIf(vResult, "G", "NG") & " )"
            Print #iFileNumber, txt
        End If
bypass_kelvin:
        'Leak
        If vTriggerLeakBin <> "" And vTriggerLeakBin <> "NA" Then
            Set objTmpSwBin = b.getBin(vTriggerLeakBin, b.SwBinSums)
            If objTmpSwBin Is Nothing Then
                GoTo bypass_leak
            End If
            vActualYied = (objTmpSwBin.Total / b.Total) * 100
            lblLeakActualYield(ixTemperature).Caption = Format(vActualYied, "#.00")
            If vActualYied > vTriggerLeakYieldValue Then
                lblLeakActualYield(ixTemperature).ForeColor = vbRed
                vResult = False
                vLotJudgmentResult = False
            Else
                vResult = True
                lblLeakActualYield(ixTemperature).ForeColor = vbBlue
            End If
            txt = "4.Trigger Leak Bin " & vTriggerLeakBin & " , Trigger  value = " & vTriggerLeakYieldValue & _
                " , Lot Bin " & vTriggerLeakBin & " = " & Format(vActualYied, "0.000") & " --> " & IIf(vResult, "G", "NG") & _
                "  (Leak Lot > Trigger value = " & IIf(vResult, "G", "NG") & " )"
            Print #iFileNumber, txt
        End If
bypass_leak:
        Print #iFileNumber, "Lot judgment : " & IIf(vLotJudgmentResult, "G", "NG")
        Print #iFileNumber, ""
    Next 'Next Temperature
    
    'End Trigger Yield
    
    Dim vResultStrTemp1 As String
    Dim vResultStrTemp2 As String
    Dim vResultStrTemp3 As String
    Dim vResultStrTemp4 As String
    
    
    vTempStep1_Result = False
    vTempStep2_Result = False
    vTempStep3_Result = False
    vTempStep4_Result = False
    
    Dim colTemp As Collection
    Set colTemp = getTemperatureCol(objs)
    
    If vTempStep1 <> "" And vTempStep1 <> "NA" Then
        'Check temp in FT summary grid
        vTempStep1_Result = findTemperature(vTempStep1, colTemp)
        lblTempStep1.Caption = lblTempStep1.Caption & " " & IIf(vTempStep1_Result, "G", "NG")
        lblTempStep1.ForeColor = IIf(vTempStep1_Result, vbBlue, vbRed)
        
        vResultStrTemp1 = "Temperature step1 = " & vTempStep1 & ", Test summary Temperature = " & IIf(vTempStep1_Result, vTempStep1, "Not Found") & " --> " & IIf(vTempStep1_Result, "G", "NG")
    Else
        vTempStep1_Result = True
    End If
    If vTempStep2 <> "" And vTempStep2 <> "NA" Then
        'Check temp in FT summary grid
        vTempStep2_Result = findTemperature(vTempStep2, colTemp)
        lblTempStep2.Caption = lblTempStep2.Caption & " " & IIf(vTempStep2_Result, "G", "NG")
        lblTempStep2.ForeColor = IIf(vTempStep2_Result, vbBlue, vbRed)
        
        vResultStrTemp2 = "Temperature step2 = " & vTempStep2 & ", Test summary Temperature = " & IIf(vTempStep2_Result, vTempStep2, "Not Found") & " --> " & IIf(vTempStep2_Result, "G", "NG")
     Else
        vTempStep2_Result = True
    End If
    If vTempStep3 <> "" And vTempStep3 <> "NA" Then
        'Check temp in FT summary grid
        vTempStep3_Result = findTemperature(vTempStep3, colTemp)
        lblTempStep3.Caption = lblTempStep3.Caption & " " & IIf(vTempStep3_Result, "G", "NG")
        lblTempStep3.ForeColor = IIf(vTempStep3_Result, vbBlue, vbRed)
        
        vResultStrTemp3 = "Temperature step3 = " & vTempStep3 & ", Test summary Temperature = " & IIf(vTempStep3_Result, vTempStep3, "Not Found") & " --> " & IIf(vTempStep3_Result, "G", "NG")
     Else
        vTempStep3_Result = True
    End If
    If vTempStep4 <> "" And vTempStep4 <> "NA" Then
        'Check temp in FT summary grid
        vTempStep4_Result = findTemperature(vTempStep4, colTemp)
        lblTempStep4.Caption = lblTempStep4.Caption & " " & IIf(vTempStep4_Result, "G", "NG")
        lblTempStep4.ForeColor = IIf(vTempStep4_Result, vbBlue, vbRed)
        
        vResultStrTemp4 = "Temperature step4 = " & vTempStep4 & ", Test summary Temperature = " & IIf(vTempStep4_Result, vTempStep4, "Not Found") & " --> " & IIf(vTempStep4_Result, "G", "NG")
     Else
        vTempStep4_Result = True
    End If
    
    Print #iFileNumber, "Lot teperature judgment = " & IIf(vTempStep1_Result And vTempStep2_Result And _
                                                    vTempStep3_Result And vTempStep4_Result, "G", "NG")
    Print #iFileNumber, vResultStrTemp1
    Print #iFileNumber, vResultStrTemp2
    Print #iFileNumber, vResultStrTemp3
    Print #iFileNumber, vResultStrTemp4
    Close #iFileNumber
End Sub

Private Function findTemperature(siteName As String, col As Collection) As Boolean
  On Error GoTo errhandler
  For Each c In col
        'remove C
        c = Replace(c, "° C", "")
        c = Replace(c, "°C", "")
        c = Replace(c, "°c", "")
        c = Replace(c, "° c", "")
        c = IIf(c = "", "25", c)
        'c = Int(c)
        If c = Trim(siteName) Then
            findTemperature = True
            Exit For
        End If
  Next
  Exit Function
errhandler:
  Set getSite = Nothing
End Function

Function get_only_fileName(vFullPath As String) As String
    Dim vFileArray() As String
    vFileArray = Split(vFullPath, "\")
    If UBound(vFileArray) > 0 Then
        get_only_fileName = vFileArray(UBound(vFileArray))
    Else
        get_only_fileName = ""
    End If
End Function



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
    
    clear_all_yield
    
    'set default Tester
    Dim vDefaultTester As String
    vDefaultTester = getSectionString("default", "tester")
    If vDefaultTester <> "" Then
        cbTester.Text = vDefaultTester
        
    End If
    
    'set update index option
    Dim vUpdateIndex As String
    vUpdateIndex = getSectionString("default", "update index")
    vUpdateIndex = UCase(vUpdateIndex)
    If vUpdateIndex = "YES" Or vUpdateIndex = "Y" Then
       chkUpdateIndex.Value = 1
    Else
       chkUpdateIndex.Value = 0
    End If
    '-----------------------
    
    'create output folder
    vReportFolder = getSectionString("output", "path")
    Dim oFs As New FileSystemObject
    
    If vReportFolder = "" Then
        vReportFolder = App.path & "\output"
        If Not oFs.FolderExists(vReportFolder) Then
            oFs.CreateFolder (vReportFolder)
        End If
    Else
        If Not oFs.FolderExists(vReportFolder) Then
            oFs.CreateFolder (vReportFolder)
        End If
    End If
    
    'create sub folder
        If Not oFs.FolderExists(vReportFolder & "\ETS") Then
            oFs.CreateFolder (vReportFolder & "\ETS")
        End If
        If Not oFs.FolderExists(vReportFolder & "\TMT") Then
            oFs.CreateFolder (vReportFolder & "\TMT")
        End If
        If Not oFs.FolderExists(vReportFolder & "\MAV") Then
            oFs.CreateFolder (vReportFolder & "\MAV")
        End If
        If Not oFs.FolderExists(vReportFolder & "\EPRO") Then
            oFs.CreateFolder (vReportFolder & "\EPRO")
        End If
    '-----------------------
    'Version 1.0.13 -- Add Test Step Temperature Validation
    Set xFolder = oFs.GetFolder(App.path)
    Dim f As Object
        For Each f In xFolder.Files
            '--add to index file--
            If f.Name Like "*.csv*" Then
                'MsgBox "Found Xlsx file : " & f.Name
                vFoundMatrixFile = True
                vMatrixFileName = f.Name
                Exit For
            End If
        Next

        lblPartId.Enabled = vFoundMatrixFile
        lblTestStep.Enabled = vFoundMatrixFile
        txtPartId.Enabled = vFoundMatrixFile
        cbTestStep.Enabled = vFoundMatrixFile
        If txtPartId.Enabled = False Then
            txtPartId.Text = "Not found Matrix file (csv)"
        Else
            txtPartId.Text = ""
        End If
        
        If vFoundMatrixFile Then
            ' Load the file.
            Dim file_name As String
            file_name = App.path & "\" & vMatrixFileName
            fnum = FreeFile
            Open file_name For Input As fnum
            whole_file = Input$(LOF(fnum), #fnum)
            Close fnum
        
            ' Break the file into lines.
            lines = Split(whole_file, vbCrLf)
        
            ' Dimension the array.
            num_rows = UBound(lines)
            one_line = Split(lines(0), ",")
            num_cols = UBound(one_line)
            ReDim Matrix_Array(num_rows, num_cols)
        
            ' Copy the data into the array.
            For R = 0 To num_rows
                If Len(lines(R)) > 0 Then
                    one_line = Split(lines(R), ",")
                    For c = 0 To num_cols
                        Matrix_Array(R, c) = one_line(c)
                    Next c
                End If
            Next R
        
        End If
        
       ' get_PartId_Matrix_Configuration "CAT4002ASD-GT3"
    
End Sub

Function get_PartId_Matrix_Configuration(PartId As String, _
            Optional TestStep As String = "1") As String
            
            If PartId = "" Then
                Exit Function
            End If
            vTempStep1 = ""
            vTempStep2 = ""
            vTempStep3 = ""
            vTempStep4 = ""
            
            lblTempStep1.Caption = "Temp1": lblTempStep1.ForeColor = vbDefault
            lblTempStep2.Caption = "Temp2": lblTempStep2.ForeColor = vbDefault
            lblTempStep3.Caption = "Temp3": lblTempStep3.ForeColor = vbDefault
            lblTempStep4.Caption = "Temp4": lblTempStep4.ForeColor = vbDefault
            
            vTriggerYieldValue = 0
            vTriggerOpenShortYieldValue = 0
            vTriggerKelvinYieldValue = 0
            vTriggerLeakYieldValue = 0
            
            vTriggerOpenShortBin = 0
            vTriggerKelvinBin = 0
            vTriggerLeakBin = 0
            
            Dim strProductId As String
            Dim strStep As String
            For i = 0 To UBound(Matrix_Array)
                strProductId = Matrix_Array(i, 2)
                strStep = Matrix_Array(i, 39) 'AN
                    If PartId = strProductId And strStep = TestStep Then
                        vTempStep1 = Matrix_Array(i, 40) 'AO
                        vTempStep2 = Matrix_Array(i, 41) 'AP
                        vTempStep3 = Matrix_Array(i, 42) 'AQ
                        vTempStep4 = Matrix_Array(i, 43) 'AR

                        lblTempStep1.Caption = "Temp1 (" & vTempStep1 & ")"
                        lblTempStep2.Caption = "Temp2 (" & vTempStep2 & ")"
                        lblTempStep3.Caption = "Temp3 (" & vTempStep3 & ")"
                        lblTempStep4.Caption = "Temp4 (" & vTempStep4 & ")"


                        'Fill all Yield Trigger value
                        Dim vTmpStr As String
                        vTmpStr = Matrix_Array(i, 32) 'oSheet.range("AG" & Trim(Str(cells.row))).Value
                        If IsNumeric(vTmpStr) Then
                            vTriggerYieldValue = Val(vTmpStr)
                        End If

                        vTmpStr = Matrix_Array(i, 33) 'oSheet.range("AH" & Trim(Str(cells.row))).Value
                        vTriggerOpenShortBin = vTmpStr
'                        If IsNumeric(vTmpStr) Then
'                             vTriggerOpenShortBin = vTmpStr
'                        End If

                        vTmpStr = Matrix_Array(i, 34) 'oSheet.range("AI" & Trim(Str(cells.row))).Value
                        vTriggerOpenShortYieldValue = vTmpStr
'                        If IsNumeric(vTmpStr) Then
'                             vTriggerOpenShortYieldValue = vTmpStr
'                        End If

                        vTmpStr = Matrix_Array(i, 35) 'oSheet.range("AJ" & Trim(Str(cells.row))).Value
                        vTriggerKelvinBin = vTmpStr
'                        If IsNumeric(vTmpStr) Then
'                             vTriggerKelvinBin = vTmpStr
'                        End If

                        vTmpStr = Matrix_Array(i, 36) 'oSheet.range("AK" & Trim(Str(cells.row))).Value
                        If IsNumeric(vTmpStr) Then
                             vTriggerKelvinYieldValue = Val(vTmpStr)
                        End If

                        vTmpStr = Matrix_Array(i, 37) 'oSheet.range("AL" & Trim(Str(cells.row))).Value
                        If IsNumeric(vTmpStr) Then
                             vTriggerLeakBin = Val(vTmpStr)
                        End If

                        vTmpStr = Matrix_Array(i, 38) 'oSheet.range("AM" & Trim(Str(cells.row))).Value
                        If IsNumeric(vTmpStr) Then
                             vTriggerLeakYieldValue = Val(vTmpStr)
                        End If
                        Exit For
                    End If
                
            Next
    
            'Start a new workbook in Excel
'           Set oExcel = CreateObject("Excel.Application")
'           Set oBook = oExcel.Workbooks.Open(App.path & "\" & vMatrixFileName, , True)
'           On Error Resume Next
'           Set oSheet = oBook.Worksheets("Consumption Factor")
'           If oSheet Is Nothing Then
'                 oBook.Close
'                oExcel.quit
'                Set oBook = Nothing
'                Set oExcel = Nothing
'                MsgBox "Not found sheet name :  Consumption Factor", vbExclamation, "Not found Matrix sheet"
'           Else
'
'            iRowCount = oSheet.UsedRange.Rows.Count
'            Dim rng1 As Object
'            Dim cells As Object
'            Set rng1 = oSheet.range("C1:C" & Trim(Str(iRowCount)))
'            Set cells = rng1.Find(PartId)
'            If Not cells Is Nothing Then
'
'                Dim firstCellAddress As String
'                firstCellAddress = cells.Address
'
'                ' Find all cells containing Elli
'                Dim vTestStepData As String
'                Do
'                    Debug.Print "Found: " & cells.Address
'                    vTestStepData = oSheet.range("AN" & Trim(Str(cells.row))).Value
'                    If TestStep = vTestStepData Then
'                        Debug.Print "Found!!!!!!!!!!!!!!!!"
'                        vTempStep1 = oSheet.range("AO" & Trim(Str(cells.row))).Value
'                        vTempStep2 = oSheet.range("AP" & Trim(Str(cells.row))).Value
'                        vTempStep3 = oSheet.range("AQ" & Trim(Str(cells.row))).Value
'                        vTempStep4 = oSheet.range("AR" & Trim(Str(cells.row))).Value
'
'                        lblTempStep1.Caption = "Temp1 (" & vTempStep1 & ")"
'                        lblTempStep2.Caption = "Temp2 (" & vTempStep2 & ")"
'                        lblTempStep3.Caption = "Temp3 (" & vTempStep3 & ")"
'                        lblTempStep4.Caption = "Temp4 (" & vTempStep4 & ")"
'
'
'                        'Fill all Yield Trigger value
'                        Dim vTmpStr As String
'                        vTmpStr = oSheet.range("AG" & Trim(Str(cells.row))).Value
'                        If IsNumeric(vTmpStr) Then
'                            vTriggerYieldValue = Val(vTmpStr)
'                        End If
'
'                        vTmpStr = oSheet.range("AH" & Trim(Str(cells.row))).Value
'                        If IsNumeric(vTmpStr) Then
'                             vTriggerOpenShortBin = Val(vTmpStr)
'                        End If
'
'                        vTmpStr = oSheet.range("AI" & Trim(Str(cells.row))).Value
'                        If IsNumeric(vTmpStr) Then
'                             vTriggerOpenShortYieldValue = Val(vTmpStr)
'                        End If
'
'                        vTmpStr = oSheet.range("AJ" & Trim(Str(cells.row))).Value
'                        If IsNumeric(vTmpStr) Then
'                             vTriggerKelvinBin = Val(vTmpStr)
'                        End If
'
'                        vTmpStr = oSheet.range("AK" & Trim(Str(cells.row))).Value
'                        If IsNumeric(vTmpStr) Then
'                             vTriggerKelvinYieldValue = Val(vTmpStr)
'                        End If
'
'                        vTmpStr = oSheet.range("AL" & Trim(Str(cells.row))).Value
'                        If IsNumeric(vTmpStr) Then
'                             vTriggerLeakBin = Val(vTmpStr)
'                        End If
'
'                        vTmpStr = oSheet.range("AM" & Trim(Str(cells.row))).Value
'                        If IsNumeric(vTmpStr) Then
'                             vTriggerLeakYieldValue = Val(vTmpStr)
'                        End If
'
'
'
'
'
'                        Exit Do
'                    End If
'                    Set cells = rng1.FindNext(cells)
'                Loop While firstCellAddress <> cells.Address
            
    
              
'            Else
'                MsgBox PartId & " not found"
'            End If
            
            lblTriggerYield.Caption = Str(vTriggerYieldValue)
                        lblOSYield.Caption = Str(vTriggerOpenShortYieldValue)
                        lblkelvinYield.Caption = Str(vTriggerKelvinYieldValue)
                        lblLeakYield.Caption = Str(vTriggerLeakYieldValue)

'           End If
           
'            oSheet.Close
'            oBook.Close
'           oExcel.quit
'           Set oSheet = Nothing
'           Set oBook = Nothing
'           Set oExcel = Nothing
           
End Function




Sub initial_Grid_Summary()
        With MSFlexGrid1
    
        .Rows = 2
        .row = 1
        .Clear
        
    .cols = 12 + 33
    .ColWidth(0) = 3200 'file name
    .ColWidth(1) = 1000 'step
    For i = 2 To .cols - 1
        .ColWidth(i) = 700
    Next
    
    .row = 0
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
    
    For i = 1 To 32
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
    
        .Rows = 1
        .row = 0
        .Clear
        
    .cols = 14
    .ColWidth(0) = 3200
    For i = 1 To .cols - 1
        .ColWidth(i) = 1000
    Next
    
    .row = 0
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
    
        .Rows = 1
        .row = 0
        .Clear
        
    .cols = 14
    .ColWidth(0) = 3200
    For i = 1 To .cols - 1
        .ColWidth(i) = 1000
    Next
    
    .row = 0
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
    If .row = 0 Then
        .row = 1
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
        .row = .row + 1
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
        Dim vBinNumberCol As New Collection
        vSeqArry = Split(obj.Seq, "_")
        If UBound(vSeqArry) > 0 Then
            If vSeqArry(0) Like "R*" Then
            
'                If vSeqArry(1) Like "B*" Then
'                    vBinNumber = Val(Replace(vSeqArry(1), "B", ""))
'                End If
            'Multiple Retest BIN
                For x = 1 To UBound(vSeqArry)
                    If vSeqArry(x) Like "B*" Then
                        vBinNumberCol.Add Val(Replace(vSeqArry(x), "B", ""))
                    End If
                Next
            
            
            End If

        End If
    
        .col = i
        Set objFind = obj.getBin(Trim(Str(ixCol)), obj.HardwareBins)
        If Not objFind Is Nothing Then
            .Text = objFind.Total
            
            
        End If
        
'        If vBinNumber = i - 4 Then
'                .CellBackColor = vbGreen
'        End If
        If Not vBinNumberCol Is Nothing Then
            For Each c In vBinNumberCol
                If c = i - 4 Then
                    .CellBackColor = vbGreen
                End If
            Next
         End If
        
        
        ixCol = ixCol + 1
    Next
    
    'Software Bin
    ixCol = 1
    For i = 13 To 13 + 31
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
    'get all Temperature
    Dim colTemp As Collection
    Set colTemp = getTemperatureCol(objs)
    Dim vRow As Integer
    vRow = 1
    Dim objTemp As Collection
    For Each c In colTemp
        Set objTemp = getCollectionByTemperature(CStr(c), objs)
        add_data_to_FT_each_Temperature CStr(c), objTemp, vRow
        vRow = vRow + 1
    Next
End Sub

Sub add_data_to_FT_each_Temperature(Temperature As String, objs As Collection, Optional row As Integer = 1)
    
    If objs.Count = 0 Then
        Exit Sub
    End If
    
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
    
    'Add for 1.0.7 , Feb 13,2019
    Dim vFailed_Fct As Long
    Dim vPassed_Retest As Long
    '----------------------

    Dim vBinNumberCol As New Collection
    
    
    '-----To Keep SW bin summary----
    Dim objSummaryReport As New clsFTsummary
    objSummaryReport.Temperature = Temperature
    
    'vFTSummaryReport.Add ""
    
    
    
    For Each obj In objs
        Dim vSeq As String
        vSeq = obj.Seq
        vLot = obj.Lot
        
       
        If vBinNumberCol.Count = 0 Then
        Dim vSeqArry() As String
        Dim vBinNumber As Integer
        
            vSeqArry = Split(obj.Seq, "_")
            If UBound(vSeqArry) > 0 Then
                If vSeqArry(0) Like "R*" Then
                    For x = 1 To UBound(vSeqArry)
                        If vSeqArry(x) Like "B*" Then
                            vBinNumberCol.Add Val(Replace(vSeqArry(x), "B", ""))
                        End If
                    Next
                End If
            End If
        End If
        
        
        'Functional
        Dim vFunctionTest As Boolean
        vFunctionTest = IIf(Mid(vSeq, 1, 1) = "F", True, False)
        If (Len(vSeq) = 2 And Mid(vSeq, 1, 1) = "F") Or Mid(vSeq, 1, 1) = "R" Then
            
            '//TODO
            Dim objTmpSwBin As Object
            Dim tmpSite As clsSite
            
            For Each y In obj.SoftwareBins
                Set tmpSite = New clsSite
                Set objTmpSwBin = obj.getBin(y.Name, objSummaryReport.SwBinSums)
                tmpSite.Name = y.Name
                tmpSite.Total = y.Total
                tmpSite.Description = y.Description
                
                If objTmpSwBin Is Nothing Then
                    objSummaryReport.SwBinSums.Add tmpSite
                Else
                    
                    If Mid(vSeq, 1, 1) = "R" Then
                        objTmpSwBin.Total = y.Total
                    Else
                        objTmpSwBin.Total = objTmpSwBin.Total + y.Total
                    End If
                End If
            Next
            
            
            'sum all FCT and Retest
            vPassed = vPassed + obj.Passed
            
            'version 1.0.7
            'For Total ,only "F" are calculated.
            If vFunctionTest Then
                vTested = vTested + obj.Tested
                vFailed_Fct = vFailed_Fct + obj.Failed
            Else
                vPassed_Retest = vPassed_Retest + obj.Passed
            End If
            
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
'                If vBinNumber = 2 Then
'                    vHWBin2 = objHWBin.Total
'                End If
                For Each c In vBinNumberCol
                    If c = 2 Then
                        vHWBin2 = objHWBin.Total
                    End If
                Next
                
            End If
            
            Set objHWBin = obj.getBin("3", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
'                If vFunctionTest Then
'                    vHWBin3 = vHWBin3 + objHWBin.Total
'                End If
                vHWBin3 = vHWBin3 + objHWBin.Total
                For Each c In vBinNumberCol
                    If c = 3 Then
                        vHWBin3 = objHWBin.Total
                    End If
                Next
            End If
            
            Set objHWBin = obj.getBin("4", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
'                If vFunctionTest Then
'                    vHWBin4 = vHWBin4 + objHWBin.Total
'                End If
                vHWBin4 = vHWBin4 + objHWBin.Total
                For Each c In vBinNumberCol
                    If c = 4 Then
                        vHWBin4 = objHWBin.Total
                    End If
                Next
            End If
            
            Set objHWBin = obj.getBin("5", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
'                If vFunctionTest Then
'                    vHWBin5 = vHWBin5 + objHWBin.Total
'                End If
                vHWBin5 = vHWBin5 + objHWBin.Total
                For Each c In vBinNumberCol
                    If c = 5 Then
                        vHWBin5 = objHWBin.Total
                    End If
                Next
            End If
            
            Set objHWBin = obj.getBin("6", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
'                If vFunctionTest Then
'                    vHWBin6 = vHWBin6 + objHWBin.Total
'                End If
                vHWBin6 = vHWBin6 + objHWBin.Total
                For Each c In vBinNumberCol
                    If c = 6 Then
                        vHWBin6 = objHWBin.Total
                    End If
                Next
            End If
            
            Set objHWBin = obj.getBin("7", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
'                If vFunctionTest Then
'                    vHWBin7 = vHWBin7 + objHWBin.Total
'                End If
                vHWBin7 = vHWBin7 + objHWBin.Total
                For Each c In vBinNumberCol
                    If c = 7 Then
                        vHWBin7 = objHWBin.Total
                    End If
                Next
            End If
        End If
        
        
        'Retest
        If Len(vSeq) >= 2 And Mid(vSeq, 1, 1) = "R" Then
            vRetestPassed = obj.Passed
        End If
        
        vTemp = obj.Temperature
                
    Next
    
    'comment for 1.0.7 ,Feb 13,2019
    'vPassed = vPassed + vRetestPassed
    vFailed = vFailed_Fct - vPassed_Retest
    
    If vTested = 0 Then
        Exit Sub
    End If
                   
    
    fGridFT.Rows = fGridFT.Rows + 1
    
    '--Add report summary
    objSummaryReport.Total = vTested
    objSummaryReport.Pass = vPassed
    objSummaryReport.Fail = vTested - vPassed
    '-------------------
    
    
    With fGridFT
    'If .row = 0 Then
        .row = fGridFT.Rows - 1
        .col = 0
        .Text = vLot
        .col = 1
        .Text = "F" 'obj.Seq
        .col = 2
        .Text = Temperature 'vTemp 'obj.Temperature
        .col = 3
        .Text = vTested 'obj.Tested
        .col = 4
        .Text = vPassed 'obj.Passed
        .col = 5
        .Text = (vTested - vPassed) 'obj.Failed
        
        .col = 6
        If vPassed <> 0 And vTested <> 0 Then
            .Text = Format((vPassed / vTested) * 100, "#0.00") 'obj.Failed
            objSummaryReport.Yield = (vPassed / vTested) * 100
        Else
            .Text = "0"
            objSummaryReport.Yield = 0
        End If
        'Hardware Bin (start bin2)
        .col = 7
        .Text = vHWBin2
            For Each c In vBinNumberCol
                If c = 2 Then
                .CellBackColor = vbGreen
            End If
            Next
            
        .col = 8
        .Text = vHWBin3
            For Each c In vBinNumberCol
                If c = 3 Then
                    .CellBackColor = vbGreen
                End If
            Next
        .col = 9
        .Text = vHWBin4
            For Each c In vBinNumberCol
                If c = 4 Then
                    .CellBackColor = vbGreen
                End If
            Next
        .col = 10
        .Text = vHWBin5
            For Each c In vBinNumberCol
                If c = 5 Then
                    .CellBackColor = vbGreen
                End If
            Next
        .col = 11
        .Text = vHWBin6
            For Each c In vBinNumberCol
                If c = 6 Then
                    .CellBackColor = vbGreen
                End If
            Next
        .col = 12
        .Text = vHWBin7
            For Each c In vBinNumberCol
                If c = 7 Then
                    .CellBackColor = vbGreen
                End If
            Next
        
        .col = 13
        .Text = vHWBin8
            For Each c In vBinNumberCol
                If c = 8 Then
                    .CellBackColor = vbGreen
                End If
            Next
   ' End If
    

        vFTSummaryReport.Add objSummaryReport
    
    End With
End Sub


Function temperatureExistInCol(Temperature As String, cols As Collection) As Boolean
    For Each c In cols
        If c = Temperature Then
            temperatureExistInCol = True
            Exit Function
        End If
    Next
End Function
Function getTemperatureCol(objs As Collection) As Collection
    Set getTemperatureCol = New Collection
    Dim vPreviousTemp As String
    Dim vCurrTemp As String
    For Each o In objs
        vCurrTemp = o.Temperature
        If Not temperatureExistInCol(vCurrTemp, getTemperatureCol) Then
            getTemperatureCol.Add vCurrTemp
            vPreviousTemp = vCurrTemp
        End If
    Next
End Function

Function getCollectionByTemperature(Temperature As String, objs As Collection) As Collection
    Dim tmpCol As New Collection
    'Set getCollectionByTemperature = New Collection
    For Each o In objs
        If Temperature = o.Temperature Then
            tmpCol.Add o
        End If
    Next
'    Set tmpCol = SortCollection(tmpCol) 'SortCollection(tmpCol)
    Set getCollectionByTemperature = tmpCol
End Function

Sub add_data_to_QA_Grid_Summary(objs As Collection)
    'get all Temperature
    Dim colTemp As Collection
    Set colTemp = getTemperatureCol(objs)
    Dim vRow As Integer
    vRow = 1
    Dim objTemp As Collection
    For Each c In colTemp
        Set objTemp = getCollectionByTemperature(CStr(c), objs)
        
        add_data_to_QA_each_Temp CStr(c), objTemp, vRow
        vRow = vRow + 1
    Next
End Sub

Sub add_data_to_QA_each_Temp(Temperature As String, objs As Collection, Optional row As Integer = 1)

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
    Dim vHWBin8 As Long


    Dim vTemp As String
    Dim vFirstTested As Boolean
    vFirstTested = True
    
    If objs Is Nothing Then
        Exit Sub
    End If
    
    
    Dim vPreviousTested As Long
    Dim vPreviousHWBin2 As Long
    Dim vPreviousHWBin3 As Long
    Dim vPreviousHWBin4 As Long
    Dim vPreviousHWBin5 As Long
    Dim vPreviousHWBin6 As Long
    Dim vPreviousHWBin7 As Long
    Dim vPreviousHWBin8 As Long
    
    For Each obj In objs
        Dim vSeq As String
        vSeq = obj.Seq
        vLot = obj.Lot
        'Functional
        If Len(vSeq) = 2 And Mid(vSeq, 1, 1) = "Q" Then
            
            If obj.Tested > vPreviousTested Then
                vTested = obj.Tested
            End If
            vPreviousTested = obj.Tested
            
'            If vFirstTested Then
'                vTested = vTested + obj.Tested
'                vFirstTested = False
'            End If
            
            vPassed = vPassed + obj.Passed

            Dim objHwBins As Object
            Set objHWBin = obj.getBin("1", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
                vHWBin1 = vHWBin1 + objHWBin.Total
            End If
            
            Set objHWBin = obj.getBin("2", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
                If vPreviousHWBin2 < objHWBin.Total Then
                    vHWBin2 = objHWBin.Total
                    vPreviousHWBin2 = vHWBin2
                End If
            End If
            
            Set objHWBin = obj.getBin("3", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
                If vPreviousHWBin3 < objHWBin.Total Then
                    vHWBin3 = objHWBin.Total
                    vPreviousHWBin3 = vHWBin3
                End If
            End If
            
            Set objHWBin = obj.getBin("4", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
                If vPreviousHWBin4 < objHWBin.Total Then
                    vHWBin4 = objHWBin.Total
                    vPreviousHWBin4 = vHWBin4
                End If
            End If
            
            Set objHWBin = obj.getBin("5", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
                If vPreviousHWBin5 < objHWBin.Total Then
                    vHWBin5 = objHWBin.Total
                    vPreviousHWBin5 = vHWBin5
                End If
            End If
            
            Set objHWBin = obj.getBin("6", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
                If vPreviousHWBin6 < objHWBin.Total Then
                    vHWBin6 = objHWBin.Total
                    vPreviousHWBin6 = vHWBin6
                End If
            End If
            
            Set objHWBin = obj.getBin("7", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
                If vPreviousHWBin7 < objHWBin.Total Then
                    vHWBin7 = objHWBin.Total
                    vPreviousHWBin7 = vHWBin7
                End If
            End If
            
            Set objHWBin = obj.getBin("8", obj.HardwareBins)
            If Not objHWBin Is Nothing Then
                If vPreviousHWBin8 < objHWBin.Total Then
                    vHWBin8 = objHWBin.Total
                    vPreviousHWBin8 = vHWBin8
                End If
            End If
            
        End If
        vTemp = obj.Temperature
    Next
    
    vFailed = vTested - vPassed
    
    If vTested = 0 Then
        Exit Sub
    End If
    
    fGridQa.Rows = fGridQa.Rows + 1
    
    With fGridQa
    'If .row = 0 Then
        .row = .Rows - 1
        .col = 0
        .Text = vLot
        .col = 1
        .Text = "Q" 'obj.Seq
        .col = 2
        .Text = Temperature 'obj.Temperature
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
        .col = 13
        .Text = vHWBin8
    'End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
            oSheet.Close
            oBook.Close
           oExcel.quit
           Set oSheet = Nothing
           Set oBook = Nothing
           Set oExcel = Nothing
End Sub




Private Sub clear_all_yield()

    For i = 0 To 3
        With lblTriggerActualYield(i)
            .Visible = False
            .Caption = ""
        End With
        With lblOSActualYield(i)
            .Visible = False
            .Caption = ""
        End With
        With lblkelvinActualYield(i)
            .Visible = False
            .Caption = ""
        End With
        With lblLeakActualYield(i)
            .Visible = False
            .Caption = ""
        End With
    Next
End Sub



Private Sub txtLotNumber_Change()
    lstFile.Clear
    initial_Grid_Summary
    initial_Grid_FT
    initial_Grid_QA
    clear_all_yield
    
    lblTempStep1.Caption = "Temp1 (" & vTempStep1 & ")": lblTempStep1.ForeColor = vbDefault
    lblTempStep2.Caption = "Temp2 (" & vTempStep2 & ")": lblTempStep2.ForeColor = vbDefault
    lblTempStep3.Caption = "Temp3 (" & vTempStep3 & ")": lblTempStep3.ForeColor = vbDefault
    lblTempStep4.Caption = "Temp4 (" & vTempStep4 & ")": lblTempStep4.ForeColor = vbDefault
                        
End Sub

Private Sub txtLotNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        filterFileByLot txtLotNumber.Text
        cmdGenerate_Click
    End If
End Sub

Private Sub txtPartId_Change()
            vTempStep1 = ""
            vTempStep2 = ""
            vTempStep3 = ""
            vTempStep4 = ""
            
            vTriggerYieldValue = 0
            vTriggerOpenShortYieldValue = 0
            vTriggerKelvinYieldValue = 0
            vTriggerLeakYieldValue = 0
            
            vTriggerOpenShortBin = 0
            vTriggerKelvinBin = 0
            vTriggerLeakBin = 0
            lblTriggerYield.Caption = Str(vTriggerYieldValue)
            lblOSYield.Caption = Str(vTriggerOpenShortYieldValue)
            lblkelvinYield.Caption = Str(vTriggerKelvinYieldValue)
            lblLeakYield.Caption = Str(vTriggerLeakYieldValue)
            
            lblTempStep1.Caption = "Temp1 (" & vTempStep1 & ")": lblTempStep1.ForeColor = vbDefault
            lblTempStep2.Caption = "Temp2 (" & vTempStep2 & ")": lblTempStep2.ForeColor = vbDefault
            lblTempStep3.Caption = "Temp3 (" & vTempStep3 & ")": lblTempStep3.ForeColor = vbDefault
            lblTempStep4.Caption = "Temp4 (" & vTempStep4 & ")": lblTempStep4.ForeColor = vbDefault
End Sub

Private Sub txtPartId_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        get_PartId_Matrix_Configuration Trim(UCase(txtPartId.Text)), Trim(cbTestStep.Text)
        
    End If
End Sub
