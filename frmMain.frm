VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Tester Yield Report"
   ClientHeight    =   5295
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14505
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   14505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   12510
      TabIndex        =   19
      Top             =   450
      Width           =   1455
   End
   Begin VB.ListBox lstFile 
      Height          =   3765
      Left            =   90
      TabIndex        =   18
      Top             =   1350
      Width           =   5370
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   495
      Left            =   10530
      TabIndex        =   16
      Top             =   1215
      Width           =   1455
   End
   Begin VB.TextBox txtFolder 
      Height          =   375
      Left            =   1170
      TabIndex        =   15
      Top             =   540
      Width           =   9780
   End
   Begin VB.TextBox txtLotNumber 
      Height          =   375
      Left            =   4365
      TabIndex        =   14
      Top             =   90
      Width           =   2580
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export Excel"
      Enabled         =   0   'False
      Height          =   495
      Left            =   12420
      TabIndex        =   12
      Top             =   4770
      Width           =   1455
   End
   Begin VB.ComboBox cbTester 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   1185
      List            =   "frmMain.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   11
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
   Begin VB.Frame Frame4 
      Caption         =   "Yield"
      Height          =   1695
      Left            =   8865
      TabIndex        =   4
      Top             =   3510
      Width           =   2655
      Begin VB.Label lblYield 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Total Failed"
      Height          =   1695
      Left            =   5850
      TabIndex        =   3
      Top             =   3825
      Width           =   2655
      Begin VB.Label lblFailed 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Total Passed"
      Height          =   1695
      Left            =   8595
      TabIndex        =   2
      Top             =   1980
      Width           =   2655
      Begin VB.Label lblPassed 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Total Tested"
      Height          =   1695
      Left            =   11925
      TabIndex        =   1
      Top             =   1800
      Width           =   2655
      Begin VB.Label lblTested 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   2175
      End
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
      TabIndex        =   17
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
      TabIndex        =   13
      Top             =   180
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Tester :"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Folder name :"
      Height          =   255
      Left            =   90
      TabIndex        =   9
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
    txtLotNumber.Text = ""
    cmdExport.Enabled = False
    'lblFileName.Caption = ""
    txtFolder.Text = ""
    lblTested.Caption = "0"
    lblPassed.Caption = "0"
    lblFailed.Caption = "0"
    lblYield.Caption = "0"
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
    For Each f In colFiles
        If f Like "*" & lotNumber & "*" Then
            lstFile.AddItem f
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

Private Sub cmdRefresh_Click()
     Me.MousePointer = 11
         Set colFiles = New Collection
        lstFile.Clear
        txtLotNumber.Text = ""
        vCurrentFileExt = txtFolder.Text
        showFileCountInFolder txtFolder.Text, vCurrentFileExt
    Me.MousePointer = 0
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " version : " & App.Major & "." & App.Minor & "." & App.Revision

End Sub

Private Sub txtLotNumber_Change()
    lstFile.Clear
End Sub

Private Sub txtLotNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        filterFileByLot txtLotNumber.Text
    End If
End Sub
