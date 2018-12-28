VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Tester Yield Report"
   ClientHeight    =   2970
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   11100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export Excel"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9600
      TabIndex        =   13
      Top             =   720
      Width           =   1455
   End
   Begin VB.ComboBox cbTester 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   960
      List            =   "frmMain.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   120
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8160
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Caption         =   "Yield"
      Height          =   1695
      Left            =   8400
      TabIndex        =   4
      Top             =   1200
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
      Left            =   5640
      TabIndex        =   3
      Top             =   1200
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
      Left            =   2880
      TabIndex        =   2
      Top             =   1200
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
      Left            =   120
      TabIndex        =   1
      Top             =   1200
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
   Begin VB.CommandButton cmdTest 
      Caption         =   "Browse file"
      Height          =   495
      Left            =   9600
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Tester :"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblFileName 
      Caption         =   "..."
      Height          =   375
      Left            =   960
      TabIndex        =   10
      Top             =   480
      Width           =   8415
   End
   Begin VB.Label Label1 
      Caption         =   "File name :"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim objEPRO As clsEPRO
Dim objETS As clsETS
Dim objMAV As clsMAV
Dim objTMT As clsTMT

Private Sub cbTester_Change()
    cmdExport.Enabled = False
End Sub

Private Sub cbTester_Click()
    clearContains
End Sub

Sub clearContains()
    cmdExport.Enabled = False
    lblFileName.Caption = ""
    lblTested.Caption = "0"
    lblPassed.Caption = "0"
    lblFailed.Caption = "0"
    lblYield.Caption = "0"
End Sub

Private Sub cmdExport_Click()
    
Select Case cbTester.Text
        Case "EPRO": EPRO_Export_Excel objEPRO
        Case "ETS": ETS_Export_Excel objETS
        Case "MAV": MAV_Export_Excel objMAV
        Case "TMT": TMT_Export_Excel objTMT
    End Select
    


End Sub




Private Sub cmdTest_Click()
    
   
    Dim vFilter As String
    Dim vExt As String
    Select Case cbTester.Text
        Case "EPRO":
                vFilter = "Apps (*.sum)|*.sum|All files (*.*)|*.sum"
                vExt = "sum"
        Case "ETS":
                vFilter = "Apps (*.txt)|*.txt|All files (*.*)|*.txt"
                vExt = "txt"
        Case "MAV":
                vFilter = "Apps (*.txt)|*.txt|All files (*.*)|*.txt"
                vExt = "txt"
        Case "TMT":
                vFilter = "Apps (*.lsr)|*.lsr|All files (*.*)|*.lsr"
                vExt = "lsr"
    End Select

    With CommonDialog1
        .Filter = vFilter
        .DefaultExt = vExt
        .DialogTitle = "Select File"
        .ShowOpen
    End With
    
    If CommonDialog1.FileName = "" Then Exit Sub
    
    lblFileName.Caption = CommonDialog1.FileName
    
    cmdExport.Enabled = False
    
    Dim objFile As Object
    Select Case cbTester.Text
        Case "EPRO":
            Set objEPRO = New clsEPRO
            objEPRO.Init CommonDialog1.FileName
            
            If objEPRO.Completed Then
                showDetail objEPRO
                cmdExport.Enabled = True
            Else
                MsgBox objEPRO.Description, vbCritical, "Profile Error"
            End If
        
        Case "ETS":
            Set objETS = New clsETS
            objETS.Init CommonDialog1.FileName
            showDetail objETS
            
            If objETS.Completed Then
                showDetail objETS
                cmdExport.Enabled = True
            Else
                MsgBox objETS.Description, vbCritical, "Profile Error"
            End If
        
        Case "MAV":
            
            Set objMAV = New clsMAV
            objMAV.Init CommonDialog1.FileName
            showDetail objMAV
            If objMAV.Completed Then
                showDetail objMAV
                cmdExport.Enabled = True
            Else
                MsgBox objMAV.Description, vbCritical, "Profile Error"
            End If
            
        Case "TMT":
            
            Set objTMT = New clsTMT
            objTMT.Init CommonDialog1.FileName
            showDetail objTMT
            If objTMT.Completed Then
                showDetail objTMT
                cmdExport.Enabled = True
            Else
                MsgBox objTMT.Description, vbCritical, "Profile Error"
            End If
    End Select
    
End Sub

Sub showDetail(vObject As Object)
    With vObject
                lblTested.Caption = Format(.Tested, "###,##0")
                lblPassed.Caption = Format(.Passed, "###,##0")
                lblFailed.Caption = Format(.Failed, "###,##0")
                lblYield.Caption = .Yield
            End With
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " version : " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
