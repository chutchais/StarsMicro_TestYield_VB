Attribute VB_Name = "mdlExcel"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                            ByVal lpOperation As String, ByVal lpFile As String, _
                            ByVal lpParameters As String, ByVal lpDirectory As String, _
                            Optional nShowCmd As Long) As Long



Sub EPRO_Export_Excel(objEPRO As clsEPRO)
    Dim oExcel As Object
   Dim oBook As Object
   Dim oSheet As Object

   'Start a new workbook in Excel
   Set oExcel = CreateObject("Excel.Application")
   Set oBook = oExcel.Workbooks.Add


   'Add data to cells of the first worksheet in the new workbook
   Set oSheet = oBook.Worksheets(1)
   'Add Column
   With oSheet
        .range("A1").Value = "Assy"
        .range("B1").Value = "Lot"
        .range("C1").Value = "Seq"
        .range("D1").Value = "Operator"
        .range("E1").Value = "Temperature"
        .range("F1").Value = "Tester"
        .range("G1").Value = "Handler"
        .range("H1").Value = "Summary"
        .range("I1").Value = "Date"
        .range("J1").Value = "Time"
        .range("K1").Value = "SystemID"
        .range("L1").Value = "ProgramName"
        .range("M1").Value = "Total"
        .range("N1").Value = "Pass"
        .range("O1").Value = "Fail"
        .range("P1").Value = "Yield"
        Dim k As Integer
        k = 16
        'IB
        For i = 1 To 6
            .cells(1, k + i).Value = "IB" & Trim(i)
        Next
        'DB
        k = k + 6
        For i = 1 To 32
            .cells(1, k + i).Value = "DB" & Trim(i)
        Next
        'BIN
        k = k + 32
        For i = 1 To 32
            .cells(1, k + i).Value = "BIN" & Trim(i)
        Next
        
   End With
   'Add Data
    With oSheet
        .range("A2").Value = objEPRO.AssyNumber
        .range("B2").Value = objEPRO.Lot
        .range("C2").Value = objEPRO.Seq
        .range("D2").Value = objEPRO.Operator
        .range("E2").Value = objEPRO.Temperature
        .range("F2").Value = objEPRO.Tester
        .range("G2").Value = objEPRO.Handler
        .range("H2").Value = objEPRO.SummaryName
        .range("I2").Value = objEPRO.StartDate
        .range("J2").Value = objEPRO.StartTime
        .range("K2").Value = objEPRO.SystemId
        .range("L2").Value = objEPRO.TestProgram
        .range("M2").Value = objEPRO.Tested
        .range("N2").Value = objEPRO.Passed
        .range("O2").Value = objEPRO.Failed
        .range("P2").Value = objEPRO.Yield
        'Dim k As Integer
        'IB
        k = 16
        For i = 1 To 6
        'Find IB from Collection
        Dim vObj As Object
            Set vObj = objEPRO.getBin(Trim(i), objEPRO.IBs)
            If vObj Is Nothing Then
                .cells(2, k + i).Value = ""
            Else
                .cells(2, k + i).Value = vObj.Total
            End If
        Next
        'DB
        k = k + 6
        For i = 1 To 32
            'Fine DB from Collection
            Set vObj = objEPRO.getBin(Trim(i), objEPRO.DBs)
                If vObj Is Nothing Then
                    .cells(2, k + i).Value = ""
                Else
                    .cells(2, k + i).Value = vObj.Total
                End If
        Next
        'Data BIN
        k = k + 32
        For i = 1 To 32
            'Fine DB from Collection
            Set vObj = objEPRO.getBin(Format(Trim(i), "0#"), objEPRO.DataBins)
                If vObj Is Nothing Then
                    .cells(2, k + i).Value = ""
                Else
                    .cells(2, k + i).Value = vObj.Description
                End If
        Next
        
        
   End With
    
    oSheet.Columns.AutoFit
   'Save the Workbook and Quit Excel
   oBook.SaveAs App.path & "\EPRO_report.xls"
   oExcel.quit
   ShellExecute hwnd, "open", App.path & "\EPRO_report.xls", vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

Sub ETS_Export_Excel(objETS As clsETS)
    Dim oExcel As Object
   Dim oBook As Object
   Dim oSheet As Object

   'Start a new workbook in Excel
   Set oExcel = CreateObject("Excel.Application")
   Set oBook = oExcel.Workbooks.Add


   'Add data to cells of the first worksheet in the new workbook
   Set oSheet = oBook.Worksheets(1)
   'Add Column
   With oSheet
        .cells(1, 1).Value = "Assy"
        .cells(1, 2).Value = "Lot"
        .cells(1, 3).Value = "Seq"
        .cells(1, 4).Value = "Test Name"
        .cells(1, 5).Value = "Tester Name"
        .cells(1, 6).Value = "Start Date"
        .cells(1, 7).Value = "Stop Date"
        .cells(1, 8).Value = "Total"
        .cells(1, 9).Value = "Pass"
        .cells(1, 10).Value = "Fail"
        .cells(1, 11).Value = "Yield"
        Dim k As Integer
        k = 11
        'IB
        For i = 1 To 32
            .cells(1, k + i).Value = "SW_Bin" & Trim(i)
        Next
        
   End With
   'Add Data
    With oSheet
        .cells(2, 1).Value = objETS.AssyNumber
        .cells(2, 2).Value = objETS.Lot
        .cells(2, 3).Value = objETS.Seq
        .cells(2, 4).Value = objETS.TestName
        .cells(2, 5).Value = objETS.Tester
        .cells(2, 6).Value = objETS.StartDate
        .cells(2, 7).Value = objETS.StartDate
        .cells(2, 8).Value = objETS.Tested
        .cells(2, 9).Value = objETS.Passed
        .cells(2, 10).Value = objETS.Failed
        .cells(2, 11).Value = objETS.Yield
        'Dim k As Integer
        'SW Bin
        k = 11
        For i = 1 To 32
        'Find IB from Collection
        Dim vObj As Object
            Set vObj = objETS.getBin(Trim(i), objETS.SoftwareBins)
            If vObj Is Nothing Then
                .cells(2, k + i).Value = ""
            Else
                .cells(2, k + i).Value = vObj.Total
            End If
        Next
      
        
   End With
    
    oSheet.Columns.AutoFit
   'Save the Workbook and Quit Excel
   oBook.SaveAs App.path & "\ETS_report.xls"
   oExcel.quit
   ShellExecute hwnd, "open", App.path & "\ETS_report.xls", vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

Sub MAV_Export_Excel(objMAV As clsMAV)
    Dim oExcel As Object
   Dim oBook As Object
   Dim oSheet As Object

   'Start a new workbook in Excel
   Set oExcel = CreateObject("Excel.Application")
   Set oBook = oExcel.Workbooks.Add


   'Add data to cells of the first worksheet in the new workbook
   Set oSheet = oBook.Worksheets(1)
   'Add Column
   With oSheet
        .cells(1, 1).Value = "Assy"
        .cells(1, 2).Value = "Lot"
        .cells(1, 3).Value = "Seq"
        .cells(1, 4).Value = "Start Date"
        .cells(1, 5).Value = "Stop Date"
        .cells(1, 6).Value = "Test Flow"
        .cells(1, 7).Value = "Test Type"
        .cells(1, 8).Value = "Device Name"
        .cells(1, 9).Value = "Lot size"
        .cells(1, 10).Value = "Test count"
        .cells(1, 11).Value = "Operator"
        .cells(1, 12).Value = "Program Name"
        .cells(1, 13).Value = "Program rev"
        .cells(1, 14).Value = "Tester ID"
        .cells(1, 15).Value = "Handler ID"
        .cells(1, 16).Value = "Total"
        .cells(1, 17).Value = "Pass"
        .cells(1, 18).Value = "Fail"
        .cells(1, 19).Value = "Yield"
        Dim k As Integer
        k = 19
        'Sw site
        For i = 1 To 8
            .cells(1, k + 1).Value = "Test_Site" & Trim(i)
            .cells(1, k + 2).Value = "Pass_Site" & Trim(i)
            .cells(1, k + 3).Value = "Fail_Site" & Trim(i)
            k = k + 3
        Next
        
   End With
   'Add Data
    With oSheet
        .cells(2, 1).Value = objMAV.AssyNumber
        .cells(2, 2).Value = objMAV.Lot
        .cells(2, 3).Value = objMAV.Seq
        .cells(2, 4).Value = objMAV.StartDate
        .cells(2, 5).Value = objMAV.StopDate
        .cells(2, 6).Value = objMAV.TestFlow
        .cells(2, 7).Value = Mid(objMAV.AssyNumber, 1, 2)
        .cells(2, 8).Value = objMAV.DeviceName
        .cells(2, 9).Value = objMAV.LotSize
        .cells(2, 10).Value = objMAV.TestCount
        .cells(2, 11).Value = objMAV.Operator
        .cells(2, 12).Value = objMAV.ProgramName
        .cells(2, 13).Value = objMAV.ProgramRev
        .cells(2, 14).Value = objMAV.Tester
        .cells(2, 15).Value = objMAV.Handler
        .cells(2, 16).Value = objMAV.Tested
        .cells(2, 17).Value = objMAV.Passed
        .cells(2, 18).Value = objMAV.Failed
        .cells(2, 19).Value = objMAV.Yield
        
        k = 19
        'Sw site
        Dim objTested As clsBin
        Dim objPassed As clsBin
        Dim objFailed As clsBin
        Set objTested = objMAV.getBin("UNITS TESTED", objMAV.UnitBins)
        Set objPassed = objMAV.getBin("UNITS PASSED", objMAV.UnitBins)
        Set objFailed = objMAV.getBin("UNITS FAILED", objMAV.UnitBins)
        Dim objTestSite As clsSite
        Dim objPassSite As clsSite
        Dim objFailSite As clsSite
        For i = 1 To 8
            Set objTestSite = getBin(Trim(i), objTested.Sites)
            Set objPassSite = getBin(Trim(i), objPassed.Sites)
            Set objFailSite = getBin(Trim(i), objFailed.Sites)
            If objTestSite Is Nothing Then
                .cells(2, k + 1).Value = ""
            Else
                .cells(2, k + 1).Value = objTestSite.Total
            End If
            '.cells(2, k + 1).Value = IIf(objTestSite Is Nothing, "", objTestSite.Total)
            
            
           ' .cells(2, k + 2).Value = IIf(IsNull(objPassSite), "", objTestSite.Total)
            If objPassSite Is Nothing Then
                .cells(2, k + 2).Value = ""
            Else
                .cells(2, k + 2).Value = objPassSite.Total
            End If
            
            '.cells(2, k + 3).Value = IIf(IsNull(objFailSite), "", objTestSite.Total)
             If objFailSite Is Nothing Then
                .cells(2, k + 3).Value = ""
            Else
                .cells(2, k + 3).Value = objFailSite.Total
            End If
            
            k = k + 3
        Next



   End With
    
    oSheet.Columns.AutoFit
   'Save the Workbook and Quit Excel
   oBook.SaveAs App.path & "\MAV_report.xls"
   oExcel.quit
   ShellExecute hwnd, "open", App.path & "\MAV_report.xls", vbNullString, vbNullString, SW_SHOWNORMAL
End Sub


Sub TMT_Export_Excel(objTMT As clsTMT)
   Dim oExcel As Object
   Dim oBook As Object
   Dim oSheet As Object

   'Start a new workbook in Excel
   Set oExcel = CreateObject("Excel.Application")
   Set oBook = oExcel.Workbooks.Add


   'Add data to cells of the first worksheet in the new workbook
   Set oSheet = oBook.Worksheets(1)
   'Add Column
   With oSheet
        .cells(1, 1).Value = "Assy"
        .cells(1, 2).Value = "Lot"
        .cells(1, 3).Value = "Seq"
        .cells(1, 4).Value = "Operator"
        .cells(1, 5).Value = "Computer"
        .cells(1, 6).Value = "Handler ID"
        .cells(1, 7).Value = "Program Name"
        .cells(1, 8).Value = "Total"
        .cells(1, 9).Value = "Pass"
        .cells(1, 10).Value = "Fail"
        .cells(1, 11).Value = "Yield"
        .cells(1, 12).Value = "Most Fail Bin"
        .cells(1, 13).Value = "Next Fail"
        .cells(1, 14).Value = "Start Date"
        .cells(1, 15).Value = "Stop Date"
                
        Dim k As Integer
        k = 15
        'Sw Bin
        For i = 1 To 8
            .cells(1, k + i).Value = "HW_Bin" & Trim(i)
        Next
        
        'Sw Bin
        k = 15 + 8
        For i = 1 To 32
            .cells(1, k + i).Value = "SW_Bin" & Trim(i)
        Next
        
        'Sw Bin Item
        k = 15 + 8 + 32
        For i = 1 To 32
            .cells(1, k + i).Value = "Bin" & Trim(i) & " Item"
        Next
        
        'Sw Bin Site
        k = 15 + 8 + 32 + 32
        For i = 1 To 32
            .cells(1, k + 1).Value = "SW" & Trim(i) & " Site1"
            .cells(1, k + 2).Value = "SW" & Trim(i) & " Site2"
            .cells(1, k + 3).Value = "SW" & Trim(i) & " Site3"
            .cells(1, k + 4).Value = "SW" & Trim(i) & " Site4"
            k = k + 4
        Next
        
        
   End With
   'Add Data
    With oSheet
        .cells(2, 1).Value = objTMT.AssyNumber
        .cells(2, 2).Value = objTMT.Lot
        .cells(2, 3).Value = objTMT.Seq
        .cells(2, 4).Value = objTMT.Operator
        .cells(2, 5).Value = objTMT.Computer
        .cells(2, 6).Value = objTMT.Handler
        .cells(2, 7).Value = objTMT.TestProgram
        .cells(2, 8).Value = objTMT.Tested
        .cells(2, 9).Value = objTMT.Passed
        .cells(2, 10).Value = objTMT.Failed
        .cells(2, 11).Value = objTMT.Yield
        .cells(2, 12).Value = objTMT.MostFailBin
        .cells(2, 13).Value = objTMT.NextSerial
        .cells(2, 14).Value = objTMT.StartDate
        .cells(2, 15).Value = objTMT.StopDate

        k = 15
        Dim objHwBins As clsBin
        For i = 1 To 8
            Set objHwBins = getBin(Trim(i), objTMT.HardwareBins)
            If objHwBins Is Nothing Then
                .cells(2, k + i).Value = ""
            Else
                .cells(2, k + i).Value = objHwBins.Total
            End If
            
        Next
        
        k = 15 + 8
        Dim s As Integer
        s = 15 + 8 + 32 + 32
        
        Dim objSwBins As clsBin
        Dim objSite As clsSite
        
        For i = 1 To 32
            Set objSwBins = getBin(Trim(i), objTMT.SoftwareBins)
            If objSwBins Is Nothing Then
                .cells(2, k + i).Value = ""
                s = s + 4
            Else
                .cells(2, k + i).Value = objSwBins.Total
                .cells(2, k + i + 32).Value = objSwBins.Description
                'Fill Site data
                Dim vSite As String
                Dim vTotal As Double
                For Each objSite In objSwBins.Sites
                    vSite = objSite.Name
                    vTotal = objSite.Total
                    .cells(2, s + Val(vSite)).Value = vTotal
                   ' Set objSite = Nothing
                Next
                s = s + 4
            End If
            
        Next

   End With
    
    oSheet.Columns.AutoFit
   'Save the Workbook and Quit Excel
   oBook.SaveAs App.path & "\TMT_report.xls"
   oExcel.quit
   ShellExecute hwnd, "open", App.path & "\TMT_report.xls", vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

Function getBin(key As String, col As Collection) As Object
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


