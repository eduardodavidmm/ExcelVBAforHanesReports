' THIS WORKBOOK

Private Sub Workbook_Open()
  executeProcess
End Sub



Private Sub executeProcess_BK()
    
    Dim finishProcess As Boolean
    finishProcess = False
    
    If verifyConections = True Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        
        'Get the hub and week to download the info
        WeekSelected = getWeekToAnalyze()

        'REGISTER WHEN START THE PROCESS
        startProcess = Now()
        
        'CALL THE FUNCTION TO GET THE CORP BUSINESS UNIT
        getCorpBusinessUnit_FromCPS False
                
        'CALL THE FUNCTION TO REMOVE ALL THE RECORDS BY WEEK.
        Delete_Rows_Based_On_WeekValue
        
        'GET THE RANGE OF DAYS IN THE WEEK
        'GetRangeOfDatesByFiscalWeek
               
        'CALL THE FUNCTION TO GET THE PLANTS TO BE ANALIZED.
        GetPlants
        
        'CALL THE FUNCTION TO UPDATE THE EXCEL FILE
        'openAndUpdateExcelFileToImportSharePoint
        
        'ImportToSharepoint
        
        'END PROCESS
        endProcess = Now()
        totalProcess = DATEDIFF("n", startProcess, endProcess)
        
        finishProcess = endProcessReg(startProcess, endProcess, totalProcess)
        
            
        If finishProcess = True Then
            Call SendEmail("the process was executed correctly", "Ended process:" & vbCrLf & "Fiscal Week: " & WeekSelected & vbCrLf & "Start Process: " & startProcess & vbCrLf & "End Process: " & endProcess & vbCrLf & "Total time: " & totalProcess & "Minutes")
        Else
            Call SendEmail("Update process not executed", "The process did not finish correctly")
        End If
            
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        
        SaveAndClose
    End If
End Sub




 Private Function endProcessReg_BK(initialDate As Date, endDate As Date, duration As Integer) As Boolean
    Dim totalRows As Integer
    Dim rowCount As Integer
    totalRows = Sheet1.Range("Calendar").Rows.Count
    rowCount = 2

    If totalRows > 0 Then
        For rowCount = 2 To totalRows + 1
            If Sheet1.Cells(rowCount, 1) = WeekSelected Then
                Sheet1.Cells(rowCount, 6) = True
                Sheet1.Cells(rowCount, 7) = initialDate
                Sheet1.Cells(rowCount, 8) = endDate
                Sheet1.Cells(rowCount, 9) = duration
                Exit For
            End If
        Next rowCount
    End If
    endProcessReg = True
End Function
