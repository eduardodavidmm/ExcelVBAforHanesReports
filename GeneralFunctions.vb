' GENERAL FUNCTIONS

Option Explicit



' =====FILL THE COMBOBOX HUB SUPPLY CHAIN WITH INFO.=====
Public Sub cmdUpdateDropDowns_Click()
    Dim totalRows As Integer
    Dim rowCount As Integer
    
    totalRows = Sheet17.Range("SupplyChainHub").Rows.Count
    
    If totalRows > 0 Then
        For rowCount = 2 To totalRows + 1
            Sheet3.cmbHub.AddItem Trim(Sheet17.Cells(rowCount, 2))
        Next rowCount
        MsgBox "Process to get Supply Chain Hub finished successfully"
    Else
        MsgBox "No data available for Supply Chain Hub.", vbCritical + vbOKOnly
        Exit Sub
    End If
End Sub






' ===== GET THE DIVISION BY STYLE =====
Public Function getCorpBusinessUnit_WorkCenter_MegaWorkCenter(mfgStyle As String) As Variant
    Dim totalRows As Integer
    Dim rowCount As Integer
    Dim corpBusinessUnit_WorkCenter_MegaWorkCenter As Variant
    Dim findValues As Boolean
    
    totalRows = Sheet8.Range("BusinessUnitData").Rows.Count
    rowCount = 2
    findValues = False
    
    For rowCount = 2 To totalRows + 1
        If Sheet8.Cells(rowCount, 1).Value = mfgStyle Then
            ReDim corpBusinessUnit_WorkCenter_MegaWorkCenter(1 To 3)
            corpBusinessUnit_WorkCenter_MegaWorkCenter(1) = Sheet8.Cells(rowCount, 2).Value
            corpBusinessUnit_WorkCenter_MegaWorkCenter(2) = Sheet8.Cells(rowCount, 3).Value
            corpBusinessUnit_WorkCenter_MegaWorkCenter(3) = Sheet8.Cells(rowCount, 4).Value
            findValues = True
            Exit For
        End If
    Next rowCount
        
    If findValues = False Then
        corpBusinessUnit_WorkCenter_MegaWorkCenter = getCorpBusinessUnit_WorkCenter_MegaWorkCenter_byStyle_FromCPS(mfgStyle)
    End If

    getCorpBusinessUnit_WorkCenter_MegaWorkCenter = corpBusinessUnit_WorkCenter_MegaWorkCenter
End Function








'GET THE CORP BUSINESS UNIT FROM CPS BY MFGSTYLE
Public Function getCorpBusinessUnit_WorkCenter_MegaWorkCenter_byStyle_FromCPS(mfgStyle As String) As Variant
    
    Dim oSheetName As Worksheet
    Dim sTableName As String
    Dim tbl As ListObject
    Dim newrow As ListRow
    ReDim corpBusinessUnit_WorkCenter_MegaWorkCenter(1 To 3)
       
    setConnSQLCPS  ' Set connection to the database.
     
    ' SQL query to fetch details about WorkOrders
    sQuery = "SELECT  B.Style,B.BusinessCode,B.WorkcenterCode,B.MegaWorkcenterCode FROm dbo.vStyleMaster B WHERE style='" & mfgStyle & "'"
    
    
    If CPSrs.State = adStateOpen Then
        CPSrs.Close
    End If
    
    ' Execute query using recordset object.
    CPSrs.CursorLocation = adUseClient
    CPSConn.CommandTimeout = 0
    CPSrs.Open sQuery, CPSConn, adOpenKeyset, adLockOptimistic
    
    'Define Variable
    sTableName = "BusinessUnitData"
    
    'Define WorkSheet object
    Set oSheetName = Sheets("CorpBusinessUnitData")
    
    'Define Table Object
    Set tbl = oSheetName.ListObjects(sTableName)
                    
    ' Finally, show the details.
    If CPSrs.RecordCount > 0 Then
         Do While Not CPSrs.EOF
            'Add New row to the table
            Set newrow = tbl.ListRows.Add
            With newrow
                .Range(1) = "'" & Trim(CPSrs.Fields("Style").Value)
                .Range(2) = CPSrs.Fields("BusinessCode").Value
                .Range(3) = CPSrs.Fields("WorkcenterCode").Value
                .Range(4) = CPSrs.Fields("MegaWorkcenterCode").Value
                
                corpBusinessUnit_WorkCenter_MegaWorkCenter(1) = CPSrs.Fields("BusinessCode").Value
                corpBusinessUnit_WorkCenter_MegaWorkCenter(2) = CPSrs.Fields("WorkcenterCode").Value
                corpBusinessUnit_WorkCenter_MegaWorkCenter(3) = CPSrs.Fields("MegaWorkcenterCode").Value
                
            End With
            CPSrs.MoveNext
        Loop
    Else
        corpBusinessUnit_WorkCenter_MegaWorkCenter(1) = "N/A"
        corpBusinessUnit_WorkCenter_MegaWorkCenter(2) = "N/A"
        corpBusinessUnit_WorkCenter_MegaWorkCenter(3) = "N/A"
    End If
    
    getCorpBusinessUnit_WorkCenter_MegaWorkCenter_byStyle_FromCPS = corpBusinessUnit_WorkCenter_MegaWorkCenter
    
End Function





' ===== GET THE PLANT NAME BY PLANT ID =====
Public Function getPlantName(plantCode As String) As String
    Dim totalRows As Integer
    Dim rowCount As Integer
    Dim plantName As String
    totalRows = Sheet5.Range("PlantsInfo").Rows.Count
    rowCount = 2
    plantName = "N/A"
    If totalRows > 0 Then
        For rowCount = 2 To totalRows + 1
            If Sheet5.Cells(rowCount, 1) = SupplyChainHub And Sheet5.Cells(rowCount, 2) = plantCode Then
                plantName = Sheet5.Cells(rowCount, 3)
                Exit For
            End If
        Next rowCount
    End If
    getPlantName = plantName
End Function




'CALL THE FUNCTION WHEN THE USER PRESS THE BUTTOM
Public Sub call_getCorpBusinessUnit_FromCPS_FromButtom()
    getCorpBusinessUnit_FromCPS True
End Sub





'GET THE CORP BUSINESS UNIT FROM CPS
Public Sub getCorpBusinessUnit_FromCPS(showMessage As Boolean)
    
    Dim oSheetName As Worksheet
    Dim sTableName As String
    Dim tbl As ListObject
    Dim newrow As ListRow

    GetActivePlants     ' get the plants to find the
    setConnSQLCPS       ' Set connection to the database.
     
    ' SQL query to fetch details about WorkOrders
    sQuery = "SELECT DISTINCT B.Style,B.BusinessCode,B.WorkcenterCode,B.MegaWorkcenterCode FROM dbo.vFullWIP a RIGHT JOIN dbo.vStyleMaster b ON a.Style=b.Style WHERE Asewplantcode IN (" & plants & ")"
    If CPSrs.State = adStateOpen Then
        CPSrs.Close
    End If
    
    ' Execute query using recordset object.
    CPSrs.CursorLocation = adUseClient
    CPSConn.CommandTimeout = 0
    CPSrs.Open sQuery, CPSConn, adOpenKeyset, adLockOptimistic
    
    'Define Variable
    sTableName = "BusinessUnitData"
    
    'Define WorkSheet object
    Set oSheetName = Sheets("CorpBusinessUnitData")
    
    'Define Table Object
    Set tbl = oSheetName.ListObjects(sTableName)
                    
     'DELETE THE TABLE INFOTMATION
    With oSheetName.ListObjects(sTableName)
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
    
    ' Finally, show the details.
    If CPSrs.RecordCount > 0 Then
        tbl.ListRows.Add
        
        'fill the table
         tbl.DataBodyRange.CopyFromRecordset CPSrs
        
        'message successfully process
        If showMessage = True Then
            MsgBox "Process to get Corp Bussines Unit finished successfully"
        End If
    End If
End Sub






'GET THE CORP BUSINESS UNIT FROM CPS BY MFGSTYLE
Public Function getCorpBusinessUnit_byStyle_FromCPS(mfgStyle As String) As String
    
    Dim oSheetName As Worksheet
    Dim sTableName As String
    Dim tbl As ListObject
    Dim newrow As ListRow
    Dim corpBusinessUnit As String
    
    
    setConnSQLCPS  ' Set connection to the database.
     
    ' SQL query to fetch details about WorkOrders
    sQuery = "SELECT  B.Style,B.BusinessCode,B.WorkcenterCode,B.MegaWorkcenterCode FROm dbo.vStyleMaster B WHERE style='" & mfgStyle & "'"
    
    
    If CPSrs.State = adStateOpen Then
        CPSrs.Close
    End If
    
    ' Execute query using recordset object.
    CPSrs.CursorLocation = adUseClient
    CPSConn.CommandTimeout = 0
    CPSrs.Open sQuery, CPSConn, adOpenKeyset, adLockOptimistic
    
    'Define Variable
    sTableName = "BusinessUnitData"
    
    'Define WorkSheet object
    Set oSheetName = Sheets("CorpBusinessUnitData")
    
    'Define Table Object
    Set tbl = oSheetName.ListObjects(sTableName)
                    
    ' Finally, show the details.
    If CPSrs.RecordCount > 0 Then
         Do While Not CPSrs.EOF
            'Add New row to the table
            Set newrow = tbl.ListRows.Add
            With newrow
                .Range(1) = "'" & Trim(CPSrs.Fields("Style").Value)
                .Range(2) = CPSrs.Fields("BusinessCode").Value
                .Range(3) = CPSrs.Fields("WorkcenterCode").Value
                .Range(4) = CPSrs.Fields("MegaWorkcenterCode").Value
                corpBusinessUnit = CPSrs.Fields("BusinessCode").Value
            End With
            CPSrs.MoveNext
        Loop
    Else
        corpBusinessUnit = "N/A"
    End If
    getCorpBusinessUnit_byStyle_FromCPS = corpBusinessUnit
End Function









' ===== OBTAIN THE ACTIVE PLANTS TO OBTAIN THE CORP BUSSINES UNIT =====
Public Sub GetActivePlants()
    Dim totalRows As Integer
    Dim rowCount As Integer
    Dim totalRowsActive As Integer
    Dim rowCountActive As Integer
    
    plants = ""
    totalRows = Sheet5.Range("PlantsInfo").Rows.Count
     
    If totalRows > 0 Then
        'count rows with the column value A
        
        For rowCount = 2 To totalRows + 1
            If Sheet5.Cells(rowCount, 4).Value = "A" Then
               totalRowsActive = totalRowsActive + 1
            End If
        Next rowCount
        
        rowCount = 0
        
        For rowCount = 2 To totalRows + 1
            If Sheet5.Cells(rowCount, 4).Value = "A" Then
                plants = plants & "'" & Trim(Sheet5.Cells(rowCount, 2).Value) & "'"
                rowCountActive = rowCountActive + 1
                If rowCountActive < totalRowsActive Then
                    plants = plants & ","
                End If
            End If
        Next rowCount
    End If
End Sub








' ===== GET THE WEEK TO BE ANALYZED =====
Public Function getWeekToAnalyze() As String
    Dim totalRows As Integer
    Dim rowCount As Integer
    Dim week As String
    Dim status As Boolean
    
    totalRows = Sheet1.Range("Calendar").Rows.Count
    rowCount = 2
    week = "N/A"
    status = True
    
    If totalRows > 0 Then
        
        While status = True
            status = Sheet1.Cells(rowCount, 6)
            week = Sheet1.Cells(rowCount, 1)
            rowCount = rowCount + 1
        Wend
             
    End If
    getWeekToAnalyze = week
End Function





' ===== SAVE AND CLOSE DE INFORMATION=====
Public Sub SaveAndClose()
    'ActiveWorkbook.Close SaveChanges:=True
    'ActiveWorkbook.Save
    'OpenAccessToImportSharepointList
    'ActiveWorkbook.Close
End Sub




' ===== GET THE EMAIL USERS=====
Public Function getEmailList() As String
    Dim totalRows As Integer
    Dim rowCount As Integer
    Dim emails As String
    emails = ""
    totalRows = Sheet35.Range("emailList").Rows.Count
    
    If totalRows > 0 Then
        For rowCount = 2 To totalRows + 1
            emails = emails & Sheet35.Cells(rowCount, 2).Value
            If rowCount < totalRows + 1 Then
                emails = emails & ";"
            End If
        Next rowCount
    End If
    getEmailList = emails
End Function




' ===== SEND THE EMAIL TO THE USERS=====
Public Function SendEmail(subject As String, body As String) As Boolean
    On Error GoTo ErrorControl
    Dim email As String
    email = getEmailList()
    
    Dim oEmail As CDO.message
    Set oEmail = CreateObject("CDO.Message")
    
    oEmail.From = "LeadtimeAutomaticProcess@hanes.com"
    oEmail.To = email
    oEmail.subject = subject
    oEmail.Textbody = body
    'oEmail.AddAttachment "C:TempTextFile.TXT"

    oEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    oEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "10.1.5.93"
    oEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/authenticate") = 1
    oEmail.Configuration.Fields.Update

    oEmail.Send
    Set oEmail = Nothing
    
done:
    SendEmail = True
    Exit Function
ErrorControl:
    SendEmail = False
    Exit Function
End Function




'OPEN EXCEL FILE UPDATE THE INFO TO IMPORT TO SHAREPOINT
Public Sub openAndUpdateExcelFileToImportSharePoint()
    Application.ScreenUpdating = False
    ActiveWorkbook.Save
    Dim wb As Excel.Workbook
    Set wb = Application.Workbooks.Open("https://hanes.sharepoint.com/sites/LeadTimeCalculationReportSourcefiles/Shared%20Documents/General/DEV/Source%20Files/Gerencial%20Lead%20Time%20Calculation%20Report.xlsx", ReadOnly:=False)
    wb.Sheets("HistoryTransaction").Range("HistoryTransaction").ListObject.QueryTable.Refresh BackgroundQuery:=False
    wb.Save
    wb.Close
    Application.ScreenUpdating = True
End Sub

