Option Explicit




'DELETE THE ROWS TABLE BY FISCAL WEEK
Public Sub Delete_Rows_Based_On_WeekValue(tableName As String, sheetName As String)
    Dim oSheetName As Worksheet
    Dim sTableName As String
    Dim tbl As ListObject
    
    'Define Variable
    sTableName = tableName
    
    'Define WorkSheet object
    Set oSheetName = Sheets(sheetName)
    
    'Define Table Object
    Set tbl = oSheetName.ListObjects(sTableName)
                    
     'DELETE THE TABLE INFOTMATION
    With oSheetName.ListObjects(sTableName)
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
End Sub





' ===== OBTAIN THE RANGE OF DATE OF THE WEEK=====
Public Sub GetRangeOfDatesByFiscalWeek()
    rowCount = 1
    Set rng = Sheet1.Range("Calendar")
    For Each cell In rng
        If cell(rowCount, 1).Value = WeekSelected Then
            initialDate = cell(rowCount, 2).Value
            endDate = cell(rowCount, 3).Value
            InitialDateTTS = cell(rowCount, 4).Value
            EndDateTTS = cell(rowCount, 5).Value
            rowCount = rowCount + 1
            Exit For
        End If
    Next cell
End Sub





' ===== OBTAIN THE PLANTS TO BE ANALIZED =====
Public Sub GetPlants()
    Dim totalRows As Integer
    Dim rowCount As Integer
    
    plants = ""
    plantCategory = ""
    plantName = ""
    
    totalRows = Sheet5.Range("PlantsInfo").Rows.Count
    
    If totalRows > 0 Then
        'count rows with the column value A
        For rowCount = 2 To totalRows + 1
        
            If Sheet5.Cells(rowCount, 4).Value = "A" Then
            
               plants = Trim(Sheet5.Cells(rowCount, 2).Value)
               plantName = Sheet5.Cells(rowCount, 3).Value
               plantCategory = Sheet5.Cells(rowCount, 5).Value
               SupplyChainHub = Sheet5.Cells(rowCount, 1).Value
               reportType = Sheet5.Cells(rowCount, 6).Value
               
               
                'CALL THE FUNCTION TO OBTAIN THE TRANSMITTED WORKLOTS.
                getTransmittedOrdersFromMADM
                
                
                '=========CALL THE FUCTION TO FIND THE ORDERS INFORMATION IN TTS.
                getWorkOrderToFindInfoFromTTS
                
                
                ' CALL THE FUNCTION BY BUCKET SOURCE(ANET/TTS)
                ' =======ANET
                ' attribution
                GetInformationFromMADN_Attribution
                'sewing
                GetInformationFromMADN_Sewing
                
                'TTS
                If totalTransmittedOrder > 0 Then
                    getGroupOfOrders
                End If
                
                
                '===Call the SKU Changes process.
                GetInformationFromMADN_SKUChangeLots
                
                
                plants = ""
                plantCategory = ""
                plantName = ""
                reportType = ""
                
            End If
        Next rowCount
    End If
End Sub







' GET ALL THE WORKLOST TRASMITTED BY WEEK
Public Sub getTransmittedOrdersFromMADM()
    
    totalTransmittedOrder = 0
    
    setConnSQLMadm     ' Set connection to the database.
    
     With MADMCommand
        Set .ActiveConnection = MADMConn
        .CommandType = adCmdText
        .CommandTimeout = 0
        
        .CommandText = "IF OBJECT_ID('tempdb..#TrasmittedManifest') IS NOT NULL DROP TABLE #TrasmittedManifest"
        .Execute
        .CommandText = "SELECT CD.WorkOrder, CD.MfgStyleCD, CD.MfgColorCD, CD.MfgSizeCD, CD.MfgSizeDESC, CD.MfgAttributeCD, CD.MfgRevisionNO, CD.PkgStyleCD, CD.PkgColorCD, CD.PkgSizeCD, CD.PkgSizeDESC, CD.PkgAttributeCD, CD.PkgRevisionNO, CD.SelStyleCD, CD.SelColorCD, CD.SelSizeCD, CD.SelSizeDESC, CD.SelAttributeCD, CD.SelRevisionNO, CD.ParentWorkOrder, CD.AssortmentParentWO, CD.FromPlantCD, CD.CutPlantCD, CD.SewPlantCD, CD.Lot, CD.ParentLot,  lastTrans.lastUpdate, SUM(round(Convert(Decimal(10, 4), CD.PiecesApplied) / Convert(Decimal(10, 4), 12), 4)) As Doz " & _
                       " into #TrasmittedManifest " & _
                       " FROM (SELECT Manifest, WorkOrder, MAX(ANETCreatedOnDate) AS lastUpdate FROM Manufacturing.dbo.ANETCOODetails WITH (NOLOCK) " & _
                       " WHERE (CONVERT(date, ANETCreatedOnDate) BETWEEN '" & initialDate & "' AND '" & endDate & "') AND (FromPlantCD IN('" & plants & "')) AND (HSCD IS NOT NULL) GROUP BY Manifest, WorkOrder) AS lastTrans INNER JOIN " & _
                       " Manufacturing.dbo.ANETCOODetails AS CD WITH (NOLOCK) ON lastTrans.Manifest = CD.Manifest AND lastTrans.WorkOrder = CD.WorkOrder AND lastTrans.lastUpdate = CD.ANETCreatedOnDate GROUP BY CD.WorkOrder, CD.MfgStyleCD, CD.MfgColorCD, CD.MfgSizeCD, CD.MfgSizeDESC, CD.MfgAttributeCD, CD.MfgRevisionNO, CD.PkgStyleCD, CD.PkgColorCD, CD.PkgSizeCD, CD.PkgSizeDESC, CD.PkgAttributeCD, CD.PkgRevisionNO, CD.SelStyleCD, CD.SelColorCD, CD.SelSizeCD, CD.SelSizeDESC, CD.SelAttributeCD, CD.SelRevisionNO, CD.ParentWorkOrder, CD.AssortmentParentWO, CD.FromPlantCD, CD.CutPlantCD, CD.SewPlantCD,CD.SewPlantCD, CD.Lot, CD.ParentLot,  lastTrans.lastUpdate"
        .Execute
            
            
        .CommandText = "IF OBJECT_ID('tempdb..#WOInfo') IS NOT NULL DROP TABLE #WOInfo"
        .Execute
        .CommandText = "select DISTINCT TM.WorkOrder ,OWT.GlobalWorkOrderID " & _
                       " INTO #WOInfo " & _
                       " from #TrasmittedManifest as TM " & _
                       " INNER JOIN Manufacturing.dbo.ANETWorkOrders AS AWO WITH (NOLOCK)  ON TM.WorkOrder = AWO.WorkOrder AND TM.MfgStyleCD = AWO.StyleCD AND TM.MfgColorCD = AWO.ColorCD AND TM.MfgSizeCD = AWO.SizeCD AND TM.MfgAttributeCD =AWO.AttributeCD   " & _
                       " INNER JOIN dbo.ANETOriginalWorkOrderTrace as OWT WITH (NOLOCK)  ON AWO.GlobalWorkOrderID  = OWT.GlobalWorkOrderID "
        .Execute
            

        .CommandText = "IF OBJECT_ID('tempdb..#TransmittedWorkLots') IS NOT NULL DROP TABLE #TransmittedWorkLots"
        .Execute
        .CommandText = "select TM.WorkOrder , TM.ParentWorkOrder, TM.AssortmentParentWO, TM.FromPlantCD, TM.CutPlantCD, " & _
                       " TM.SewPlantCD,SUM(TM.Doz) AS DOZ, OWT.OriginalWorkOrder, OWT.ApparelNETWorkLot, OWT.MfgStyle, OWT.MfgColor, OWT.MfgSizeCD, OWT.MfgSizeDesc, OWT.PkgStyle, OWT.PkgColor, " & _
                       " OWT.PkgSizeCD, OWT.PkgSizeDesc, OWT.SellStyle, OWT.SellColor, OWT.SellSizeCD, " & _
                       " OWT.SellSizeDesc, OWT.CutLoc, OWT.SewLoc, ISNULL(TRY_CONVERT(int,OWT.OriginalWorkOrder),0) as WO_NUMBER  " & _
                       " INTO #TransmittedWorkLots " & _
                       " from #TrasmittedManifest as TM " & _
                       " INNER JOIN #WOInfo AS WO on TM.WorkOrder =  wo.WorkOrder " & _
                       " LEFT JOIN dbo.ANETOriginalWorkOrderTrace as OWT WITH (NOLOCK)  ON WO.GlobalWorkOrderID  = OWT.GlobalWorkOrderID " & _
                       " GROUP BY TM.WorkOrder, TM.ParentWorkOrder, TM.AssortmentParentWO, TM.FromPlantCD, TM.CutPlantCD, " & _
                       " TM.SewPlantCD, OWT.OriginalWorkOrder, OWT.ApparelNETWorkLot, OWT.MfgStyle, OWT.MfgColor, OWT.MfgSizeCD, OWT.MfgSizeDesc, OWT.PkgStyle, OWT.PkgColor, " & _
                       " OWT.PkgSizeCD, OWT.PkgSizeDesc, OWT.SellStyle, OWT.SellColor, OWT.SellSizeCD, " & _
                       " OWT.SellSizeDesc , OWT.CutLoc, OWT.SewLoc " & _
                       " ORDER BY TM.WorkOrder "
        .Execute
       
    End With

     
    sQuery = "SELECT DISTINCT WO_NUMBER AS WO,SewPlantCD as Plant FROM #TransmittedWorkLots"
        
    ' Execute query using recordset object.
    MADMrs.CursorLocation = adUseClient
    MADMConn.CommandTimeout = 0
    MADMrs.Open sQuery, MADMConn, adOpenKeyset, adLockOptimistic
        
    Dim iCnt As Integer
    iCnt = 1
    
    Sheet6.Columns(1).ClearContents
        
    ' Finally, show the details.
    If MADMrs.RecordCount > 0 Then
        totalTransmittedOrder = MADMrs.RecordCount
        Do While Not MADMrs.EOF
            Sheet6.Cells(iCnt, 1) = "'" & MADMrs.Fields("WO").Value
            Sheet6.Cells(iCnt, 2) = "'" & MADMrs.Fields("Plant").Value
            MADMrs.MoveNext
            iCnt = iCnt + 1
        Loop
    End If
End Sub





'OBTAIN THE BUCKETS BY HUB AND SOURCE TO GET THE DATA.
Public Function getBucketInfoBySupplyHub(source As String, Process As String) As Integer
    Dim totalRows As Integer
    Dim rowCount As Integer
    totalRows = Sheet2.Range("Buckets").Rows.Count
    
    Dim totalRowsByBucket As Integer
    Dim oSheetName As Worksheet
    Dim sTableName As String
    Dim tbl As ListObject
    Dim newrow As ListRow
    
    'Define Variable
    sTableName = "BucketsAnalysis"
                    
    'Define WorkSheet object
    Set oSheetName = Sheets("BucketInfoAnalysis")
                                     
    'Define Table Object
    Set tbl = oSheetName.ListObjects(sTableName)
    
    'DELETE THE TABLE INFOTMATION
    With oSheetName.ListObjects(sTableName)
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
    For rowCount = 2 To totalRows + 1
        If Sheet2.Cells(rowCount, 1) = SupplyChainHub And Sheet2.Cells(rowCount, 8) = source And Sheet2.Cells(rowCount, 20) = reportType And Sheet2.Cells(rowCount, 21) = Process Then
            Set newrow = tbl.ListRows.Add
            With newrow
                .Range(1) = "'" & Sheet2.Cells(rowCount, 1).Value
                .Range(2) = "'" & Sheet2.Cells(rowCount, 2).Value
                .Range(3) = "'" & Sheet2.Cells(rowCount, 3).Value
                .Range(4) = "'" & Sheet2.Cells(rowCount, 4).Value
                .Range(5) = "'" & Sheet2.Cells(rowCount, 5).Value
                .Range(6) = "'" & Sheet2.Cells(rowCount, 6).Value
                .Range(7) = "'" & Sheet2.Cells(rowCount, 7).Value
                .Range(8) = "'" & Sheet2.Cells(rowCount, 8).Value
                .Range(9) = "'" & Sheet2.Cells(rowCount, 9).Value
                .Range(10) = "'" & Sheet2.Cells(rowCount, 10).Value
                .Range(11) = "'" & Sheet2.Cells(rowCount, 11).Value
                .Range(12) = "'" & Sheet2.Cells(rowCount, 12).Value
                .Range(13) = "'" & Sheet2.Cells(rowCount, 13).Value
                .Range(14) = "'" & Sheet2.Cells(rowCount, 14).Value
                .Range(15) = "'" & Sheet2.Cells(rowCount, 20).Value
                .Range(16) = "'" & Sheet2.Cells(rowCount, 21).Value
            End With
            totalRowsByBucket = totalRowsByBucket + 1
        End If
    Next rowCount
    getBucketInfoBySupplyHub = totalRowsByBucket
End Function








'FUNCTION TO OBTAIN THE WORKLOTS TRANSACTIONS FROM MADM FOR THE ATTRIBUTION PROCESS
Public Sub GetInformationFromMADN_Attribution()
    Dim oSheetName As Worksheet
    Dim sTableName As String
    Dim tbl As ListObject
    Dim newrow As ListRow

    Dim divisionCode As String
    Dim workCenter As String
    Dim megaWorkCenter As String
    Dim sellingStyle As String
    Dim CorpBusinessHubQty As Integer
    Dim outputArr As Variant
    Dim toConsiderMsg As String
    
    Dim getResultsFromTTS As Variant
    Dim sellingFromTTS As String
    Dim cutDueDateFromTTS As String
    Dim sewDueDateFromTTS As String
    Dim DCDueDateFromTTS As String
    Dim OriginalWOFromTTS As Single
     
    Dim totalRows As Long
    Dim rowCount As Long
    
    Dim Process As String
    
    totalRows = Sheet7.Range("CorpBusinessUnitHub").Rows.Count
    rowCount = 2
    
    divisionCode = ""
    workCenter = ""
    megaWorkCenter = ""
    sellingStyle = ""
    CorpBusinessHubQty = 0
    toConsiderMsg = ""
    
    'GET THE QUANTITY OF BUCKETS IN THE TABLE.
    Process = "Attribution"
    BucketRow = getBucketInfoBySupplyHub("ANET", Process)
    
    If BucketRow = 0 Then
        Exit Sub
    End If
    
    setConnSQLMadm
    
    With MADMCommand
        Set .ActiveConnection = MADMConn
        .CommandType = adCmdText
        .CommandTimeout = 0
        
        .CommandText = "IF OBJECT_ID('tempdb..#TrasmittedManifest') IS NOT NULL DROP TABLE #TrasmittedManifest"
        .Execute
        .CommandText = "SELECT CD.WorkOrder, CD.MfgStyleCD, CD.MfgColorCD, CD.MfgSizeCD, CD.MfgSizeDESC, CD.MfgAttributeCD, CD.MfgRevisionNO, CD.PkgStyleCD, CD.PkgColorCD, CD.PkgSizeCD, CD.PkgSizeDESC, CD.PkgAttributeCD, CD.PkgRevisionNO, CD.SelStyleCD, CD.SelColorCD, CD.SelSizeCD, CD.SelSizeDESC, CD.SelAttributeCD, CD.SelRevisionNO, CD.ParentWorkOrder, CD.AssortmentParentWO, CD.FromPlantCD, CD.CutPlantCD, CD.SewPlantCD, CD.Lot, CD.ParentLot,  lastTrans.lastUpdate, SUM(round(Convert(Decimal(10, 4), CD.PiecesApplied) / Convert(Decimal(10, 4), 12), 4)) As Doz " & _
                       " into #TrasmittedManifest " & _
                       " FROM (SELECT Manifest, WorkOrder, MAX(ANETCreatedOnDate) AS lastUpdate FROM Manufacturing.dbo.ANETCOODetails WITH (NOLOCK) " & _
                       " WHERE (CONVERT(date, ANETCreatedOnDate) BETWEEN '" & initialDate & "' AND '" & endDate & "') AND (FromPlantCD IN('" & plants & "')) AND (HSCD IS NOT NULL) GROUP BY Manifest, WorkOrder) AS lastTrans INNER JOIN " & _
                       " Manufacturing.dbo.ANETCOODetails AS CD WITH (NOLOCK) ON lastTrans.Manifest = CD.Manifest AND lastTrans.WorkOrder = CD.WorkOrder AND lastTrans.lastUpdate = CD.ANETCreatedOnDate GROUP BY CD.WorkOrder, CD.MfgStyleCD, CD.MfgColorCD, CD.MfgSizeCD, CD.MfgSizeDESC, CD.MfgAttributeCD, CD.MfgRevisionNO, CD.PkgStyleCD, CD.PkgColorCD, CD.PkgSizeCD, CD.PkgSizeDESC, CD.PkgAttributeCD, CD.PkgRevisionNO, CD.SelStyleCD, CD.SelColorCD, CD.SelSizeCD, CD.SelSizeDESC, CD.SelAttributeCD, CD.SelRevisionNO, CD.ParentWorkOrder, CD.AssortmentParentWO, CD.FromPlantCD, CD.CutPlantCD, CD.SewPlantCD,CD.SewPlantCD, CD.Lot, CD.ParentLot,  lastTrans.lastUpdate"
        .Execute
            
            
        .CommandText = "IF OBJECT_ID('tempdb..#WOInfo') IS NOT NULL DROP TABLE #WOInfo"
        .Execute
        .CommandText = "select DISTINCT TM.WorkOrder ,OWT.GlobalWorkOrderID, CONVERT(INT,AWO.PriorityCD)  AS PriorityCD " & _
                       " INTO #WOInfo " & _
                       " from #TrasmittedManifest as TM " & _
                       " INNER JOIN Manufacturing.dbo.ANETWorkOrders AS AWO WITH (NOLOCK)  ON TM.WorkOrder = AWO.WorkOrder AND TM.MfgStyleCD = AWO.StyleCD AND TM.MfgColorCD = AWO.ColorCD AND TM.MfgSizeCD = AWO.SizeCD AND TM.MfgAttributeCD =AWO.AttributeCD " & _
                       " INNER JOIN dbo.ANETOriginalWorkOrderTrace as OWT WITH (NOLOCK)  ON AWO.GlobalWorkOrderID  = OWT.GlobalWorkOrderID"
        .Execute
            

        .CommandText = "IF OBJECT_ID('tempdb..#TransmittedWorkLots') IS NOT NULL DROP TABLE #TransmittedWorkLots"
        .Execute
        .CommandText = "select TM.WorkOrder , TM.ParentWorkOrder, TM.AssortmentParentWO, TM.FromPlantCD, TM.CutPlantCD,  " & _
                        " TM.SewPlantCD,SUM(TM.Doz) AS DOZ, OWT.OriginalWorkOrder, OWT.ApparelNETWorkLot, " & _
                        " TM.MfgStyleCD AS MfgStyle, TM.MfgColorCD AS MfgColor, TM.MfgSizeCD AS MfgSizeCD, TM.MfgSizeDesc AS MfgSizeDesc, TM.PkgStyleCD AS PkgStyle, TM.PkgColorCD AS PkgColor, " & _
                        " TM.PkgSizeCD , TM.PkgSizeDesc, TM.SelStyleCD AS SellStyle, TM.SelColorCD AS SellColor, TM.SelSizeCD AS SellSizeCD, " & _
                        " TM.SelSizeDESC AS SellSizeDesc, TM.CutPlantCD AS CutLoc, TM.SewPlantCD AS SewLoc, " & _
                        " ISNULL(TRY_CONVERT(int,OWT.OriginalWorkOrder),0) as WO_NUMBER " & _
                        " ,MAX(TM.lastUpdate) AS lastUpdate, " & _
                        " ISNULL(TM.AssortmentParentWO, TM.WorkOrder) As WorkLot " & _
                        " ,WO.PriorityCD " & _
                        " INTO #TransmittedWorkLots " & _
                        " from #TrasmittedManifest as TM " & _
                        " INNER JOIN #WOInfo AS WO on TM.WorkOrder =  wo.WorkOrder " & _
                        " LEFT JOIN dbo.ANETOriginalWorkOrderTrace as OWT WITH (NOLOCK)  ON WO.GlobalWorkOrderID  = OWT.GlobalWorkOrderID " & _
                        " GROUP BY TM.WorkOrder, TM.ParentWorkOrder, TM.AssortmentParentWO, TM.FromPlantCD, TM.CutPlantCD, " & _
                        " TM.SewPlantCD, OWT.OriginalWorkOrder, OWT.ApparelNETWorkLot, " & _
                        " TM.MfgStyleCD, TM.MfgColorCD, TM.MfgSizeCD, TM.MfgSizeDesc, TM.PkgStyleCD, TM.PkgColorCD, " & _
                        " TM.PkgSizeCD , TM.PkgSizeDesc, TM.SelStyleCD, TM.SelColorCD, TM.SelSizeCD, " & _
                        " TM.SelSizeDESC , TM.CutPlantCD, TM.SewPlantCD " & _
                        " ,WO.PriorityCD " & _
                        " ORDER BY TM.WorkOrder "
        .Execute
       
        
        
        .CommandText = "IF OBJECT_ID('tempdb..#TransmittedWorkLotsByDate') IS NOT NULL DROP TABLE #TransmittedWorkLotsByDate"
        .Execute
        .CommandText = "select TW.WorkOrder , TW.ParentWorkOrder, TW.AssortmentParentWO, TW.FromPlantCD, TW.CutPlantCD, TW.SewPlantCD,TW.DOZ, TW.OriginalWorkOrder, TW.ApparelNETWorkLot, TW.MfgStyle, TW.MfgColor, TW.MfgSizeCD, TW.MfgSizeDesc, TW.PkgStyle, TW.PkgColor,TW.PkgSizeCD, TW.PkgSizeDesc, TW.SellStyle, TW.SellColor, TW.SellSizeCD,TW.SellSizeDesc, TW.CutLoc, TW.SewLoc, TW.WO_NUMBER,TW.WorkLot, MAX(WOAD.ANETCreatedOnDate) as TrasmittedDate,  TW.PriorityCD " & _
                       " INTO #TransmittedWorkLotsByDate " & _
                       " from #TransmittedWorkLots as TW left join dbo.ANETWorkOrderActionDetails as WOAD on  TW.WorkLot = WOAD.WorkOrder " & _
                       " WHERE CONVERT(DATE,WOAD.ANETCreatedOnDate) <='" & endDate & "' AND WOAD.ActionCD = 'SH' AND WOAD.Quantity > 0 " & _
                       " GROUP BY TW.WorkOrder , TW.ParentWorkOrder, TW.AssortmentParentWO, TW.FromPlantCD, TW.CutPlantCD, " & _
                       " TW.SewPlantCD,TW.DOZ, TW.OriginalWorkOrder, TW.ApparelNETWorkLot, TW.MfgStyle, TW.MfgColor, TW.MfgSizeCD, TW.MfgSizeDesc, TW.PkgStyle, TW.PkgColor, " & _
                       " TW.PkgSizeCD, TW.PkgSizeDesc, TW.SellStyle, TW.SellColor, TW.SellSizeCD, " & _
                       " TW.SellSizeDesc, TW.CutLoc, TW.SewLoc, TW.WO_NUMBER,TW.WorkLot,TW.PriorityCD"
        .Execute
        
                
        
         ' CALL THE FUNCTION BY BUCKET SOURCE(ANET/TTS)
        For f = 1 To BucketRow
            initialTransaction = Sheet21.Cells(f + 1, 5)
            endTransaction = Sheet21.Cells(f + 1, 6)
            bucketId = Sheet21.Cells(f + 1, 3)
            bucketSource = Sheet21.Cells(f + 1, 8)
            InitialTransCondition = Sheet21.Cells(f + 1, 9)
            EndTransCondition = Sheet21.Cells(f + 1, 10)
            CutDueDateAnalysis = Sheet21.Cells(f + 1, 12)
            SewDueDateAnalysis = Sheet21.Cells(f + 1, 13)
            DCDueDateAnalysis = Sheet21.Cells(f + 1, 14)
            BucketDescription = Sheet21.Cells(f + 1, 4)
            BucketGoalInDays = Sheet21.Cells(f + 1, 11)
            
            
            .CommandText = "IF OBJECT_ID('tempdb..#InitialTrans') IS NOT NULL DROP TABLE #InitialTrans"
            .Execute
            .CommandText = "SELECT * " & _
                           " into #InitialTrans FROM ( " & _
                           " SELECT TWL.WorkOrder,ActionCD," & InitialTransCondition & "(ANETCreatedOnDate) AS DT, 'wo' as obs  FROM #TransmittedWorkLotsByDate  AS TWL INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.WorkOrder = WOD.WorkOrder WHERE ActionCD ='" & initialTransaction & "' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate GROUP BY TWL.WorkOrder,ActionCD " & _
                           " Union " & _
                           " SELECT TWL.WorkOrder,ActionCD," & InitialTransCondition & "(ANETCreatedOnDate) AS DT, 'assort' as obs  FROM #TransmittedWorkLotsByDate AS TWL INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.AssortmentParentWO = WOD.WorkOrder WHERE ActionCD ='" & initialTransaction & "' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate GROUP BY TWL.WorkOrder,ActionCD " & _
                           " Union " & _
                           " SELECT TWL.WorkOrder,ActionCD," & InitialTransCondition & "(ANETCreatedOnDate) AS DT, 'Pare' as obs  FROM #TransmittedWorkLotsByDate AS TWL INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.ParentWorkOrder = WOD.WorkOrder WHERE ActionCD ='" & initialTransaction & "' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate GROUP BY TWL.WorkOrder,ActionCD) " & _
                           " AS InitialTrans "
            .Execute
        
        
            .CommandText = "IF OBJECT_ID('tempdb..#EndTrans') IS NOT NULL DROP TABLE #EndTrans"
            .Execute
                
            If endTransaction = "SH" Then
                .CommandText = "SELECT WorkOrder, 'SH' AS ActionCD, Max(TrasmittedDate) as DT, 'wo' AS obs into #EndTrans FROM  #TransmittedWorkLotsByDate group by WorkOrder"
            Else
                .CommandText = "SELECT * " & _
                               " into #EndTrans FROM ( " & _
                                " SELECT TWL.WorkOrder,ActionCD," & EndTransCondition & "(ANETCreatedOnDate) AS DT, 'wo' as obs  FROM #TransmittedWorkLotsByDate AS TWL INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.WorkOrder = WOD.WorkOrder WHERE ActionCD ='" & endTransaction & "' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate GROUP BY TWL.WorkOrder,ActionCD " & _
                                " Union " & _
                                " SELECT TWL.WorkOrder,ActionCD," & EndTransCondition & "(ANETCreatedOnDate) AS DT, 'assort' as obs  FROM #TransmittedWorkLotsByDate AS TWL INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.AssortmentParentWO = WOD.WorkOrder WHERE ActionCD ='" & endTransaction & "' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate GROUP BY TWL.WorkOrder,ActionCD " & _
                                " Union " & _
                                " SELECT TWL.WorkOrder,ActionCD," & EndTransCondition & "(ANETCreatedOnDate) AS DT, 'Pare' as obs  FROM #TransmittedWorkLotsByDate AS TWL INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.ParentWorkOrder = WOD.WorkOrder WHERE ActionCD ='" & endTransaction & "' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate GROUP BY TWL.WorkOrder,ActionCD) " & _
                                " AS EndTrans"
            End If
            .Execute
        
        
            .CommandText = "IF OBJECT_ID('tempdb..#TransactionInfo') IS NOT NULL DROP TABLE #TransactionInfo"
            .Execute
            .CommandText = "select DISTINCT *, " & _
                            " 'ID' AS InitialTrans, " & _
                            " ISNULL((SELECT dt FROM #InitialTrans WHERE WorkOrder = TWL.WorkOrder AND obs ='WO'), " & _
                            " ISNULL((SELECT dt FROM #InitialTrans WHERE WorkOrder = TWL.WorkOrder AND obs ='assort'),(SELECT dt FROM #InitialTrans WHERE WorkOrder = TWL.WorkOrder AND obs ='Pare'))) AS InitialDate, " & _
                            " 'FQ' AS EndTrans, " & _
                            " ISNULL((SELECT dt FROM #EndTrans  WHERE WorkOrder = TWL.WorkOrder AND obs ='WO'), " & _
                            " ISNULL((SELECT dt FROM #EndTrans WHERE WorkOrder = TWL.WorkOrder AND obs ='assort'),(SELECT dt FROM #EndTrans WHERE WorkOrder = TWL.WorkOrder AND obs ='Pare'))) AS EndDate, " & _
                            " TWL.DOZ  AS Quantity " & _
                            " into #TransactionInfo " & _
                            " from #TransmittedWorkLots as TWL"
                
            .Execute
                
            .CommandText = "IF OBJECT_ID('tempdb..#CurrentWO') IS NOT NULL DROP TABLE #CurrentWO"
            .Execute
            .CommandText = "SELECT  dbo.TTSCutOrderRoutingDetail.WorkOrderNumber, MAX(dbo.TTSCutOrderRoutingDetail.CutOrderSequence) AS CutOrderSequence, RIGHT('000000' + CONVERT(varchar, dbo.TTSCutOrderRoutingDetail.WorkOrderNumber), 6)AS CurrentWorkOrderNumber, SellStyle , MfgStyle, MfgColor, MfgSizeDesc, SizeDescription " & _
                               " INTO #CurrentWO " & _
                               " FROM  dbo.TTSCutOrderRoutingDetail WITH (NOLOCK) INNER JOIN #TransactionInfo AS TransmittesdOrder ON dbo.TTSCutOrderRoutingDetail.WorkOrderNumber = TransmittesdOrder.WO_NUMBER AND dbo.TTSCutOrderRoutingDetail.GarmentStyle = TransmittesdOrder.MfgStyle AND dbo.TTSCutOrderRoutingDetail.GarmentColor = TransmittesdOrder.MfgColor  And TransmittesdOrder.MfgSizeCD = dbo.TTSCutOrderRoutingDetail.SizeDescription GROUP BY dbo.TTSCutOrderRoutingDetail.WorkOrderNumber, SellStyle,MfgStyle,MfgColor,MfgSizeDesc,SizeDescription"
            .Execute
                  
                  
            .CommandText = "IF OBJECT_ID('tempdb..#LeadTimeInfo') IS NOT NULL DROP TABLE #LeadTimeInfo"
            .Execute
            .CommandText = "select DISTINCT TI.CutLoc as CutPlant,TI.SewPlantCD as SewPlant  " & _
                           " , CONVERT(varchar, CONVERT(date, CONVERT(VARCHAR(8), CORD.CutDueDate)), 101) AS CutDueDate " & _
                           " ,CONVERT(varchar, CONVERT(date, CONVERT(VARCHAR(8), CORD.SeDueDate)), 101) AS SewDueDate " & _
                           " ,CONVERT(varchar, CONVERT(date, CONVERT(VARCHAR(8), CORD.DCDueDate)), 101) AS DCDueDate " & _
                           " ,TI.OriginalWorkOrder AS WorkOrder, " & _
                           " CORD.OriginalTTSWO AS OriginalWO, " & _
                           " TI.WorkOrder AS WorkLot, " & _
                           " ISNULL(TI.PriorityCD ,ISNULL(CORD.Priority,0)) AS [Priority], " & _
                           " TI.SellStyle AS SellingStyle " & _
                           " ,TI.MfgStyle AS MFGStyle " & _
                           " ,TI.MfgColor AS MFGColor " & _
                           " ,TI.MfgSizeDesc AS MFGSize, " & _
                           " '" & initialTransaction & "' AS InitialTransCode " & _
                           " ,TI.InitialDate, '" & endTransaction & "' AS EndTransCode, TI.EndDate " & _
                           " ,TI.Quantity AS Doz " & _
                           " ,round(CONVERT(decimal(30,2),DATEDIFF (Second, TI.InitialDate, TI.EndDate)) / CONVERT(decimal(30,2), 86400),2) AS LTDays " & _
                           " INTO #LeadTimeInfo " & _
                           " From " & _
                           " #CurrentWO as co INNER JOIN " & _
                           " Manufacturing.dbo.TTSCutOrderRoutingDetail AS CORD WITH (NOLOCK) ON CO.WorkOrderNumber  = CORD.WorkOrderNumber " & _
                           " AND  CO.CutOrderSequence = CORD.CutOrderSequence " & _
                           " RIGHT JOIN #TransactionInfo as TI ON CORD.WorkOrderNumber = TI.WO_NUMBER AND " & _
                           " CORD.GarmentStyle = TI.mfgStyle And CORD.GarmentColor = TI.MfgColor " & _
                           " And CORD.SizeDescription = TI.MfgSizeCD " & _
                           " ORDER BY WorkLot"
                .Execute
                
                sQuery = "SELECT ISNULL(CutPlant,'') as CutPlant, SewPlant , ISNULL(CutDueDate,'09/09/1999') as CutDueDate, ISNULL(SewDueDate,'09/09/1999') as SewDueDate, ISNULL(DCDueDate,'09/09/1999') AS DCDueDate, " & _
                            " ISNULL(TRY_CONVERT(BIGINT,WorkOrder),0) as WorkOrder, " & _
                            " convert(BIGINT, isnull(OriginalWO,ISNULL(TRY_CONVERT(BIGINT,OriginalWO),0))) as OriginalWO, " & _
                            " isnull(WorkLot,'') as WorkLot, [Priority], isnull(SellingStyle,'') as SellingStyle, isnull(MFGStyle,'') as MFGStyle, isnull(MFGColor,'') as MFGColor, isnull(MFGSize,'') as MFGSize, isnull(InitialTransCode,'') as InitialTransCode, ISNULL(InitialDate,'09/09/1999 15:46:30') AS  InitialDate " & _
                            " ,isnull(EndTransCode,'') as EndTransCode, ISNULL(EndDate,'09/09/1999 15:46:30') AS EndDate, Doz " & _
                            " ,ISNULL(LTDays,0) AS LTDays " & _
                            " ,CASE WHEN InitialDate is null OR EndDate is null THEN 'Exclude' ELSE CASE WHEN LTDays < 0 THEN 'Exclude' ELSE 'Include' END END AS [ToConsider?] " & _
                            " ,P.[PlantDESC] AS SewPlantName " & _
                            " FROM #LeadTimeInfo as LT " & _
                            " LEFT JOIN Manufacturing.dbo.ANETFacilities as P with (nolock)  on LT.SewPlant = P.PlantCD " & _
                            " ORDER BY SellingStyle"

                If MADMrs.State = adStateOpen Then
                    MADMrs.Close
                End If
                
                ' Execute query using recordset object.
                MADMrs.CursorType = adOpenForwardOnly
                MADMrs.LockType = adLockReadOnly
                MADMrs.CursorLocation = adUseClient
                MADMConn.CommandTimeout = 0
                MADMrs.Open sQuery, MADMConn, adOpenKeyset, adLockOptimistic
                                                   
                                                   
                'Define Variable
                sTableName = "HistoryTransaction"
                                
                'Define WorkSheet object
                Set oSheetName = Sheets("TransactionInfo")
                                                 
                'Define Table Object
                Set tbl = oSheetName.ListObjects(sTableName)

                If MADMrs.RecordCount > 0 Then
                    Do While Not MADMrs.EOF
                    
                        'GET THE SELLING STYLE FROM TTS.
                        If IsNumeric(MADMrs.Fields("WorkOrder").Value) = False Then
                            sellingFromTTS = Trim(MADMrs.Fields("SellingStyle").Value)
                            cutDueDateFromTTS = MADMrs.Fields("CutDueDate").Value
                            sewDueDateFromTTS = MADMrs.Fields("SewDueDate").Value
                            DCDueDateFromTTS = MADMrs.Fields("DCDueDate").Value
                            OriginalWOFromTTS = MADMrs.Fields("OriginalWO").Value
                        Else
                            getResultsFromTTS = getSellingStyleFromLocalSheet(MADMrs.Fields("WorkOrder").Value, Trim(MADMrs.Fields("SellingStyle").Value), MADMrs.Fields("CutDueDate").Value, MADMrs.Fields("SewDueDate").Value, MADMrs.Fields("DCDueDate").Value, MADMrs.Fields("OriginalWO").Value)
                            sellingFromTTS = getResultsFromTTS(1)
                            cutDueDateFromTTS = getResultsFromTTS(2)
                            sewDueDateFromTTS = getResultsFromTTS(3)
                            DCDueDateFromTTS = getResultsFromTTS(4)
                            OriginalWOFromTTS = getResultsFromTTS(5)
                        End If
                                                
                                                
                        'GET THE CORP BUSINESS UNIT AND WORKCENTER CODE
                        If sellingStyle <> sellingFromTTS Then
                            outputArr = getCorpBusinessUnit_WorkCenter_MegaWorkCenter(MADMrs.Fields("SellingStyle").Value)
                            divisionCode = outputArr(1)
                            workCenter = outputArr(2)
                            megaWorkCenter = outputArr(3)
                            
                            CorpBusinessHubQty = 0
                            
                            For rowCount = 2 To totalRows + 1
                                If Sheet7.Cells(rowCount, 1).Value = SupplyChainHub And Sheet7.Cells(rowCount, 2).Value = divisionCode Then
                                    CorpBusinessHubQty = 1
                                    Exit For
                                End If
                            Next rowCount
                        End If
                        
                                                
                        'INCLUDE OR EXCLUDE FROM THE DATA BY DIVISION
                        If divisionCode = "N/A" Or CorpBusinessHubQty > 0 Then
                            toConsiderMsg = MADMrs.Fields("ToConsider?").Value
                        Else
                            toConsiderMsg = "Exclude"
                        End If
                            
                        Set newrow = tbl.ListRows.Add
                        With newrow
                            .Range(1) = WeekSelected
                            .Range(2) = "'" & plants
                            .Range(3) = plantName
                            .Range(4) = "'" & Trim(MADMrs.Fields("CutPlant").Value)
                            .Range(5) = "'" & Trim(MADMrs.Fields("SewPlant").Value)
                            .Range(6) = cutDueDateFromTTS
                            .Range(7) = sewDueDateFromTTS
                            .Range(8) = DCDueDateFromTTS
                            .Range(9) = MADMrs.Fields("WorkOrder").Value
                            .Range(10) = OriginalWOFromTTS
                            .Range(11) = "'" & Trim(MADMrs.Fields("Priority").Value)
                            .Range(12) = "'" & Trim(MADMrs.Fields("WorkLot").Value)
                            .Range(13) = "'" & sellingFromTTS
                            .Range(14) = "'" & Trim(MADMrs.Fields("MFGStyle").Value)
                            .Range(15) = "'" & Trim(MADMrs.Fields("MFGColor").Value)
                            .Range(16) = "'" & Trim(MADMrs.Fields("MFGSize").Value)
                            .Range(17) = divisionCode
                            .Range(18) = workCenter
                            .Range(19) = megaWorkCenter
                            .Range(20) = bucketId
                            .Range(21) = "'" & MADMrs.Fields("InitialTransCode").Value
                            .Range(22) = MADMrs.Fields("InitialDate").Value
                            .Range(23) = "'" & MADMrs.Fields("EndTransCode").Value
                            .Range(24) = MADMrs.Fields("EndDate").Value
                            .Range(25) = MADMrs.Fields("Doz").Value
                            .Range(26) = MADMrs.Fields("LTDays").Value
                            .Range(27) = toConsiderMsg
                            .Range(28) = CutDueDateAnalysis
                            .Range(29) = SewDueDateAnalysis
                            .Range(30) = DCDueDateAnalysis
                            .Range(31) = SupplyChainHub
                            .Range(32) = BucketDescription
                            .Range(33) = plantCategory
                            .Range(34) = reportType
                            .Range(35) = Process
                        End With
                        sellingStyle = sellingFromTTS
                        MADMrs.MoveNext
                    Loop
                End If
        Next f
    End With
End Sub






'FUNCTION TO OBTAIN THE WORKLOTS TRANSACTIONS FROM MADM FOR THE SEWING PROCESS
Public Sub GetInformationFromMADN_Sewing()
    Dim oSheetName As Worksheet
    Dim sTableName As String
    Dim tbl As ListObject
    Dim newrow As ListRow

    Dim divisionCode As String
    Dim workCenter As String
    Dim megaWorkCenter As String
    Dim sellingStyle As String
    Dim CorpBusinessHubQty As Integer
    Dim outputArr As Variant
    Dim toConsiderMsg As String
    
    Dim getResultsFromTTS As Variant
    Dim sellingFromTTS As String
    Dim cutDueDateFromTTS As String
    Dim sewDueDateFromTTS As String
    Dim DCDueDateFromTTS As String
    Dim OriginalWOFromTTS As Single
     
    Dim totalRows As Long
    Dim rowCount As Long
    Dim Process As String
      
      
    totalRows = Sheet7.Range("CorpBusinessUnitHub").Rows.Count
    rowCount = 2
    
    divisionCode = ""
    workCenter = ""
    megaWorkCenter = ""
    sellingStyle = ""
    CorpBusinessHubQty = 0
    toConsiderMsg = ""
    
    'GET THE QUANTITY OF BUCKETS IN THE TABLE.
    Process = "Sew"
    BucketRow = getBucketInfoBySupplyHub("ANET", Process)

    If BucketRow = 0 Then
        Exit Sub
    End If
    
    setConnSQLMadm
    
    With MADMCommand
        Set .ActiveConnection = MADMConn
        .CommandType = adCmdText
        .CommandTimeout = 0
        
        '---#AttributionLotsWithWO---
        .CommandText = "IF OBJECT_ID('tempdb..#AttributionLotsWithWO') IS NOT NULL DROP TABLE #AttributionLotsWithWO"
        .Execute
        .CommandText = "SELECT distinct ACD.WorkOrder,ACD.MfgStyleCD,ACD.MfgColorCD,ACD.MfgSizeCD,ACD.MfgAttributeCD " & _
                    " ,MfgRevisionCD  ,ACD.SewPlantCD ,OWT.OriginalWorkOrder, convert(varchar(3), ACD.MfgColorCD) as MfgColorCDWithThreeCharacters " & _
                    " into #AttributionLotsWithWO " & _
                    " FROM   Manufacturing.dbo.ANETCOODetails AS ACD  WITH (NOLOCK) " & _
                    " INNER JOIN Manufacturing.dbo.ANETWorkOrders AS AWO  WITH (NOLOCK) ON  AWO.WorkOrder = ACD.WorkOrder " & _
                    " INNER JOIN Manufacturing.dbo.ANETOriginalWorkOrderTrace as OWT WITH (NOLOCK) ON OWT.GlobalWorkOrderID =AWO.GlobalWorkOrderID AND OWT.MfgStyle  = ACD.MfgStyleCD AND convert(varchar(3),OWT.MfgColor)  = convert(varchar(3),ACD.MfgColorCD) " & _
                    " AND OWT.MfgSizeCD  = ACD.MfgSizeCD AND OWT.MfgAttributeCD  =ACD.MfgAttributeCD   AND OWT.SewLoc = ACD.SewPlantCD " & _
                    " WHERE (CONVERT(date, ACD.ANETCreatedOnDate) BETWEEN '" & initialDate & "' AND '" & endDate & "') AND (ACD.FromPlantCD IN('" & plants & "')) AND (ACD.HSCD IS NOT NULL)"
        .Execute
        '---
        
        
        '---#AttributionSewMergeLots---
        .CommandText = "IF OBJECT_ID('tempdb..#AttributionSewMergeLots') IS NOT NULL DROP TABLE #AttributionSewMergeLots"
        .Execute
        .CommandText = "SELECT distinct ACD.*, OWT.ApparelNETWorkLot " & _
                        " into #AttributionSewMergeLots " & _
                        " FROM   #AttributionLotsWithWO AS ACD " & _
                        " INNER JOIN Manufacturing.dbo.ANETOriginalWorkOrderTrace as OWT WITH (NOLOCK) ON  OWT.OriginalWorkOrder  = ACD.OriginalWorkOrder AND  OWT.MfgStyle  = ACD.MfgStyleCD AND convert(varchar(3),OWT.MfgColor)  = convert(varchar(3),ACD.MfgColorCD) AND OWT.MfgSizeCD  = ACD.MfgSizeCD AND OWT.MfgAttributeCD  =ACD.MfgAttributeCD   AND ACD.SewPlantCD = OWT.SewLoc " & _
                        " WHERE OWT.AttrSewLoc = '   '"
                        
        .Execute
        '---
        
        
        
        '---#lastTrans---
        .CommandText = "IF OBJECT_ID('tempdb..#lastTrans') IS NOT NULL DROP TABLE #lastTrans"
        .Execute
        .CommandText = "SELECT ACD.Manifest, ACD.WorkOrder, MAX(ACD.ANETCreatedOnDate) AS lastUpdate " & _
                        " INTO #lastTrans " & _
                        " FROM Manufacturing.dbo.ANETCOODetails AS ACD  WITH (NOLOCK) " & _
                        " INNER JOIN  #AttributionSewMergeLots AS M on ACD.WorkOrder = m.ApparelNETWorkLot  and ACD.FromPlantCD = m.SewPlantCD " & _
                        " WHERE (CONVERT(date, ANETCreatedOnDate) <= '" & endDate & "') " & _
                        " AND (HSCD IS NOT NULL) GROUP BY Manifest, ACD.WorkOrder"
        .Execute
        '---
        
        
        '---#TrasmittedManifest---
        .CommandText = "IF OBJECT_ID('tempdb..#TrasmittedManifest') IS NOT NULL DROP TABLE #TrasmittedManifest"
        .Execute
        .CommandText = "SELECT CD.WorkOrder, CD.MfgStyleCD, CD.MfgColorCD, CD.MfgSizeCD, CD.MfgSizeDESC, CD.MfgAttributeCD, CD.MfgRevisionNO, CD.PkgStyleCD, CD.PkgColorCD, CD.PkgSizeCD, CD.PkgSizeDESC, CD.PkgAttributeCD, CD.PkgRevisionNO, CD.SelStyleCD, CD.SelColorCD, CD.SelSizeCD, CD.SelSizeDESC, CD.SelAttributeCD, CD.SelRevisionNO, CD.ParentWorkOrder, CD.AssortmentParentWO, CD.FromPlantCD, CD.CutPlantCD, CD.SewPlantCD, CD.Lot, CD.ParentLot,  lastTrans.lastUpdate, SUM(round(Convert(Decimal(10, 4), CD.PiecesApplied) / Convert(Decimal(10, 4), 12), 4)) As Doz " & _
                        " ,convert(varchar(3), CD.MfgColorCD) as MfgColorCDWithThreeCharacters " & _
                        " into #TrasmittedManifest " & _
                        " FROM #lastTrans AS lastTrans INNER JOIN " & _
                        " Manufacturing.dbo.ANETCOODetails AS CD WITH (NOLOCK) ON lastTrans.Manifest = CD.Manifest AND lastTrans.WorkOrder = CD.WorkOrder AND lastTrans.lastUpdate = CD.ANETCreatedOnDate " & _
                        " GROUP BY CD.WorkOrder, CD.MfgStyleCD, CD.MfgColorCD, CD.MfgSizeCD, CD.MfgSizeDESC, CD.MfgAttributeCD, CD.MfgRevisionNO, CD.PkgStyleCD, CD.PkgColorCD, CD.PkgSizeCD, CD.PkgSizeDESC, CD.PkgAttributeCD, CD.PkgRevisionNO, CD.SelStyleCD, CD.SelColorCD, CD.SelSizeCD, CD.SelSizeDESC, CD.SelAttributeCD, CD.SelRevisionNO, CD.ParentWorkOrder, CD.AssortmentParentWO, CD.FromPlantCD, CD.CutPlantCD, CD.SewPlantCD,CD.SewPlantCD, CD.Lot, CD.ParentLot,  lastTrans.lastUpdate"
        .Execute
        '---
         
         
         '---#WOInfo---
        .CommandText = "IF OBJECT_ID('tempdb..#WOInfo') IS NOT NULL DROP TABLE #WOInfo"
        .Execute
        .CommandText = "select DISTINCT TM.WorkOrder ,OWT.GlobalWorkOrderID, CONVERT(INT,AWO.PriorityCD)  AS PriorityCD " & _
                        " INTO #WOInfo " & _
                        " from #TrasmittedManifest as TM " & _
                        " INNER JOIN Manufacturing.dbo.ANETWorkOrders AS AWO WITH (NOLOCK)  ON TM.WorkOrder = AWO.WorkOrder AND TM.MfgStyleCD = AWO.StyleCD AND convert(varchar(3),TM.MfgColorCD) = convert(varchar(3),AWO.ColorCD) AND TM.MfgSizeCD = AWO.SizeCD AND TM.MfgAttributeCD =AWO.AttributeCD " & _
                        " INNER JOIN Manufacturing.dbo.ANETOriginalWorkOrderTrace as OWT WITH (NOLOCK)  ON AWO.GlobalWorkOrderID  = OWT.GlobalWorkOrderID"
        .Execute
        '---


         '---#TransmittedWorkLots---
        .CommandText = "IF OBJECT_ID('tempdb..#TransmittedWorkLots') IS NOT NULL DROP TABLE #TransmittedWorkLots"
        .Execute
        .CommandText = "select TM.WorkOrder , TM.ParentWorkOrder, TM.AssortmentParentWO, TM.FromPlantCD, TM.CutPlantCD,  " & _
                        " TM.SewPlantCD,SUM(TM.Doz) AS DOZ, OWT.OriginalWorkOrder, OWT.ApparelNETWorkLot, " & _
                        " TM.MfgStyleCD AS MfgStyle, TM.MfgColorCD AS MfgColor, TM.MfgSizeCD AS MfgSizeCD, TM.MfgSizeDesc AS MfgSizeDesc, TM.PkgStyleCD AS PkgStyle, TM.PkgColorCD AS PkgColor, " & _
                        " TM.PkgSizeCD , TM.PkgSizeDesc, TM.SelStyleCD AS SellStyle, TM.SelColorCD AS SellColor, TM.SelSizeCD AS SellSizeCD, " & _
                        " TM.SelSizeDESC AS SellSizeDesc, TM.CutPlantCD AS CutLoc, TM.SewPlantCD AS SewLoc, " & _
                        " ISNULL(TRY_CONVERT(int,OWT.OriginalWorkOrder),0) as WO_NUMBER " & _
                        " ,MAX(TM.lastUpdate) AS lastUpdate, " & _
                        " ISNULL(TM.AssortmentParentWO, TM.WorkOrder) As WorkLot " & _
                        " ,WO.PriorityCD " & _
                        " INTO #TransmittedWorkLots " & _
                        " from #TrasmittedManifest as TM " & _
                        " INNER JOIN #WOInfo AS WO on TM.WorkOrder =  wo.WorkOrder " & _
                        " LEFT JOIN Manufacturing .dbo.ANETOriginalWorkOrderTrace as OWT WITH (NOLOCK)  ON WO.GlobalWorkOrderID  = OWT.GlobalWorkOrderID " & _
                        " GROUP BY TM.WorkOrder, TM.ParentWorkOrder, TM.AssortmentParentWO, TM.FromPlantCD, TM.CutPlantCD, " & _
                        " TM.SewPlantCD, OWT.OriginalWorkOrder, OWT.ApparelNETWorkLot, " & _
                        " TM.MfgStyleCD, TM.MfgColorCD, TM.MfgSizeCD, TM.MfgSizeDesc, TM.PkgStyleCD, TM.PkgColorCD, " & _
                        " TM.PkgSizeCD , TM.PkgSizeDesc, TM.SelStyleCD, TM.SelColorCD, TM.SelSizeCD, " & _
                        " TM.SelSizeDESC , TM.CutPlantCD, TM.SewPlantCD " & _
                        " ,WO.PriorityCD " & _
                        " ORDER BY TM.WorkOrder "
        .Execute
       '---
        
        
        
        '---#TransmittedWorkLotsByDate---
        .CommandText = "IF OBJECT_ID('tempdb..#TransmittedWorkLotsByDate') IS NOT NULL DROP TABLE #TransmittedWorkLotsByDate"
        .Execute
         .CommandText = "select TW.WorkOrder , TW.ParentWorkOrder, TW.AssortmentParentWO, TW.FromPlantCD, TW.CutPlantCD, TW.SewPlantCD,TW.DOZ, TW.OriginalWorkOrder, TW.ApparelNETWorkLot, TW.MfgStyle, TW.MfgColor, TW.MfgSizeCD, TW.MfgSizeDesc, TW.PkgStyle, TW.PkgColor,TW.PkgSizeCD, TW.PkgSizeDesc, TW.SellStyle, TW.SellColor, TW.SellSizeCD,TW.SellSizeDesc, TW.CutLoc, TW.SewLoc, TW.WO_NUMBER,TW.WorkLot, MAX(WOAD.ANETCreatedOnDate) as TrasmittedDate " & _
                    " INTO #TransmittedWorkLotsByDate " & _
                    " from #TransmittedWorkLots as TW left join Manufacturing.dbo.ANETWorkOrderActionDetails as WOAD on  TW.WorkLot = WOAD.WorkOrder " & _
                    " WHERE CONVERT(DATE,WOAD.ANETCreatedOnDate) <='" & endDate & "'  AND WOAD.ActionCD = 'SH' AND WOAD.Quantity > 0 " & _
                    " GROUP BY TW.WorkOrder , TW.ParentWorkOrder, TW.AssortmentParentWO, TW.FromPlantCD, TW.CutPlantCD, " & _
                    " TW.SewPlantCD,TW.DOZ, TW.OriginalWorkOrder, TW.ApparelNETWorkLot, TW.MfgStyle, TW.MfgColor, TW.MfgSizeCD, TW.MfgSizeDesc, TW.PkgStyle, TW.PkgColor, " & _
                    " TW.PkgSizeCD, TW.PkgSizeDesc, TW.SellStyle, TW.SellColor, TW.SellSizeCD, " & _
                    " TW.SellSizeDesc, TW.CutLoc, TW.SewLoc, TW.WO_NUMBER,TW.WorkLot"
        .Execute
        '---
                
        
         ' CALL THE FUNCTION BY BUCKET SOURCE(ANET/TTS)
        For f = 1 To BucketRow
            initialTransaction = Sheet21.Cells(f + 1, 5)
            endTransaction = Sheet21.Cells(f + 1, 6)
            bucketId = Sheet21.Cells(f + 1, 3)
            bucketSource = Sheet21.Cells(f + 1, 8)
            InitialTransCondition = Sheet21.Cells(f + 1, 9)
            EndTransCondition = Sheet21.Cells(f + 1, 10)
            CutDueDateAnalysis = Sheet21.Cells(f + 1, 12)
            SewDueDateAnalysis = Sheet21.Cells(f + 1, 13)
            DCDueDateAnalysis = Sheet21.Cells(f + 1, 14)
            BucketDescription = Sheet21.Cells(f + 1, 4)
            BucketGoalInDays = Sheet21.Cells(f + 1, 11)
            
            '---#InitialTrans---
            .CommandText = "IF OBJECT_ID('tempdb..#InitialTrans') IS NOT NULL DROP TABLE #InitialTrans"
            .Execute
             .CommandText = "SELECT * " & _
                           " into #InitialTrans FROM ( " & _
                           " SELECT TWL.WorkOrder,ActionCD," & InitialTransCondition & "(ANETCreatedOnDate) AS DT, 'wo' as obs  FROM #TransmittedWorkLotsByDate  AS TWL INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.WorkOrder = WOD.WorkOrder WHERE ActionCD ='" & initialTransaction & "' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate GROUP BY TWL.WorkOrder,ActionCD " & _
                           " Union " & _
                           " SELECT TWL.WorkOrder,ActionCD," & InitialTransCondition & "(ANETCreatedOnDate) AS DT, 'assort' as obs  FROM #TransmittedWorkLotsByDate AS TWL INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.AssortmentParentWO = WOD.WorkOrder WHERE ActionCD ='" & initialTransaction & "' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate GROUP BY TWL.WorkOrder,ActionCD " & _
                           " Union " & _
                           " SELECT TWL.WorkOrder,ActionCD," & InitialTransCondition & "(ANETCreatedOnDate) AS DT, 'Pare' as obs  FROM #TransmittedWorkLotsByDate AS TWL INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.ParentWorkOrder = WOD.WorkOrder WHERE ActionCD ='" & initialTransaction & "' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate GROUP BY TWL.WorkOrder,ActionCD) " & _
                           " AS InitialTrans "
            .Execute
            '---
        
            '---#EndTrans---
            .CommandText = "IF OBJECT_ID('tempdb..#EndTrans') IS NOT NULL DROP TABLE #EndTrans"
            .Execute
                
            If endTransaction = "SH" Then
                .CommandText = "SELECT WorkOrder, 'SH' AS ActionCD, Max(TrasmittedDate) as DT, 'wo' AS obs into #EndTrans FROM  #TransmittedWorkLotsByDate group by WorkOrder"
            Else
                .CommandText = "SELECT * " & _
                               " into #EndTrans FROM ( " & _
                                " SELECT TWL.WorkOrder,ActionCD," & EndTransCondition & "(ANETCreatedOnDate) AS DT, 'wo' as obs  FROM #TransmittedWorkLotsByDate AS TWL INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.WorkOrder = WOD.WorkOrder WHERE ActionCD ='" & endTransaction & "' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate GROUP BY TWL.WorkOrder,ActionCD " & _
                                " Union " & _
                                " SELECT TWL.WorkOrder,ActionCD," & EndTransCondition & "(ANETCreatedOnDate) AS DT, 'assort' as obs  FROM #TransmittedWorkLotsByDate AS TWL INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.AssortmentParentWO = WOD.WorkOrder WHERE ActionCD ='" & endTransaction & "' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate GROUP BY TWL.WorkOrder,ActionCD " & _
                                " Union " & _
                                " SELECT TWL.WorkOrder,ActionCD," & EndTransCondition & "(ANETCreatedOnDate) AS DT, 'Pare' as obs  FROM #TransmittedWorkLotsByDate AS TWL INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.ParentWorkOrder = WOD.WorkOrder WHERE ActionCD ='" & endTransaction & "' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate GROUP BY TWL.WorkOrder,ActionCD) " & _
                                " AS EndTrans"
            End If
            .Execute
             '---
        
            '---#TransactionInfo---
            .CommandText = "IF OBJECT_ID('tempdb..#TransactionInfo') IS NOT NULL DROP TABLE #TransactionInfo"
            .Execute
            .CommandText = "select DISTINCT *, " & _
                            " 'ID' AS InitialTrans, " & _
                            " ISNULL((SELECT dt FROM #InitialTrans WHERE WorkOrder = TWL.WorkOrder AND obs ='WO'), " & _
                            " ISNULL((SELECT dt FROM #InitialTrans WHERE WorkOrder = TWL.WorkOrder AND obs ='assort'),(SELECT dt FROM #InitialTrans WHERE WorkOrder = TWL.WorkOrder AND obs ='Pare'))) AS InitialDate, " & _
                            " 'FQ' AS EndTrans, " & _
                            " ISNULL((SELECT dt FROM #EndTrans  WHERE WorkOrder = TWL.WorkOrder AND obs ='WO'), " & _
                            " ISNULL((SELECT dt FROM #EndTrans WHERE WorkOrder = TWL.WorkOrder AND obs ='assort'),(SELECT dt FROM #EndTrans WHERE WorkOrder = TWL.WorkOrder AND obs ='Pare'))) AS EndDate, " & _
                            " TWL.DOZ  AS Quantity " & _
                            " into #TransactionInfo " & _
                            " from #TransmittedWorkLots as TWL"
                
            .Execute
            '---
              
            '---#CurrentWO---
            .CommandText = "IF OBJECT_ID('tempdb..#CurrentWO') IS NOT NULL DROP TABLE #CurrentWO"
            .Execute
            .CommandText = "SELECT  Manufacturing.dbo.TTSCutOrderRoutingDetail.WorkOrderNumber, MAX(Manufacturing.dbo.TTSCutOrderRoutingDetail.CutOrderSequence) AS CutOrderSequence, RIGHT('000000' + CONVERT(varchar, Manufacturing.dbo.TTSCutOrderRoutingDetail.WorkOrderNumber), 6)AS CurrentWorkOrderNumber, SellStyle , MfgStyle, MfgColor, MfgSizeDesc, SizeDescription  " & _
                            " INTO #CurrentWO " & _
                            " FROM  Manufacturing.dbo.TTSCutOrderRoutingDetail WITH (NOLOCK) INNER JOIN #TransactionInfo AS TransmittesdOrder ON Manufacturing.dbo.TTSCutOrderRoutingDetail.WorkOrderNumber = TransmittesdOrder.WO_NUMBER AND Manufacturing.dbo.TTSCutOrderRoutingDetail.GarmentStyle = TransmittesdOrder.MfgStyle AND convert(varchar(3),Manufacturing.dbo.TTSCutOrderRoutingDetail.GarmentColor) = convert(varchar(3),TransmittesdOrder.MfgColor)  And TransmittesdOrder.MfgSizeCD = Manufacturing.dbo.TTSCutOrderRoutingDetail.SizeDescription GROUP BY Manufacturing.dbo.TTSCutOrderRoutingDetail.WorkOrderNumber, SellStyle,MfgStyle,MfgColor,MfgSizeDesc,SizeDescription"
            .Execute
             '---
                 
                 
             '---#LeadTimeInfo---
            .CommandText = "IF OBJECT_ID('tempdb..#LeadTimeInfo') IS NOT NULL DROP TABLE #LeadTimeInfo"
            .Execute
            .CommandText = "select DISTINCT TI.CutLoc as CutPlant,TI.SewPlantCD as SewPlant  " & _
                            " , CONVERT(varchar, CONVERT(date, CONVERT(VARCHAR(8), CORD.CutDueDate)), 101) AS CutDueDate " & _
                            " ,CONVERT(varchar, CONVERT(date, CONVERT(VARCHAR(8), CORD.SeDueDate)), 101) AS SewDueDate " & _
                            " ,CONVERT(varchar, CONVERT(date, CONVERT(VARCHAR(8), CORD.DCDueDate)), 101) AS DCDueDate " & _
                            " ,TI.OriginalWorkOrder AS WorkOrder, " & _
                            " CORD.OriginalTTSWO AS OriginalWO, " & _
                            " TI.WorkOrder AS WorkLot, " & _
                            " ISNULL(TI.PriorityCD ,ISNULL(CORD.Priority,0)) AS [Priority], " & _
                            " TI.SellStyle AS SellingStyle " & _
                            " ,TI.MfgStyle AS MFGStyle " & _
                            " ,TI.MfgColor AS MFGColor " & _
                            " ,TI.MfgSizeDesc AS MFGSize, " & _
                            " '" & initialTransaction & "' AS InitialTransCode " & _
                            " ,TI.InitialDate, '" & endTransaction & "' AS EndTransCode, " & _
                            " TI.EndDate ,TI.Quantity AS Doz " & _
                            " ,round(CONVERT(decimal(30,2),DATEDIFF (Second, TI.InitialDate, TI.EndDate)) / CONVERT(decimal(30,2), 86400),2) AS LTDays " & _
                            " INTO #LeadTimeInfo " & _
                            " From " & _
                            " #CurrentWO as co INNER JOIN " & _
                            " Manufacturing.dbo.TTSCutOrderRoutingDetail AS CORD WITH (NOLOCK) ON CO.WorkOrderNumber  = CORD.WorkOrderNumber " & _
                            " AND  CO.CutOrderSequence = CORD.CutOrderSequence " & _
                            " RIGHT JOIN #TransactionInfo as TI ON CORD.WorkOrderNumber = TI.WO_NUMBER AND " & _
                            " CORD.GarmentStyle = TI.mfgStyle And Convert(varchar(3), CORD.GarmentColor) = Convert(varchar(3), TI.MfgColor) " & _
                            " And CORD.SizeDescription = TI.MfgSizeCD " & _
                            " ORDER BY WorkLot"
                .Execute
                '---
                
                
                '---QUERY FINAL---
                sQuery = "SELECT ISNULL(CutPlant,'') as CutPlant, SewPlant , ISNULL(CutDueDate,'09/09/1999') as CutDueDate, ISNULL(SewDueDate,'09/09/1999') as SewDueDate, ISNULL(DCDueDate,'09/09/1999') AS DCDueDate, " & _
                            " ISNULL(TRY_CONVERT(BIGINT,WorkOrder),0) as WorkOrder, " & _
                            " convert(BIGINT, isnull(OriginalWO,ISNULL(TRY_CONVERT(BIGINT,OriginalWO),0))) as OriginalWO, " & _
                            " isnull(WorkLot,'') as WorkLot, [Priority], isnull(SellingStyle,'') as SellingStyle, isnull(MFGStyle,'') as MFGStyle, isnull(MFGColor,'') as MFGColor, isnull(MFGSize,'') as MFGSize, isnull(InitialTransCode,'') as InitialTransCode, ISNULL(InitialDate,'09/09/1999 15:46:30') AS  InitialDate " & _
                            " ,isnull(EndTransCode,'') as EndTransCode, ISNULL(EndDate,'09/09/1999 15:46:30') AS EndDate, Doz " & _
                            " ,ISNULL(LTDays,0) AS LTDays " & _
                            " ,CASE WHEN InitialDate is null OR EndDate is null THEN 'Exclude' ELSE CASE WHEN LTDays < 0 THEN 'Exclude' ELSE 'Include' END END AS [ToConsider?] " & _
                            " ,P.[PlantDESC] AS SewPlantName " & _
                            " FROM #LeadTimeInfo as LT " & _
                            " LEFT JOIN Manufacturing.dbo.ANETFacilities as P with (nolock)  on LT.SewPlant = P.PlantCD " & _
                            " ORDER BY SellingStyle"
                '---
                
                
                
                If MADMrs.State = adStateOpen Then
                    MADMrs.Close
                End If
                
                ' Execute query using recordset object.
                MADMrs.CursorType = adOpenForwardOnly
                MADMrs.LockType = adLockReadOnly
                MADMrs.CursorLocation = adUseClient
                MADMConn.CommandTimeout = 0
                MADMrs.Open sQuery, MADMConn, adOpenKeyset, adLockOptimistic
                                                   
                                                   
                'Define Variable
                sTableName = "HistoryTransaction"
                                
                'Define WorkSheet object
                Set oSheetName = Sheets("TransactionInfo")
                                                 
                'Define Table Object
                Set tbl = oSheetName.ListObjects(sTableName)

                If MADMrs.RecordCount > 0 Then
                    Do While Not MADMrs.EOF
                    
                        'GET THE SELLING STYLE FROM TTS.
                        If IsNumeric(MADMrs.Fields("WorkOrder").Value) = False Then
                            sellingFromTTS = Trim(MADMrs.Fields("SellingStyle").Value)
                            cutDueDateFromTTS = MADMrs.Fields("CutDueDate").Value
                            sewDueDateFromTTS = MADMrs.Fields("SewDueDate").Value
                            DCDueDateFromTTS = MADMrs.Fields("DCDueDate").Value
                            OriginalWOFromTTS = MADMrs.Fields("OriginalWO").Value
                        Else
                            getResultsFromTTS = getSellingStyleFromLocalSheet(MADMrs.Fields("WorkOrder").Value, Trim(MADMrs.Fields("SellingStyle").Value), MADMrs.Fields("CutDueDate").Value, MADMrs.Fields("SewDueDate").Value, MADMrs.Fields("DCDueDate").Value, MADMrs.Fields("OriginalWO").Value)
                            sellingFromTTS = getResultsFromTTS(1)
                            cutDueDateFromTTS = getResultsFromTTS(2)
                            sewDueDateFromTTS = getResultsFromTTS(3)
                            DCDueDateFromTTS = getResultsFromTTS(4)
                            OriginalWOFromTTS = getResultsFromTTS(5)
                        End If
                                                
                                                
                        'GET THE CORP BUSINESS UNIT AND WORKCENTER CODE
                        If sellingStyle <> sellingFromTTS Then
                            outputArr = getCorpBusinessUnit_WorkCenter_MegaWorkCenter(MADMrs.Fields("SellingStyle").Value)
                            divisionCode = outputArr(1)
                            workCenter = outputArr(2)
                            megaWorkCenter = outputArr(3)
                            
                            CorpBusinessHubQty = 0
                            
                            For rowCount = 2 To totalRows + 1
                                If Sheet7.Cells(rowCount, 1).Value = SupplyChainHub And Sheet7.Cells(rowCount, 2).Value = divisionCode Then
                                    CorpBusinessHubQty = 1
                                    Exit For
                                End If
                            Next rowCount
                        End If
                        
                                                
                        'INCLUDE OR EXCLUDE FROM THE DATA BY DIVISION
                        If divisionCode = "N/A" Or CorpBusinessHubQty > 0 Then
                            toConsiderMsg = MADMrs.Fields("ToConsider?").Value
                        Else
                            toConsiderMsg = "Exclude"
                        End If
                            
                        Set newrow = tbl.ListRows.Add
                        With newrow
                            .Range(1) = WeekSelected
                            .Range(2) = "'" & plants
                            .Range(3) = plantName
                            .Range(4) = "'" & Trim(MADMrs.Fields("CutPlant").Value)
                            .Range(5) = "'" & Trim(MADMrs.Fields("SewPlant").Value)
                            .Range(6) = cutDueDateFromTTS
                            .Range(7) = sewDueDateFromTTS
                            .Range(8) = DCDueDateFromTTS
                            .Range(9) = MADMrs.Fields("WorkOrder").Value
                            .Range(10) = OriginalWOFromTTS
                            .Range(11) = "'" & Trim(MADMrs.Fields("Priority").Value)
                            .Range(12) = "'" & Trim(MADMrs.Fields("WorkLot").Value)
                            .Range(13) = "'" & sellingFromTTS
                            .Range(14) = "'" & Trim(MADMrs.Fields("MFGStyle").Value)
                            .Range(15) = "'" & Trim(MADMrs.Fields("MFGColor").Value)
                            .Range(16) = "'" & Trim(MADMrs.Fields("MFGSize").Value)
                            .Range(17) = divisionCode
                            .Range(18) = workCenter
                            .Range(19) = megaWorkCenter
                            .Range(20) = bucketId
                            .Range(21) = "'" & MADMrs.Fields("InitialTransCode").Value
                            .Range(22) = MADMrs.Fields("InitialDate").Value
                            .Range(23) = "'" & MADMrs.Fields("EndTransCode").Value
                            .Range(24) = MADMrs.Fields("EndDate").Value
                            .Range(25) = MADMrs.Fields("Doz").Value
                            .Range(26) = MADMrs.Fields("LTDays").Value
                            .Range(27) = toConsiderMsg
                            .Range(28) = CutDueDateAnalysis
                            .Range(29) = SewDueDateAnalysis
                            .Range(30) = DCDueDateAnalysis
                            .Range(31) = SupplyChainHub
                            .Range(32) = BucketDescription
                            .Range(33) = plantCategory
                            .Range(34) = reportType
                            .Range(35) = Process
                        End With
                        sellingStyle = sellingFromTTS
                        MADMrs.MoveNext
                    Loop
                End If
        Next f
    End With
End Sub






' FUNCTION THAT GROUPS THE ORDERS IN RANGES
Public Sub getGroupOfOrders()
    Dim rangeRows As Integer
    Dim r As Integer
    
    'GET THE QUANTITY OF BUCKETS IN THE TABLE.
    BucketRow = getBucketInfoBySupplyHub("TTS", "Textiles")
    If BucketRow = 0 Then
        Exit Sub
    End If
        
    ' CALL THE FUNCTION BY BUCKET SOURCE(ANET/TTS)
    For f = 1 To BucketRow
            
            initialTransaction = Sheet21.Cells(f + 1, 5)
            endTransaction = Sheet21.Cells(f + 1, 6)
            bucketId = Sheet21.Cells(f + 1, 3)
            bucketSource = Sheet21.Cells(f + 1, 8)
            InitialTransCondition = Sheet21.Cells(f + 1, 9)
            EndTransCondition = Sheet21.Cells(f + 1, 10)
            CutDueDateAnalysis = Sheet21.Cells(f + 1, 12)
            SewDueDateAnalysis = Sheet21.Cells(f + 1, 13)
            DCDueDateAnalysis = Sheet21.Cells(f + 1, 14)
            BucketDescription = Sheet21.Cells(f + 1, 4)
            BucketGoalInDays = Sheet21.Cells(f + 1, 11)
                       
            rangeRows = 0
            rowCount = 0
            transmittedOrders = ""
    
            For r = 1 To totalTransmittedOrder
                
                rowCount = rowCount + 1
                rangeRows = rangeRows + 1
                
               If rowCount = totalTransmittedOrder Then
                    transmittedOrders = transmittedOrders & Trim(Sheet6.Cells(r, 1).Value)
                    getDataFromTTS
                    Exit For
                Else
                    If rangeRows = 3000 Then
                        transmittedOrders = transmittedOrders & Trim(Sheet6.Cells(r, 1).Value)
                        getDataFromTTS
                        transmittedOrders = ""
                        rangeRows = 0
                    Else
                        transmittedOrders = transmittedOrders & Trim(Sheet6.Cells(r, 1).Value) & ","
                    End If
                End If
            Next r
        Next f
End Sub








' =====  GET THE TRANSACTION INFO FROM TTS=====
Public Sub getDataFromTTS()

    Dim oSheetName As Worksheet
    Dim sTableName As String
    Dim tbl As ListObject
    Dim newrow As ListRow
     
    Dim divisionCode As String
    Dim workCenter As String
    Dim megaWorkCenter As String
    Dim sellingStyle As String
    Dim CorpBusinessHubQty As Integer
    Dim outputArr As Variant
    Dim getResultsFromLocalSheet As Variant
    Dim toConsiderMsg As String
    
    Dim totalRows As Long
    Dim rowCount As Long
    Dim Process As String
    
    Dim sewPlantFromLocalSheet As String
    
    totalRows = Sheet7.Range("CorpBusinessUnitHub").Rows.Count
    rowCount = 2
    
    divisionCode = ""
    workCenter = ""
    megaWorkCenter = ""
    sellingStyle = ""
    CorpBusinessHubQty = 0
    toConsiderMsg = ""
    Process = "Textiles"
    
    setConnAS400_TTS  ' Set connection to the database.
    
    ' SQL query to fetch details about WorkOrders
    sQuery = "SELECT DISTINCT CutPlant,SewPlant,CutDueDate,SewDueDate,DCDueDate,WorkOrder,OriginalWO,WorkLot,Priority,SellingStyle,MFGStyle,MFGColor,MFGSize,InitialTransCode,VARCHAR_FORMAT(COALESCE(initialDate,COALESCE(initialDateOriginalWO,'2020-03-19-13.33.42')), 'MM/DD/YYYY HH24:MI:SS') AS InitialDate,EndTransCode,VARCHAR_FORMAT(COALESCE(endDate,COALESCE(endDateOriginalWO,'2020-03-19-13.33.42')), 'MM/DD/YYYY HH24:MI:SS') AS EndDate,Doz, ROUND(CAST(TIMESTAMPDIFF(2,CHAR(TIMESTAMP(COALESCE(endDate,COALESCE(endDateOriginalWO,'2020-03-19-13.33.42'))) - TIMESTAMP(COALESCE(initialDate,COALESCE(initialDateOriginalWO,'2020-03-19-13.33.42')))))  AS DEC(60,2))/ CAST(86400  AS DEC(60,2)),2) as LTDays  " & _
             " ,CASE WHEN COALESCE(initialDate,COALESCE(initialDateOriginalWO,''))='' OR  COALESCE(endDate,COALESCE(endDateOriginalWO,''))='' THEN 'Exclude' ELSE CASE WHEN ROUND(CAST(TIMESTAMPDIFF(2,CHAR(TIMESTAMP(COALESCE(endDate,endDateOriginalWO)) - TIMESTAMP(COALESCE(initialDate,initialDateOriginalWO)))) AS DEC(60,2))/ CAST(86400  AS DEC(60,2)),2)<0 THEN 'Exclude' ELSE 'Include' END END AS ToConsider " & _
             " FROM (SELECT DISTINCT CutPlant,SewPlant,Substring(CutDueDate,3,2) || '/' || Right(CutDueDate,2) || '/20' || Left(CutDueDate,2) as CutDueDate, Substring(SewDueDate,3,2) || '/' || Right(SewDueDate,2) || '/20' || Left(SewDueDate,2) as SewDueDate, Substring(DCDueDate,3,2) || '/' || Right(DCDueDate,2) || '/20' || Left(DCDueDate,2) as DCDueDate, WorkOrder, COALESCE(OriginalWO,WorkOrder) AS OriginalWO,WorkLot,SellingStyle,MFGStyle,MFGColor,MFGSize, " & _
             " '" & initialTransaction & "' as InitialTransCode, " & _
             " (SELECT DISTINCT   " & InitialTransCondition & " ('20' || Left(ICLIB.ICP2060.IPTRDT,2) || '-' || Substring(ICLIB.ICP2060.IPTRDT,3,2) || '-' || Right(ICLIB.ICP2060.IPTRDT,2) || '-'||  Left(RIGHT('000000' || Ltrim(Rtrim( ICLIB.ICP2060.IPTRTM)),6),2) || '.' || Substring(RIGHT('000000' || Ltrim(Rtrim( ICLIB.ICP2060.IPTRTM)),6),3,2) || '.' || Right(RIGHT('000000' || Ltrim(Rtrim( ICLIB.ICP2060.IPTRTM)),6),2) ) AS TransDate  FROM ICLIB.ICP2060  WHERE (((ICLIB.ICP2060.IPWLBR)>=0) AND ((ICLIB.ICP2060.IPTRAN)='" & initialTransaction & "') AND (ICLIB.ICP2060.IPTRDT<= " & EndDateTTS & " ) AND  ((ICLIB.ICP2060.IPLTNO =WorkOrder)))) as initialDate, " & _
             " (SELECT DISTINCT   " & InitialTransCondition & " ('20' || Left(ICLIB.ICP2060.IPTRDT,2) || '-' || Substring(ICLIB.ICP2060.IPTRDT,3,2) || '-' || Right(ICLIB.ICP2060.IPTRDT,2) || '-'||  Left(RIGHT('000000' || Ltrim(Rtrim( ICLIB.ICP2060.IPTRTM)),6),2) || '.' || Substring(RIGHT('000000' || Ltrim(Rtrim( ICLIB.ICP2060.IPTRTM)),6),3,2) || '.' || Right(RIGHT('000000' || Ltrim(Rtrim( ICLIB.ICP2060.IPTRTM)),6),2) ) AS TransDate  FROM ICLIB.ICP2060  WHERE (((ICLIB.ICP2060.IPWLBR)>=0) AND ((ICLIB.ICP2060.IPTRAN)='" & initialTransaction & "') AND (ICLIB.ICP2060.IPTRDT<= " & EndDateTTS & " ) AND ((ICLIB.ICP2060.IPLTNO=OriginalWO)))) as initialDateOriginalWO, " & _
             " '" & endTransaction & "' AS EndTransCode, " & _
             " (SELECT DISTINCT  " & EndTransCondition & " ('20' || Left(ICLIB.ICP2060.IPTRDT,2) || '-' || Substring(ICLIB.ICP2060.IPTRDT,3,2) || '-' || Right(ICLIB.ICP2060.IPTRDT,2) || '-'||  Left(RIGHT('000000' || Ltrim(Rtrim( ICLIB.ICP2060.IPTRTM)),6),2) || '.' || Substring(RIGHT('000000' || Ltrim(Rtrim( ICLIB.ICP2060.IPTRTM)),6),3,2) || '.' || Right(RIGHT('000000' || Ltrim(Rtrim( ICLIB.ICP2060.IPTRTM)),6),2) ) AS TransDate FROM ICLIB.ICP2060  WHERE (((ICLIB.ICP2060.IPWLBR)>=0) AND ((ICLIB.ICP2060.IPTRAN)='" & endTransaction & "') AND (ICLIB.ICP2060.IPTRDT<= " & EndDateTTS & " ) AND ((ICLIB.ICP2060.IPLTNO =WorkOrder))))  AS endDate, " & _
             " (SELECT DISTINCT  " & EndTransCondition & " ('20' || Left(ICLIB.ICP2060.IPTRDT,2) || '-' || Substring(ICLIB.ICP2060.IPTRDT,3,2) || '-' || Right(ICLIB.ICP2060.IPTRDT,2) || '-'||  Left(RIGHT('000000' || Ltrim(Rtrim( ICLIB.ICP2060.IPTRTM)),6),2) || '.' || Substring(RIGHT('000000' || Ltrim(Rtrim( ICLIB.ICP2060.IPTRTM)),6),3,2) || '.' || Right(RIGHT('000000' || Ltrim(Rtrim( ICLIB.ICP2060.IPTRTM)),6),2) ) AS TransDate FROM ICLIB.ICP2060  WHERE (((ICLIB.ICP2060.IPWLBR)>=0) AND ((ICLIB.ICP2060.IPTRAN)='" & endTransaction & "') AND (ICLIB.ICP2060.IPTRDT<= " & EndDateTTS & " ) AND ((ICLIB.ICP2060.IPLTNO = OriginalWO))))  AS endDateOriginalWO ,Doz ,Priority " & _
             " from ( " & _
             " SELECT DISTINCT OPLIB.OPV070.C1CUTP AS CutPlant,OPLIB.OPV070.C1SPLT AS SewPlant,OPLIB.OPV070.C1CDUE AS CutDueDate,OPLIB.OPV070.C1SDUE AS SewDueDate, OPLIB.OPV070.C1SDUE " & _
             " AS  DCDueDate,OPLIB.OPV070.C1CUTO AS WorkOrder,(SELECT  DISTINCT ICLIB.ICP2060.IPRMRF  FROM ICLIB.ICP2060  WHERE (((ICLIB.ICP2060.IPLTTP)<>'L')  AND ((ICLIB.ICP2060.IPTRAN)='SP') AND ((ICLIB.ICP2060.IPLTNO=(SELECT CASE WHEN COUNT(ICLIB.ICP2060.IPLTNO)= 0 THEN OPLIB.OPV070.C1CUTO ELSE 0 END AS WO FROM ICLIB.ICP2060  WHERE (((ICLIB.ICP2060.IPLTTP)<>'L')  AND ((ICLIB.ICP2060.IPTRAN)='IR') AND ((ICLIB.ICP2060.IPLTNO=OPLIB.OPV070.C1CUTO))))))) fetch first 1 row only) AS OriginalWO,OPLIB.OPP074.C4ANET_LOT AS WorkLot, OPLIB.OPL770A.C2SSTY AS SellingStyle, OPLIB.OPV070.C1MSTY AS MFGStyle, OPLIB.OPV070.C1GCOL AS MFGColor,'N/A' AS MFGSize, (PLUN01+PLUN02+PLUN03+PLUN04+PLUN05+PLUN06+PLUN07+PLUN08+PLUN09+PLUN10)/12 AS Doz, OPLIB.OPV070.C1OPRI as Priority  FROM (OPLIB.OPV070 INNER JOIN OPLIB.OPP074 ON OPLIB.OPV070.C1CUTO = OPLIB.OPP074.C4CUTO) " & _
             " INNER JOIN OPLIB.OPL770A ON OPLIB.OPV070.C1CUTO = OPLIB.OPL770A.C2CUTO WHERE (((OPLIB.OPV070.C1CUTO) In ( " & transmittedOrders & " )))) as WoInfo) AS WO_Trans_Info ORDER BY SellingStyle"
              
              
  
    If AS400_TTSRs.State = adStateOpen Then
        AS400_TTSRs.Close
    End If
    
    ' Execute query using recordset object.
    AS400_TTSRs.CursorLocation = adUseClient
    AS400_TTSConn.CommandTimeout = 0
    AS400_TTSRs.Open sQuery, AS400_TTSConn, adOpenKeyset, adLockOptimistic
                    
    'Define Variable
    sTableName = "HistoryTransaction"
                    
    'Define WorkSheet object
    Set oSheetName = Sheets("TransactionInfo")
                                     
    'Define Table Object
    Set tbl = oSheetName.ListObjects(sTableName)

    ' Finally, show the details.
    If AS400_TTSRs.RecordCount > 0 Then
        Do While Not AS400_TTSRs.EOF
            
            
            'GET THE SEW PLANT FROM LOCAL SHEET.
            If IsNumeric(AS400_TTSRs.Fields("WorkOrder").Value) = False Then
                sewPlantFromLocalSheet = Trim(AS400_TTSRs.Fields("SewPlant").Value)
            Else
                getResultsFromLocalSheet = getSewPlantFromTransmittedSheet(AS400_TTSRs.Fields("WorkOrder").Value, Trim(AS400_TTSRs.Fields("SewPlant").Value))
                sewPlantFromLocalSheet = getResultsFromLocalSheet(2)
            End If
            
            
            'GET THE CORP BUSINESS UNIT AND WORKCENTER CODE
            If sellingStyle <> AS400_TTSRs.Fields("SellingStyle").Value Then
                outputArr = getCorpBusinessUnit_WorkCenter_MegaWorkCenter(AS400_TTSRs.Fields("SellingStyle").Value)
                divisionCode = outputArr(1)
                workCenter = outputArr(2)
                megaWorkCenter = outputArr(3)
                            
                CorpBusinessHubQty = 0
                For rowCount = 2 To totalRows + 1
                    If Sheet7.Cells(rowCount, 1).Value = SupplyChainHub And Sheet7.Cells(rowCount, 2).Value = divisionCode Then
                       CorpBusinessHubQty = 1
                        Exit For
                    End If
                Next rowCount
            End If
            
            'INCLUDE OR EXCLUDE FROM THE DATA BY DIVISION
            If divisionCode = "N/A" Or CorpBusinessHubQty > 0 Then
                toConsiderMsg = AS400_TTSRs.Fields("ToConsider").Value
            Else
                toConsiderMsg = "Exclude"
            End If
            
            Set newrow = tbl.ListRows.Add
            With newrow
                .Range(1) = WeekSelected
                .Range(2) = "'" & plants
                .Range(3) = plantName
                .Range(4) = "'" & Trim(AS400_TTSRs.Fields("CutPlant").Value)
                '.Range(5) = "'" & Trim(AS400_TTSRs.Fields("SewPlant").Value)
                .Range(5) = "'" & Trim(sewPlantFromLocalSheet)
                .Range(6) = AS400_TTSRs.Fields("CutDueDate").Value
                .Range(7) = AS400_TTSRs.Fields("SewDueDate").Value
                .Range(8) = AS400_TTSRs.Fields("DCDueDate").Value
                .Range(9) = AS400_TTSRs.Fields("WorkOrder").Value
                .Range(10) = AS400_TTSRs.Fields("OriginalWO").Value
                .Range(11) = "'" & Trim(AS400_TTSRs.Fields("Priority").Value)
                .Range(12) = "'" & Trim(AS400_TTSRs.Fields("WorkLot").Value)
                .Range(13) = "'" & Trim(AS400_TTSRs.Fields("SellingStyle").Value)
                .Range(14) = "'" & Trim(AS400_TTSRs.Fields("MFGStyle").Value)
                .Range(15) = "'" & Trim(AS400_TTSRs.Fields("MFGColor").Value)
                .Range(16) = "'" & Trim(AS400_TTSRs.Fields("MFGSize").Value)
                .Range(17) = divisionCode
                .Range(18) = workCenter
                .Range(19) = megaWorkCenter
                .Range(20) = bucketId
                .Range(21) = "'" & AS400_TTSRs.Fields("InitialTransCode").Value
                .Range(22) = AS400_TTSRs.Fields("InitialDate").Value
                .Range(23) = "'" & AS400_TTSRs.Fields("EndTransCode").Value
                .Range(24) = AS400_TTSRs.Fields("EndDate").Value
                .Range(25) = AS400_TTSRs.Fields("Doz").Value
                .Range(26) = AS400_TTSRs.Fields("LTDays").Value
                .Range(27) = toConsiderMsg
                .Range(28) = CutDueDateAnalysis
                .Range(29) = SewDueDateAnalysis
                .Range(30) = DCDueDateAnalysis
                .Range(31) = SupplyChainHub
                .Range(32) = BucketDescription
                .Range(33) = plantCategory
                .Range(34) = reportType
                .Range(35) = Process
            End With
            sellingStyle = AS400_TTSRs.Fields("SellingStyle").Value
            AS400_TTSRs.MoveNext
        Loop
    End If
End Sub






'FUNCTION FOR GROUPING ORDERS AND CALL THE FUNCTION TO GET THE INFO FORM TTS
Public Sub getWorkOrderToFindInfoFromTTS()
    Dim rangeRows As Integer
    Dim r As Integer
    
    rangeRows = 0
    rowCount = 0
    transmittedOrders = ""
    
    'DELETE THE TABLE INFOTMATION
    With Sheet4.ListObjects("WOinfo")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
    
    For r = 1 To totalTransmittedOrder
        rowCount = rowCount + 1
        rangeRows = rangeRows + 1
        
        If rowCount = totalTransmittedOrder Then
            transmittedOrders = transmittedOrders & Trim(Sheet6.Cells(r, 1).Value)
            getWorkOrderInfoFromTTS
            Exit For
        Else
            If rangeRows = 3000 Then
                transmittedOrders = transmittedOrders & Trim(Sheet6.Cells(r, 1).Value)
                getWorkOrderInfoFromTTS
                transmittedOrders = ""
                rangeRows = 0
            Else
                transmittedOrders = transmittedOrders & Trim(Sheet6.Cells(r, 1).Value) & ","
            End If
        End If
    Next r
End Sub





' =====  GET THE WO INFO FROM TTS =====
Public Sub getWorkOrderInfoFromTTS()

    Dim oSheetName As Worksheet
    Dim sTableName As String
    Dim tbl As ListObject
    Dim newrow As ListRow
    
    setConnAS400_TTS  ' Set connection to the database.
     
    ' SQL query to fetch details about WorkOrders
    sQuery = "SELECT DISTINCT OPLIB.OPV070.C1CUTP AS CutPlant,OPLIB.OPV070.C1SPLT AS SewPlant,Substring(OPLIB.OPV070.C1CDUE,3,2) || '/' || Right(OPLIB.OPV070.C1CDUE,2) || '/20' || Left(OPLIB.OPV070.C1CDUE,2) as CutDueDate,Substring(OPLIB.OPV070.C1SDUE,3,2) || '/' || Right(OPLIB.OPV070.C1SDUE,2) || '/20' || Left(OPLIB.OPV070.C1SDUE,2) as SewDueDate, Substring(OPLIB.OPV070.C1ADUE,3,2) || '/' || Right(OPLIB.OPV070.C1ADUE,2) || '/20' || Left(OPLIB.OPV070.C1SDUE,2) AS DCDueDate, " & _
             " OPLIB.OPV070.C1CUTO AS WorkOrder,COALESCE((SELECT  DISTINCT ICLIB.ICP2060.IPRMRF  FROM ICLIB.ICP2060  WHERE (((ICLIB.ICP2060.IPLTTP)<>'L')  AND ((ICLIB.ICP2060.IPTRAN)='SP') AND ((ICLIB.ICP2060.IPLTNO=(SELECT CASE WHEN COUNT(ICLIB.ICP2060.IPLTNO)= 0 THEN OPLIB.OPV070.C1CUTO ELSE 0 END AS WO FROM ICLIB.ICP2060  WHERE (((ICLIB.ICP2060.IPLTTP)<>'L')  AND ((ICLIB.ICP2060.IPTRAN)='IR') AND ((ICLIB.ICP2060.IPLTNO=OPLIB.OPV070.C1CUTO))))))) fetch first 1 row only),OPLIB.OPV070.C1CUTO) AS OriginalWO,OPLIB.OPP074.C4ANET_LOT AS WorkLot, OPLIB.OPL770A.C2SSTY AS SellingStyle, OPLIB.OPV070.C1GSTY AS MFGStyle, OPLIB.OPV070.C1GCOL AS MFGColor,'N/A' AS MFGSize, (PLUN01+PLUN02+PLUN03+PLUN04+PLUN05+PLUN06+PLUN07+PLUN08+PLUN09+PLUN10)/12 AS Doz, OPLIB.OPV070.C1OPRI as Priority  FROM (OPLIB.OPV070 INNER JOIN OPLIB.OPP074 ON OPLIB.OPV070.C1CUTO = OPLIB.OPP074.C4CUTO) " & _
             " INNER JOIN OPLIB.OPL770A ON OPLIB.OPV070.C1CUTO = OPLIB.OPL770A.C2CUTO WHERE (((OPLIB.OPV070.C1CUTO) In (" & transmittedOrders & ")))"
             
    If AS400_TTSRs.State = adStateOpen Then
        AS400_TTSRs.Close
    End If
    
    ' Execute query using recordset object.
    AS400_TTSRs.CursorLocation = adUseClient
    AS400_TTSConn.CommandTimeout = 0
    AS400_TTSRs.Open sQuery, AS400_TTSConn, adOpenKeyset, adLockOptimistic
                    
    'Define Variable
    sTableName = "WOinfo"
                    
    'Define WorkSheet object
    Set oSheetName = Sheets("WorkOrderInfo")
                                     
    'Define Table Object
    Set tbl = oSheetName.ListObjects(sTableName)

     ' Finally, show the details.
    If AS400_TTSRs.RecordCount > 0 Then
        Do While Not AS400_TTSRs.EOF
            Set newrow = tbl.ListRows.Add
            With newrow
                .Range(1) = "'" & Trim(AS400_TTSRs.Fields("CUTPLANT").Value)
                .Range(2) = "'" & Trim(AS400_TTSRs.Fields("SEWPLANT").Value)
                .Range(3) = AS400_TTSRs.Fields("CUTDUEDATE").Value
                .Range(4) = AS400_TTSRs.Fields("SEWDUEDATE").Value
                .Range(5) = AS400_TTSRs.Fields("DCDUEDATE").Value
                .Range(6) = AS400_TTSRs.Fields("WORKORDER").Value
                .Range(7) = AS400_TTSRs.Fields("ORIGINALWO").Value
                .Range(8) = "'" & Trim(AS400_TTSRs.Fields("WORKLOT").Value)
                .Range(9) = "'" & Trim(AS400_TTSRs.Fields("SELLINGSTYLE").Value)
                .Range(10) = "'" & Trim(AS400_TTSRs.Fields("MFGSTYLE").Value)
                .Range(11) = "'" & Trim(AS400_TTSRs.Fields("MFGCOLOR").Value)
                .Range(12) = "'" & Trim(AS400_TTSRs.Fields("MFGSIZE").Value)
                .Range(13) = AS400_TTSRs.Fields("DOZ").Value
                .Range(14) = AS400_TTSRs.Fields("PRIORITY").Value
            End With
            AS400_TTSRs.MoveNext
        Loop
    End If
End Sub






' =====  FUNCTION RETURN THE SELLING STYLE FROM THE ORDER IN TSS. =====
Public Function getSellingStyleFromLocalSheet(wo As Single, actualSelling As String, actualCutDueDate As String, actualSewDueDate As String, actualDCDueDate As String, actualOriginalWO As String) As Variant
    Dim totalRows As Integer
    Dim rowCount As Integer
    Dim newWOInfo As Variant
    totalRows = Sheet4.Range("WOinfo").Rows.Count
    
    ReDim newWOInfo(1 To 5)
    
    newWOInfo(1) = actualSelling
    newWOInfo(2) = actualCutDueDate
    newWOInfo(3) = actualSewDueDate
    newWOInfo(4) = actualDCDueDate
    newWOInfo(5) = actualOriginalWO
    
    For rowCount = 2 To totalRows + 1
        If Sheet4.Cells(rowCount, 6) = wo Then
            newWOInfo(1) = Sheet4.Cells(rowCount, 9)
            newWOInfo(2) = Sheet4.Cells(rowCount, 3)
            newWOInfo(3) = Sheet4.Cells(rowCount, 4)
            newWOInfo(4) = Sheet4.Cells(rowCount, 5)
            newWOInfo(5) = Sheet4.Cells(rowCount, 7)
            Exit For
        End If
    Next rowCount
    
    getSellingStyleFromLocalSheet = newWOInfo
End Function





' =====  FUNCTION RETURN THE SEW PLANT FROM THE  TRANSMITTED ORDER SHEET. =====
Public Function getSewPlantFromTransmittedSheet(wo As Single, actualSewPlant As String) As Variant
    Dim rowCount As Integer
    Dim newWOInfo As Variant

    ReDim newWOInfo(1 To 2)
    
    newWOInfo(1) = wo
    newWOInfo(2) = actualSewPlant

    For rowCount = 2 To totalTransmittedOrder
        If Sheet6.Cells(rowCount, 1) = wo Then
            newWOInfo(1) = Sheet6.Cells(rowCount, 1)
            newWOInfo(2) = Sheet6.Cells(rowCount, 2)
            Exit For
        End If
    Next rowCount
    getSewPlantFromTransmittedSheet = newWOInfo
End Function





Public Sub OpenAccessToImportSharepointList()
'Access object
Dim appAccess As Access.Application

'create new access object
Set appAccess = New Access.Application
'open the acces project
Call appAccess.OpenCurrentDatabase( _
"C:\Lead time Automatic run\DEV\Consolidated LeadTimeCalculationReport_ImportToSharepoint.accdb")
appAccess.Visible = True
End Sub









'FUNCTION TO OBTAIN THE WORKLOTS TRANSACTIONS FROM MADM FOR THE ATTRIBUTION PROCESS
Public Sub GetInformationFromMADN_SKUChangeLots()
    Dim oSheetName As Worksheet
    Dim sTableName As String
    Dim tbl As ListObject
    Dim newrow As ListRow

    'sQuery = "SELECT DISTINCT PlantCD,WorkOrder from Staging.InventoryTransactions with (nolock) WHERE PlantCD ='" & plants & "' AND convert(date,CreatedOnDate) BETWEEN '" & initialDate & "' AND '" & endDate & "'  AND AdjustmentCD ='S' AND AdjustmentDESC ='SKU Change'"
    
    sQuery = "SELECT DISTINCT PlantCD,WorkOrder from Staging.InventoryTransactions with (nolock) WHERE PlantCD ='" & plants & "' AND convert(date,CreatedOnDate) BETWEEN '" & initialDate & "' AND '" & endDate & "'  AND AdjustmentCD ='S' AND AdjustmentDESC ='SKU Change' " & _
            " Union " & _
            " SELECT DISTINCT a.PlantCD, b.WorkOrder " & _
            " from Staging.InventoryTransactions as a with (nolock) inner join Staging.InventoryTransactions as b with (nolock) on a.WorkOrder = b.ParentWorkOrder " & _
            " WHERE a.PlantCD ='" & plants & "' AND convert(date,a.CreatedOnDate) BETWEEN '" & initialDate & "' AND '" & endDate & "'  AND a.AdjustmentCD ='S' AND a.AdjustmentDESC ='SKU Change'"
    
    
    'sQuery = "SELECT DISTINCT PlantCD,WorkOrder from Staging.InventoryTransactions with (nolock) WHERE PlantCD ='67' AND convert(date,CreatedOnDate) BETWEEN '1/1/2023' AND '10/24/2023'  AND AdjustmentCD ='S' AND AdjustmentDESC ='SKU Change' " & _
            " Union " & _
            " SELECT DISTINCT a.PlantCD, b.WorkOrder " & _
            " from Staging.InventoryTransactions as a with (nolock) inner join Staging.InventoryTransactions as b with (nolock) on a.WorkOrder = b.ParentWorkOrder " & _
            " WHERE a.PlantCD ='67' AND convert(date,a.CreatedOnDate) BETWEEN '1/1/2023' AND '10/24/2023'  AND a.AdjustmentCD ='S' AND a.AdjustmentDESC ='SKU Change'"
    
    If MADMrs.State = adStateOpen Then
        MADMrs.Close
    End If
    
    ' Execute query using recordset object.
    MADMrs.CursorType = adOpenForwardOnly
    MADMrs.LockType = adLockReadOnly
    MADMrs.CursorLocation = adUseClient
    MADMConn.CommandTimeout = 0
    MADMrs.Open sQuery, MADMConn, adOpenKeyset, adLockOptimistic
    
    'Define Variable
    sTableName = "SKUChangeWorkLots"
    
    'Define WorkSheet object
    Set oSheetName = Sheets("SKUChangeLots")
    
    'Define Table Object
    Set tbl = oSheetName.ListObjects(sTableName)
    
    If MADMrs.RecordCount > 0 Then
        tbl.ListRows.Add
        tbl.DataBodyRange(tbl.ListRows.Count, 1).CopyFromRecordset MADMrs 'Agrega los datos al final de la tabla
    End If
End Sub
