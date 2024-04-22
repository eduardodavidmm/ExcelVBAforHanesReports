IF OBJECT_ID('tempdb..#TrasmittedManifest') IS NOT NULL DROP TABLE #TrasmittedManifest
SELECT CD.WorkOrder, CD.MfgStyleCD, CD.MfgColorCD, CD.MfgSizeCD, CD.MfgSizeDESC, CD.MfgAttributeCD, 
CD.MfgRevisionNO, CD.PkgStyleCD, CD.PkgColorCD, CD.PkgSizeCD, CD.PkgSizeDESC, CD.PkgAttributeCD, CD.PkgRevisionNO, CD.SelStyleCD, 
CD.SelColorCD, CD.SelSizeCD, CD.SelSizeDESC, CD.SelAttributeCD, CD.SelRevisionNO, CD.ParentWorkOrder, CD.AssortmentParentWO, 
CD.FromPlantCD, CD.CutPlantCD, CD.SewPlantCD, CD.Lot, CD.ParentLot,  lastTrans.lastUpdate, SUM(round(Convert(Decimal(10, 4), 
CD.PiecesApplied) / Convert(Decimal(10, 4), 12), 4)) As Doz  into #TrasmittedManifest  FROM (SELECT Manifest, WorkOrder, MAX(ANETCreatedOnDate) AS 
lastUpdate FROM Manufacturing.dbo.ANETCOODetails WITH (NOLOCK)  WHERE (CONVERT(date, ANETCreatedOnDate) BETWEEN '4/7/2024' AND '4/13/2024') AND (FromPlantCD IN('67')) 
AND (HSCD IS NOT NULL) GROUP BY Manifest, WorkOrder) AS lastTrans INNER JOIN  Manufacturing.dbo.ANETCOODetails AS CD WITH (NOLOCK) ON lastTrans.Manifest = CD.Manifest 
AND lastTrans.WorkOrder = CD.WorkOrder AND lastTrans.lastUpdate = CD.ANETCreatedOnDate GROUP BY CD.WorkOrder, CD.MfgStyleCD, CD.MfgColorCD, CD.MfgSizeCD, CD.MfgSizeDESC, CD.MfgAttributeCD, 
CD.MfgRevisionNO, CD.PkgStyleCD, CD.PkgColorCD, CD.PkgSizeCD, CD.PkgSizeDESC, CD.PkgAttributeCD, CD.PkgRevisionNO, CD.SelStyleCD, CD.SelColorCD, CD.SelSizeCD, CD.SelSizeDESC, CD.SelAttributeCD, 
CD.SelRevisionNO, CD.ParentWorkOrder, CD.AssortmentParentWO, CD.FromPlantCD, CD.CutPlantCD, CD.SewPlantCD,CD.SewPlantCD, CD.Lot, CD.ParentLot,  lastTrans.lastUpdate

IF OBJECT_ID('tempdb..#WOInfo') IS NOT NULL DROP TABLE #WOInfo
select DISTINCT TM.WorkOrder ,OWT.GlobalWorkOrderID  INTO #WOInfo  from #TrasmittedManifest as TM  INNER JOIN Manufacturing.dbo.ANETWorkOrders AS AWO WITH (NOLOCK)  
ON TM.WorkOrder = AWO.WorkOrder AND TM.MfgStyleCD = AWO.StyleCD AND TM.MfgColorCD = AWO.ColorCD AND TM.MfgSizeCD = AWO.SizeCD AND TM.MfgAttributeCD =AWO.AttributeCD  
INNER JOIN dbo.ANETOriginalWorkOrderTrace as OWT WITH (NOLOCK)  ON AWO.GlobalWorkOrderID  = OWT.GlobalWorkOrderID 

IF OBJECT_ID('tempdb..#TransmittedWorkLots') IS NOT NULL DROP TABLE #TransmittedWorkLots
select TM.WorkOrder , TM.ParentWorkOrder, TM.AssortmentParentWO, TM.FromPlantCD, TM.CutPlantCD,  TM.SewPlantCD,SUM(TM.Doz) AS DOZ, OWT.OriginalWorkOrder, OWT.ApparelNETWorkLot, 
OWT.MfgStyle, OWT.MfgColor, OWT.MfgSizeCD, OWT.MfgSizeDesc, OWT.PkgStyle, OWT.PkgColor,  OWT.PkgSizeCD, OWT.PkgSizeDesc, OWT.SellStyle, OWT.SellColor, OWT.SellSizeCD,  OWT.SellSizeDesc, 
OWT.CutLoc, OWT.SewLoc, ISNULL(TRY_CONVERT(int,OWT.OriginalWorkOrder),0) as WO_NUMBER   INTO #TransmittedWorkLots  from #TrasmittedManifest as TM  
INNER JOIN #WOInfo AS WO on TM.WorkOrder =  wo.WorkOrder  LEFT JOIN dbo.ANETOriginalWorkOrderTrace as OWT WITH (NOLOCK)  ON WO.GlobalWorkOrderID  = OWT.GlobalWorkOrderID  
GROUP BY TM.WorkOrder, TM.ParentWorkOrder, TM.AssortmentParentWO, TM.FromPlantCD, TM.CutPlantCD,  TM.SewPlantCD, OWT.OriginalWorkOrder, OWT.ApparelNETWorkLot, OWT.MfgStyle, OWT.MfgColor, 
OWT.MfgSizeCD, OWT.MfgSizeDesc, OWT.PkgStyle, OWT.PkgColor,  OWT.PkgSizeCD, OWT.PkgSizeDesc, OWT.SellStyle, OWT.SellColor, OWT.SellSizeCD,  OWT.SellSizeDesc , OWT.CutLoc, OWT.SewLoc  ORDER BY TM.WorkOrder 

IF OBJECT_ID('tempdb..#TrasmittedManifest') IS NOT NULL DROP TABLE #TrasmittedManifest
SELECT CD.WorkOrder, CD.MfgStyleCD, CD.MfgColorCD, CD.MfgSizeCD, CD.MfgSizeDESC, CD.MfgAttributeCD, CD.MfgRevisionNO, CD.PkgStyleCD, CD.PkgColorCD, CD.PkgSizeCD, CD.PkgSizeDESC, 
CD.PkgAttributeCD, CD.PkgRevisionNO, CD.SelStyleCD, CD.SelColorCD, CD.SelSizeCD, CD.SelSizeDESC, CD.SelAttributeCD, CD.SelRevisionNO, CD.ParentWorkOrder, CD.AssortmentParentWO, CD.FromPlantCD, 
CD.CutPlantCD, CD.SewPlantCD, CD.Lot, CD.ParentLot,  lastTrans.lastUpdate, SUM(round(Convert(Decimal(10, 4), CD.PiecesApplied) / Convert(Decimal(10, 4), 12), 4)) As Doz  into #TrasmittedManifest 
FROM (SELECT Manifest, WorkOrder, MAX(ANETCreatedOnDate) AS lastUpdate FROM Manufacturing.dbo.ANETCOODetails WITH (NOLOCK)  WHERE (CONVERT(date, ANETCreatedOnDate) BETWEEN '4/7/2024' AND '4/13/2024') 
AND (FromPlantCD IN('67')) AND (HSCD IS NOT NULL) GROUP BY Manifest, WorkOrder) AS lastTrans INNER JOIN  Manufacturing.dbo.ANETCOODetails AS CD WITH (NOLOCK) ON lastTrans.Manifest = CD.Manifest 
AND lastTrans.WorkOrder = CD.WorkOrder AND lastTrans.lastUpdate = CD.ANETCreatedOnDate GROUP BY CD.WorkOrder, CD.MfgStyleCD, CD.MfgColorCD, CD.MfgSizeCD, CD.MfgSizeDESC, CD.MfgAttributeCD, 
CD.MfgRevisionNO, CD.PkgStyleCD, CD.PkgColorCD, CD.PkgSizeCD, CD.PkgSizeDESC, CD.PkgAttributeCD, CD.PkgRevisionNO, CD.SelStyleCD, CD.SelColorCD, CD.SelSizeCD, CD.SelSizeDESC, CD.SelAttributeCD, 
CD.SelRevisionNO, CD.ParentWorkOrder, CD.AssortmentParentWO, CD.FromPlantCD, CD.CutPlantCD, CD.SewPlantCD,CD.SewPlantCD, CD.Lot, CD.ParentLot,  lastTrans.lastUpdate

IF OBJECT_ID('tempdb..#WOInfo') IS NOT NULL DROP TABLE #WOInfo
select DISTINCT TM.WorkOrder ,OWT.GlobalWorkOrderID, CONVERT(INT,AWO.PriorityCD)  AS PriorityCD  INTO #WOInfo  from #TrasmittedManifest as TM  INNER JOIN Manufacturing.dbo.ANETWorkOrders AS 
AWO WITH (NOLOCK)  ON TM.WorkOrder = AWO.WorkOrder AND TM.MfgStyleCD = AWO.StyleCD AND TM.MfgColorCD = AWO.ColorCD AND TM.MfgSizeCD = AWO.SizeCD AND TM.MfgAttributeCD =AWO.AttributeCD  
INNER JOIN dbo.ANETOriginalWorkOrderTrace as OWT WITH (NOLOCK)  ON AWO.GlobalWorkOrderID  = OWT.GlobalWorkOrderID

IF OBJECT_ID('tempdb..#TransmittedWorkLots') IS NOT NULL DROP TABLE #TransmittedWorkLots
select TM.WorkOrder , TM.ParentWorkOrder, TM.AssortmentParentWO, TM.FromPlantCD, TM.CutPlantCD,   TM.SewPlantCD,SUM(TM.Doz) AS DOZ, OWT.OriginalWorkOrder, OWT.ApparelNETWorkLot,  TM.MfgStyleCD AS 
MfgStyle, TM.MfgColorCD AS MfgColor, TM.MfgSizeCD AS MfgSizeCD, TM.MfgSizeDesc AS MfgSizeDesc, TM.PkgStyleCD AS PkgStyle, TM.PkgColorCD AS PkgColor,  TM.PkgSizeCD , TM.PkgSizeDesc, TM.SelStyleCD AS 
SellStyle, TM.SelColorCD AS SellColor, TM.SelSizeCD AS SellSizeCD,  TM.SelSizeDESC AS SellSizeDesc, TM.CutPlantCD AS CutLoc, TM.SewPlantCD AS SewLoc,  ISNULL(TRY_CONVERT(int,OWT.OriginalWorkOrder),0) as 
WO_NUMBER  ,MAX(TM.lastUpdate) AS lastUpdate,  ISNULL(TM.AssortmentParentWO, TM.WorkOrder) As WorkLot  ,WO.PriorityCD  INTO #TransmittedWorkLots  from #TrasmittedManifest as TM  INNER JOIN #WOInfo AS WO on 
TM.WorkOrder =  wo.WorkOrder  LEFT JOIN dbo.ANETOriginalWorkOrderTrace as OWT WITH (NOLOCK)  ON WO.GlobalWorkOrderID  = OWT.GlobalWorkOrderID  GROUP BY TM.WorkOrder, TM.ParentWorkOrder, TM.AssortmentParentWO, 
TM.FromPlantCD, TM.CutPlantCD,  TM.SewPlantCD, OWT.OriginalWorkOrder, OWT.ApparelNETWorkLot,  TM.MfgStyleCD, TM.MfgColorCD, TM.MfgSizeCD, TM.MfgSizeDesc, TM.PkgStyleCD, TM.PkgColorCD,  TM.PkgSizeCD , 
TM.PkgSizeDesc, TM.SelStyleCD, TM.SelColorCD, TM.SelSizeCD,  TM.SelSizeDESC , TM.CutPlantCD, TM.SewPlantCD  ,WO.PriorityCD  ORDER BY TM.WorkOrder 

IF OBJECT_ID('tempdb..#TransmittedWorkLotsByDate') IS NOT NULL DROP TABLE #TransmittedWorkLotsByDate
select TW.WorkOrder , TW.ParentWorkOrder, TW.AssortmentParentWO, TW.FromPlantCD, TW.CutPlantCD, TW.SewPlantCD,TW.DOZ, TW.OriginalWorkOrder, TW.ApparelNETWorkLot, TW.MfgStyle, TW.MfgColor, TW.MfgSizeCD, 
TW.MfgSizeDesc, TW.PkgStyle, TW.PkgColor,TW.PkgSizeCD, TW.PkgSizeDesc, TW.SellStyle, TW.SellColor, TW.SellSizeCD,TW.SellSizeDesc, TW.CutLoc, TW.SewLoc, TW.WO_NUMBER,TW.WorkLot, MAX(WOAD.ANETCreatedOnDate) as 
TrasmittedDate,  TW.PriorityCD  INTO #TransmittedWorkLotsByDate  from #TransmittedWorkLots as TW left join dbo.ANETWorkOrderActionDetails as WOAD on  TW.WorkLot = WOAD.WorkOrder  
WHERE CONVERT(DATE,WOAD.ANETCreatedOnDate) <='4/13/2024' AND WOAD.ActionCD = 'SH' AND WOAD.Quantity > 0  GROUP BY TW.WorkOrder , TW.ParentWorkOrder, TW.AssortmentParentWO, TW.FromPlantCD, TW.CutPlantCD,  
TW.SewPlantCD,TW.DOZ, TW.OriginalWorkOrder, TW.ApparelNETWorkLot, TW.MfgStyle, TW.MfgColor, TW.MfgSizeCD, TW.MfgSizeDesc, TW.PkgStyle, TW.PkgColor,  TW.PkgSizeCD, TW.PkgSizeDesc, TW.SellStyle, TW.SellColor, 
TW.SellSizeCD,  TW.SellSizeDesc, TW.CutLoc, TW.SewLoc, TW.WO_NUMBER,TW.WorkLot,TW.PriorityCD

IF OBJECT_ID('tempdb..#InitialTrans') IS NOT NULL DROP TABLE #InitialTrans
SELECT *  into #InitialTrans FROM (  SELECT TWL.WorkOrder,ActionCD,Min(ANETCreatedOnDate) AS DT, 'wo' as obs  FROM #TransmittedWorkLotsByDate  AS TWL INNER JOIN dbo.ANETWorkOrderActionDetails as 
WOD WITH (NOLOCK) ON TWL.WorkOrder = WOD.WorkOrder WHERE ActionCD ='ID' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate GROUP BY TWL.WorkOrder,ActionCD  Union  
SELECT TWL.WorkOrder,ActionCD,Min(ANETCreatedOnDate) AS DT, 'assort' as obs  FROM #TransmittedWorkLotsByDate AS TWL INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON 
TWL.AssortmentParentWO = WOD.WorkOrder WHERE ActionCD ='ID' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate GROUP BY TWL.WorkOrder,ActionCD  Union  SELECT 
TWL.WorkOrder,ActionCD,Min(ANETCreatedOnDate) AS DT, 'Pare' as obs  FROM #TransmittedWorkLotsByDate AS TWL INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON 
TWL.ParentWorkOrder = WOD.WorkOrder WHERE ActionCD ='ID' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate GROUP BY TWL.WorkOrder,ActionCD)  AS InitialTrans 

IF OBJECT_ID('tempdb..#EndTrans') IS NOT NULL DROP TABLE #EndTrans
SELECT *  into #EndTrans FROM (  SELECT TWL.WorkOrder,ActionCD,Max(ANETCreatedOnDate) AS DT, 'wo' as obs  FROM #TransmittedWorkLotsByDate AS TWL 
INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.WorkOrder = WOD.WorkOrder WHERE ActionCD ='FQ' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate 
GROUP BY TWL.WorkOrder,ActionCD  Union  SELECT TWL.WorkOrder,ActionCD,Max(ANETCreatedOnDate) AS DT, 'assort' as obs  FROM #TransmittedWorkLotsByDate AS TWL 
INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.AssortmentParentWO = WOD.WorkOrder WHERE ActionCD ='FQ' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate 
GROUP BY TWL.WorkOrder,ActionCD  Union  SELECT TWL.WorkOrder,ActionCD,Max(ANETCreatedOnDate) AS DT, 'Pare' as obs  FROM #TransmittedWorkLotsByDate AS TWL 
INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.ParentWorkOrder = WOD.WorkOrder WHERE ActionCD ='FQ' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate 
GROUP BY TWL.WorkOrder,ActionCD)  AS EndTrans

IF OBJECT_ID('tempdb..#TransactionInfo') IS NOT NULL DROP TABLE #TransactionInfo
select DISTINCT *,  'ID' AS InitialTrans,  ISNULL((SELECT dt FROM #InitialTrans WHERE WorkOrder = 
TWL.WorkOrder AND obs ='WO'),  ISNULL((SELECT dt FROM #InitialTrans WHERE WorkOrder = TWL.WorkOrder AND obs ='assort'),(SELECT dt 
FROM #InitialTrans WHERE WorkOrder = TWL.WorkOrder AND obs ='Pare'))) AS InitialDate,  'FQ' AS EndTrans,  ISNULL((SELECT dt FROM #EndTrans  WHERE WorkOrder = TWL.WorkOrder AND obs ='WO'),  
ISNULL((SELECT dt FROM #EndTrans WHERE WorkOrder = TWL.WorkOrder AND obs ='assort'),(SELECT dt FROM #EndTrans WHERE WorkOrder = TWL.WorkOrder AND obs ='Pare'))) AS EndDate,  TWL.DOZ  AS 
Quantity  into #TransactionInfo  from #TransmittedWorkLots as TWL

IF OBJECT_ID('tempdb..#CurrentWO') IS NOT NULL DROP TABLE #CurrentWO
SELECT  dbo.TTSCutOrderRoutingDetail.WorkOrderNumber, MAX(dbo.TTSCutOrderRoutingDetail.CutOrderSequence) AS CutOrderSequence, 
RIGHT('000000' + CONVERT(varchar, dbo.TTSCutOrderRoutingDetail.WorkOrderNumber), 6)AS CurrentWorkOrderNumber, SellStyle , MfgStyle, MfgColor, MfgSizeDesc, 
SizeDescription  INTO #CurrentWO  FROM  dbo.TTSCutOrderRoutingDetail WITH (NOLOCK) INNER JOIN #TransactionInfo AS TransmittesdOrder 
ON dbo.TTSCutOrderRoutingDetail.WorkOrderNumber = TransmittesdOrder.WO_NUMBER AND dbo.TTSCutOrderRoutingDetail.GarmentStyle = TransmittesdOrder.MfgStyle 
AND dbo.TTSCutOrderRoutingDetail.GarmentColor = TransmittesdOrder.MfgColor  And TransmittesdOrder.MfgSizeCD = dbo.TTSCutOrderRoutingDetail.SizeDescription GROUP BY 
dbo.TTSCutOrderRoutingDetail.WorkOrderNumber, SellStyle,MfgStyle,MfgColor,MfgSizeDesc,SizeDescription

IF OBJECT_ID('tempdb..#LeadTimeInfo') IS NOT NULL DROP TABLE #LeadTimeInfo
select DISTINCT TI.CutLoc as CutPlant,TI.SewPlantCD as SewPlant   , CONVERT(varchar, CONVERT(date, CONVERT(VARCHAR(8), CORD.CutDueDate)), 101) 
AS CutDueDate  ,CONVERT(varchar, CONVERT(date, CONVERT(VARCHAR(8), CORD.SeDueDate)), 101) AS SewDueDate  ,CONVERT(varchar, CONVERT(date, CONVERT(VARCHAR(8), CORD.DCDueDate)), 101) 
AS DCDueDate  ,TI.OriginalWorkOrder AS WorkOrder,  CORD.OriginalTTSWO AS OriginalWO,  TI.WorkOrder AS WorkLot,  ISNULL(TI.PriorityCD ,ISNULL(CORD.Priority,0)) AS [Priority],  
TI.SellStyle AS SellingStyle  ,TI.MfgStyle AS MFGStyle  ,TI.MfgColor AS MFGColor  ,TI.MfgSizeDesc AS MFGSize,  'ID' AS InitialTransCode  ,TI.InitialDate, 'FQ' 
AS EndTransCode, TI.EndDate  ,TI.Quantity AS Doz  ,round(CONVERT(decimal(30,2),DATEDIFF (Second, TI.InitialDate, TI.EndDate)) / CONVERT(decimal(30,2), 86400),2) AS LTDays  
INTO #LeadTimeInfo  From  #CurrentWO as co INNER JOIN  Manufacturing.dbo.TTSCutOrderRoutingDetail AS CORD WITH (NOLOCK) ON CO.WorkOrderNumber  = CORD.WorkOrderNumber  
AND  CO.CutOrderSequence = CORD.CutOrderSequence  RIGHT JOIN #TransactionInfo as TI ON CORD.WorkOrderNumber = TI.WO_NUMBER AND  CORD.GarmentStyle = TI.mfgStyle And CORD.GarmentColor = TI.MfgColor 
And CORD.SizeDescription = TI.MfgSizeCD  ORDER BY WorkLot

