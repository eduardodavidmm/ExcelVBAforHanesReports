--Tasnitted Manifest 
IF OBJECT_ID('tempdb..#TrasmittedManifest') IS NOT NULL DROP TABLE #TrasmittedManifest
SELECT CD.WorkOrder, CD.MfgStyleCD, CD.MfgColorCD, CD.MfgSizeCD, CD.MfgSizeDESC, CD.MfgAttributeCD, CD.MfgRevisionNO, 
CD.PkgStyleCD, CD.PkgColorCD, CD.PkgSizeCD, CD.PkgSizeDESC, CD.PkgAttributeCD, CD.PkgRevisionNO, CD.SelStyleCD, CD.SelColorCD, 
CD.SelSizeCD, CD.SelSizeDESC, CD.SelAttributeCD, CD.SelRevisionNO, CD.ParentWorkOrder, CD.AssortmentParentWO, CD.FromPlantCD, 
CD.CutPlantCD, CD.SewPlantCD, CD.Lot, CD.ParentLot,  lastTrans.lastUpdate, 
UM(round(Convert(Decimal(10, 4), CD.PiecesApplied) / Convert(Decimal(10, 4), 12), 4)) As Doz  into #TrasmittedManifest  
FROM (SELECT Manifest, WorkOrder, MAX(ANETCreatedOnDate) AS lastUpdate FROM Manufacturing.dbo.ANETCOODetails WITH (NOLOCK)  
WHERE (CONVERT(date, ANETCreatedOnDate) BETWEEN '4/14/2024' AND '4/20/2024') AND (FromPlantCD IN('86')) AND (HSCD IS NOT NULL) 
GROUP BY Manifest, WorkOrder) AS lastTrans INNER JOIN  Manufacturing.dbo.ANETCOODetails AS CD WITH (NOLOCK) 
ON lastTrans.Manifest = CD.Manifest AND lastTrans.WorkOrder = CD.WorkOrder AND lastTrans.lastUpdate = CD.ANETCreatedOnDate 
GROUP BY CD.WorkOrder, CD.MfgStyleCD, CD.MfgColorCD, CD.MfgSizeCD, CD.MfgSizeDESC, CD.MfgAttributeCD, CD.MfgRevisionNO, 
CD.PkgStyleCD, CD.PkgColorCD, CD.PkgSizeCD, CD.PkgSizeDESC, CD.PkgAttributeCD, CD.PkgRevisionNO, CD.SelStyleCD, CD.SelColorCD, 
CD.SelSizeCD, CD.SelSizeDESC, CD.SelAttributeCD, CD.SelRevisionNO, CD.ParentWorkOrder, CD.AssortmentParentWO, CD.FromPlantCD, 
CD.CutPlantCD, CD.SewPlantCD,CD.SewPlantCD, CD.Lot, CD.ParentLot,  lastTrans.lastUpdate

-- Work Order Info
IF OBJECT_ID('tempdb..#WOInfo') IS NOT NULL DROP TABLE #WOInfo
select DISTINCT TM.WorkOrder ,OWT.GlobalWorkOrderID  INTO #WOInfo  from #TrasmittedManifest as TM  
INNER JOIN Manufacturing.dbo.ANETWorkOrders AS AWO WITH (NOLOCK)  ON TM.WorkOrder = AWO.WorkOrder AND TM.MfgStyleCD = AWO.StyleCD 
AND TM.MfgColorCD = AWO.ColorCD AND TM.MfgSizeCD = AWO.SizeCD AND TM.MfgAttributeCD =AWO.AttributeCD    INNER JOIN dbo.ANETOriginalWorkOrderTrace as 
OWT WITH (NOLOCK)  ON AWO.GlobalWorkOrderID  = OWT.GlobalWorkOrderID 

-- Get Transmitted Workots
IF OBJECT_ID('tempdb..#TransmittedWorkLots') IS NOT NULL DROP TABLE #TransmittedWorkLots
select TM.WorkOrder , TM.ParentWorkOrder, TM.AssortmentParentWO, TM.FromPlantCD, TM.CutPlantCD,  TM.SewPlantCD,SUM(TM.Doz) 
AS DOZ, OWT.OriginalWorkOrder, OWT.ApparelNETWorkLot, OWT.MfgStyle, OWT.MfgColor, OWT.MfgSizeCD, OWT.MfgSizeDesc, OWT.PkgStyle, 
OWT.PkgColor,  OWT.PkgSizeCD, OWT.PkgSizeDesc, OWT.SellStyle, OWT.SellColor, OWT.SellSizeCD,  OWT.SellSizeDesc, OWT.CutLoc, OWT.SewLoc, 
ISNULL(TRY_CONVERT(int,OWT.OriginalWorkOrder),0) as WO_NUMBER   INTO #TransmittedWorkLots  from #TrasmittedManifest as TM  
INNER JOIN #WOInfo AS WO on TM.WorkOrder =  wo.WorkOrder  LEFT JOIN dbo.ANETOriginalWorkOrderTrace as OWT WITH (NOLOCK)  
ON WO.GlobalWorkOrderID  = OWT.GlobalWorkOrderID  GROUP BY TM.WorkOrder, TM.ParentWorkOrder, TM.AssortmentParentWO, TM.FromPlantCD, 
TM.CutPlantCD,  TM.SewPlantCD, OWT.OriginalWorkOrder, OWT.ApparelNETWorkLot, OWT.MfgStyle, OWT.MfgColor, OWT.MfgSizeCD, OWT.MfgSizeDesc, 
OWT.PkgStyle, OWT.PkgColor,  OWT.PkgSizeCD, OWT.PkgSizeDesc, OWT.SellStyle, OWT.SellColor, OWT.SellSizeCD,  OWT.SellSizeDesc , OWT.CutLoc, OWT.SewLoc  
ORDER BY TM.WorkOrder 

-- Get Transmitted Manifest 
IF OBJECT_ID('tempdb..#TrasmittedManifest') IS NOT NULL DROP TABLE #TrasmittedManifest
SELECT CD.WorkOrder, CD.MfgStyleCD, CD.MfgColorCD, CD.MfgSizeCD, CD.MfgSizeDESC, CD.MfgAttributeCD, CD.MfgRevisionNO, 
CD.PkgStyleCD, CD.PkgColorCD, CD.PkgSizeCD, CD.PkgSizeDESC, CD.PkgAttributeCD, CD.PkgRevisionNO, CD.SelStyleCD, CD.SelColorCD, 
CD.SelSizeCD, CD.SelSizeDESC, CD.SelAttributeCD, CD.SelRevisionNO, CD.ParentWorkOrder, CD.AssortmentParentWO, CD.FromPlantCD, CD.CutPlantCD, 
CD.SewPlantCD, CD.Lot, CD.ParentLot,  lastTrans.lastUpdate, SUM(round(Convert(Decimal(10, 4), CD.PiecesApplied) / Convert(Decimal(10, 4), 12), 4)) 
As Doz  into #TrasmittedManifest  FROM (SELECT Manifest, WorkOrder, MAX(ANETCreatedOnDate) AS lastUpdate FROM Manufacturing.dbo.ANETCOODetails WITH (NOLOCK)  
WHERE (CONVERT(date, ANETCreatedOnDate) BETWEEN '4/14/2024' AND '4/20/2024') AND (FromPlantCD IN('86')) AND (HSCD IS NOT NULL) GROUP BY Manifest, WorkOrder) 
AS lastTrans INNER JOIN  Manufacturing.dbo.ANETCOODetails AS CD WITH (NOLOCK) ON lastTrans.Manifest = CD.Manifest AND lastTrans.WorkOrder = CD.WorkOrder 
AND lastTrans.lastUpdate = CD.ANETCreatedOnDate GROUP BY CD.WorkOrder, CD.MfgStyleCD, CD.MfgColorCD, CD.MfgSizeCD, CD.MfgSizeDESC, CD.MfgAttributeCD, 
CD.MfgRevisionNO, CD.PkgStyleCD, CD.PkgColorCD, CD.PkgSizeCD, CD.PkgSizeDESC, CD.PkgAttributeCD, CD.PkgRevisionNO, CD.SelStyleCD, CD.SelColorCD, CD.SelSizeCD, 
CD.SelSizeDESC, CD.SelAttributeCD, CD.SelRevisionNO, CD.ParentWorkOrder, CD.AssortmentParentWO, CD.FromPlantCD, CD.CutPlantCD, CD.SewPlantCD,CD.SewPlantCD, 
CD.Lot, CD.ParentLot,  lastTrans.lastUpdate

-- Get Work Order Info
IF OBJECT_ID('tempdb..#WOInfo') IS NOT NULL DROP TABLE #WOInfo
select DISTINCT TM.WorkOrder ,OWT.GlobalWorkOrderID, CONVERT(INT,AWO.PriorityCD)  AS PriorityCD  INTO #WOInfo  
from #TrasmittedManifest as TM  INNER JOIN Manufacturing.dbo.ANETWorkOrders AS AWO WITH (NOLOCK)  ON TM.WorkOrder = AWO.WorkOrder 
AND TM.MfgStyleCD = AWO.StyleCD AND TM.MfgColorCD = AWO.ColorCD AND TM.MfgSizeCD = AWO.SizeCD AND TM.MfgAttributeCD =AWO.AttributeCD  
INNER JOIN dbo.ANETOriginalWorkOrderTrace as OWT WITH (NOLOCK)  ON AWO.GlobalWorkOrderID  = OWT.GlobalWorkOrderID

-- Get Transmitted Worklots
IF OBJECT_ID('tempdb..#TransmittedWorkLots') IS NOT NULL DROP TABLE #TransmittedWorkLots
select TM.WorkOrder , TM.ParentWorkOrder, TM.AssortmentParentWO, TM.FromPlantCD, TM.CutPlantCD, TM.SewPlantCD,SUM(TM.Doz) 
AS DOZ, OWT.OriginalWorkOrder, OWT.ApparelNETWorkLot,  TM.MfgStyleCD AS MfgStyle, TM.MfgColorCD AS MfgColor, TM.MfgSizeCD AS MfgSizeCD, 
TM.MfgSizeDesc AS MfgSizeDesc, TM.PkgStyleCD AS PkgStyle, TM.PkgColorCD AS PkgColor,  TM.PkgSizeCD , TM.PkgSizeDesc, TM.SelStyleCD 
AS SellStyle, TM.SelColorCD AS SellColor, TM.SelSizeCD AS SellSizeCD,  TM.SelSizeDESC AS SellSizeDesc, TM.CutPlantCD AS CutLoc, TM.SewPlantCD 
AS SewLoc,  ISNULL(TRY_CONVERT(int,OWT.OriginalWorkOrder),0) as WO_NUMBER  ,MAX(TM.lastUpdate) AS lastUpdate,  ISNULL(TM.AssortmentParentWO, TM.WorkOrder) 
As WorkLot  ,WO.PriorityCD  INTO #TransmittedWorkLots  from #TrasmittedManifest as TM  INNER JOIN #WOInfo AS WO on TM.WorkOrder =  wo.WorkOrder  
LEFT JOIN dbo.ANETOriginalWorkOrderTrace as OWT WITH (NOLOCK)  ON WO.GlobalWorkOrderID  = OWT.GlobalWorkOrderID  GROUP BY TM.WorkOrder, TM.ParentWorkOrder, 
TM.AssortmentParentWO, TM.FromPlantCD, TM.CutPlantCD,  TM.SewPlantCD, OWT.OriginalWorkOrder, OWT.ApparelNETWorkLot,  TM.MfgStyleCD, TM.MfgColorCD, 
TM.MfgSizeCD, TM.MfgSizeDesc, TM.PkgStyleCD, TM.PkgColorCD,  TM.PkgSizeCD , TM.PkgSizeDesc, TM.SelStyleCD, TM.SelColorCD, TM.SelSizeCD,  TM.SelSizeDESC , 
TM.CutPlantCD, TM.SewPlantCD  ,WO.PriorityCD  ORDER BY TM.WorkOrder 

--SQLQuery
SELECT DISTINCT WO_NUMBER AS WO,SewPlantCD as Plant FROM #TransmittedWorkLots

--Get Transmitted Worklot by Dates
IF OBJECT_ID('tempdb..#TransmittedWorkLotsByDate') IS NOT NULL DROP TABLE #TransmittedWorkLotsByDate
select TW.WorkOrder , TW.ParentWorkOrder, TW.AssortmentParentWO, TW.FromPlantCD, TW.CutPlantCD, TW.SewPlantCD,TW.DOZ, 
TW.OriginalWorkOrder, TW.ApparelNETWorkLot, TW.MfgStyle, TW.MfgColor, TW.MfgSizeCD, TW.MfgSizeDesc, TW.PkgStyle, TW.PkgColor,TW.PkgSizeCD, 
TW.PkgSizeDesc, TW.SellStyle, TW.SellColor, TW.SellSizeCD,TW.SellSizeDesc, TW.CutLoc, TW.SewLoc, TW.WO_NUMBER,TW.WorkLot, MAX(WOAD.ANETCreatedOnDate) 
as TrasmittedDate,  TW.PriorityCD  INTO #TransmittedWorkLotsByDate  from #TransmittedWorkLots as TW left join dbo.ANETWorkOrderActionDetails as WOAD on  
TW.WorkLot = WOAD.WorkOrder  WHERE CONVERT(DATE,WOAD.ANETCreatedOnDate) <='4/20/2024' AND WOAD.ActionCD = 'SH' AND WOAD.Quantity > 0  GROUP BY TW.WorkOrder , 
TW.ParentWorkOrder, TW.AssortmentParentWO, TW.FromPlantCD, TW.CutPlantCD,  TW.SewPlantCD,TW.DOZ, TW.OriginalWorkOrder, TW.ApparelNETWorkLot, TW.MfgStyle, 
TW.MfgColor, TW.MfgSizeCD, TW.MfgSizeDesc, TW.PkgStyle, TW.PkgColor,  TW.PkgSizeCD, TW.PkgSizeDesc, TW.SellStyle, TW.SellColor, TW.SellSizeCD,  TW.SellSizeDesc, 
TW.CutLoc, TW.SewLoc, TW.WO_NUMBER,TW.WorkLot,TW.PriorityCD

--Get Initial Trans
IF OBJECT_ID('tempdb..#InitialTrans') IS NOT NULL DROP TABLE #InitialTrans
SELECT *  into #InitialTrans FROM (  SELECT TWL.WorkOrder,ActionCD,Max(ANETCreatedOnDate) AS DT, 'wo' as obs  FROM #TransmittedWorkLotsByDate  
AS TWL INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.WorkOrder = WOD.WorkOrder WHERE ActionCD ='IT' AND Quantity>0 AND 
ANETCreatedOnDate<= TWL.TrasmittedDate GROUP BY TWL.WorkOrder,ActionCD  Union  SELECT TWL.WorkOrder,ActionCD,Max(ANETCreatedOnDate) AS DT, 'assort' as obs  
FROM #TransmittedWorkLotsByDate AS TWL INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.AssortmentParentWO = WOD.WorkOrder 
WHERE ActionCD ='IT' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate GROUP BY TWL.WorkOrder,ActionCD  Union  
SELECT TWL.WorkOrder,ActionCD,Max(ANETCreatedOnDate) AS DT, 'Pare' as obs  FROM #TransmittedWorkLotsByDate AS TWL 
INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.ParentWorkOrder = WOD.WorkOrder WHERE ActionCD ='IT' AND Quantity>0 
AND ANETCreatedOnDate<= TWL.TrasmittedDate GROUP BY TWL.WorkOrder,ActionCD)  AS InitialTrans 

--Get End Transaction Date
IF OBJECT_ID('tempdb..#EndTrans') IS NOT NULL DROP TABLE #EndTrans
SELECT WorkOrder, 'SH' AS ActionCD, Max(TrasmittedDate) as DT, 'wo' AS obs into #EndTrans FROM  #TransmittedWorkLotsByDate group by WorkOrder
SELECT *  into #EndTrans FROM (  SELECT TWL.WorkOrder,ActionCD,Max(ANETCreatedOnDate) AS DT, 'wo' as obs  FROM #TransmittedWorkLotsByDate 
AS TWL INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.WorkOrder = WOD.WorkOrder WHERE ActionCD ='OS' AND Quantity>0 
AND ANETCreatedOnDate<= TWL.TrasmittedDate GROUP BY TWL.WorkOrder,ActionCD  Union  SELECT TWL.WorkOrder,ActionCD,Max(ANETCreatedOnDate) 
AS DT, 'assort' as obs  FROM #TransmittedWorkLotsByDate AS TWL INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) 
ON TWL.AssortmentParentWO = WOD.WorkOrder WHERE ActionCD ='OS' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate GROUP BY TWL.WorkOrder,ActionCD  
Union  SELECT TWL.WorkOrder,ActionCD,Max(ANETCreatedOnDate) AS DT, 'Pare' as obs  FROM #TransmittedWorkLotsByDate AS TWL INNER JOIN dbo.ANETWorkOrderActionDetails 
as WOD WITH (NOLOCK) ON TWL.ParentWorkOrder = WOD.WorkOrder WHERE ActionCD ='OS' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate 
GROUP BY TWL.WorkOrder,ActionCD)  AS EndTrans

--Get Transaction Info
IF OBJECT_ID('tempdb..#TransactionInfo') IS NOT NULL DROP TABLE #TransactionInfo
select DISTINCT *,  'ID' AS InitialTrans,  ISNULL((SELECT dt FROM #InitialTrans WHERE WorkOrder = TWL.WorkOrder AND obs ='WO'), 
ISNULL((SELECT dt FROM #InitialTrans WHERE WorkOrder = TWL.WorkOrder AND obs ='assort'),(SELECT dt FROM #InitialTrans WHERE WorkOrder = TWL.WorkOrder 
AND obs ='Pare'))) AS InitialDate,  'FQ' AS EndTrans,  ISNULL((SELECT dt FROM #EndTrans  WHERE WorkOrder = TWL.WorkOrder AND obs ='WO'),  
ISNULL((SELECT dt FROM #EndTrans WHERE WorkOrder = TWL.WorkOrder AND obs ='assort'),(SELECT dt FROM #EndTrans WHERE WorkOrder = TWL.WorkOrder AND obs ='Pare'))) 
AS EndDate,  TWL.DOZ  AS Quantity  into #TransactionInfo  from #TransmittedWorkLots as TWL

--Get Current WorkOrder
IF OBJECT_ID('tempdb..#CurrentWO') IS NOT NULL DROP TABLE #CurrentWO
SELECT  dbo.TTSCutOrderRoutingDetail.WorkOrderNumber, MAX(dbo.TTSCutOrderRoutingDetail.CutOrderSequence) AS CutOrderSequence, 
RIGHT('000000' + CONVERT(varchar, dbo.TTSCutOrderRoutingDetail.WorkOrderNumber), 6)AS CurrentWorkOrderNumber, SellStyle , MfgStyle, MfgColor, MfgSizeDesc, 
SizeDescription  INTO #CurrentWO  FROM  dbo.TTSCutOrderRoutingDetail WITH (NOLOCK) INNER JOIN #TransactionInfo AS TransmittesdOrder 
ON dbo.TTSCutOrderRoutingDetail.WorkOrderNumber = TransmittesdOrder.WO_NUMBER AND dbo.TTSCutOrderRoutingDetail.GarmentStyle = TransmittesdOrder.MfgStyle AND 
dbo.TTSCutOrderRoutingDetail.GarmentColor = TransmittesdOrder.MfgColor  And TransmittesdOrder.MfgSizeCD = dbo.TTSCutOrderRoutingDetail.SizeDescription 
GROUP BY dbo.TTSCutOrderRoutingDetail.WorkOrderNumber, SellStyle,MfgStyle,MfgColor,MfgSizeDesc,SizeDescription

--Get LeadTime Info
IF OBJECT_ID('tempdb..#LeadTimeInfo') IS NOT NULL DROP TABLE #LeadTimeInfo
select DISTINCT TI.CutLoc as CutPlant,TI.SewPlantCD as SewPlant   , CONVERT(varchar, CONVERT(date, CONVERT(VARCHAR(8), CORD.CutDueDate)), 101) 
AS CutDueDate  ,CONVERT(varchar, CONVERT(date, CONVERT(VARCHAR(8), CORD.SeDueDate)), 101) AS SewDueDate  ,CONVERT(varchar, CONVERT(date, 
CONVERT(VARCHAR(8), CORD.DCDueDate)), 101) AS DCDueDate  ,TI.OriginalWorkOrder AS WorkOrder,  CORD.OriginalTTSWO AS OriginalWO,  TI.WorkOrder AS WorkLot,  
ISNULL(TI.PriorityCD ,ISNULL(CORD.Priority,0)) AS [Priority],  TI.SellStyle AS SellingStyle  ,TI.MfgStyle AS MFGStyle  ,TI.MfgColor AS MFGColor  ,TI.MfgSizeDesc 
AS MFGSize,  'IT' AS InitialTransCode  ,TI.InitialDate, 'OS' AS EndTransCode, TI.EndDate  ,TI.Quantity AS Doz  ,round(CONVERT(decimal(30,2),DATEDIFF 
(Second, TI.InitialDate, TI.EndDate)) / CONVERT(decimal(30,2), 86400),2) AS LTDays  INTO #LeadTimeInfo  From  #CurrentWO as co INNER JOIN  
Manufacturing.dbo.TTSCutOrderRoutingDetail AS CORD WITH (NOLOCK) ON CO.WorkOrderNumber  = CORD.WorkOrderNumber  AND  CO.CutOrderSequence = CORD.CutOrderSequence  
RIGHT JOIN #TransactionInfo as TI ON CORD.WorkOrderNumber = TI.WO_NUMBER AND  CORD.GarmentStyle = TI.mfgStyle And CORD.GarmentColor = TI.MfgColor  And CORD.SizeDescription = TI.MfgSizeCD  
ORDER BY WorkLot

--SQUERY LTINFO
SELECT ISNULL(CutPlant,'') as CutPlant, SewPlant , ISNULL(CutDueDate,'09/09/1999') as CutDueDate, ISNULL(SewDueDate,'09/09/1999') as SewDueDate, ISNULL(DCDueDate,'09/09/1999') AS DCDueDate, 
ISNULL(TRY_CONVERT(BIGINT,WorkOrder),0) as WorkOrder, 
convert(BIGINT, isnull(OriginalWO,ISNULL(TRY_CONVERT(BIGINT,OriginalWO),0))) as OriginalWO, 
isnull(WorkLot,'') as WorkLot, [Priority], isnull(SellingStyle,'') as SellingStyle, isnull(MFGStyle,'') as MFGStyle, isnull(MFGColor,'') as MFGColor, isnull(MFGSize,'') 
as MFGSize, isnull(InitialTransCode,'') as InitialTransCode, ISNULL(InitialDate,'09/09/1999 15:46:30') AS  InitialDate 
,isnull(EndTransCode,'') as EndTransCode, ISNULL(EndDate,'09/09/1999 15:46:30') AS EndDate, Doz 
,ISNULL(LTDays,0) AS LTDays 
,CASE WHEN InitialDate is null OR EndDate is null THEN 'Exclude' ELSE CASE WHEN LTDays < 0 THEN 'Exclude' ELSE 'Include' END END AS [ToConsider?] 
,P.[PlantDESC] AS SewPlantName 
FROM #LeadTimeInfo as LT 
LEFT JOIN Manufacturing.dbo.ANETFacilities as P with (nolock)  on LT.SewPlant = P.PlantCD 
ORDER BY SellingStyle

-- Get AttributionLots with Work Order
IF OBJECT_ID('tempdb..#AttributionLotsWithWO') IS NOT NULL DROP TABLE #AttributionLotsWithWO
SELECT distinct ACD.WorkOrder,ACD.MfgStyleCD,ACD.MfgColorCD,ACD.MfgSizeCD,ACD.MfgAttributeCD  ,MfgRevisionCD  ,ACD.SewPlantCD ,OWT.OriginalWorkOrder, 
convert(varchar(3), ACD.MfgColorCD) as MfgColorCDWithThreeCharacters  into #AttributionLotsWithWO  FROM   Manufacturing.dbo.ANETCOODetails AS ACD  
WITH (NOLOCK)  INNER JOIN Manufacturing.dbo.ANETWorkOrders AS AWO  WITH (NOLOCK) ON  AWO.WorkOrder = ACD.WorkOrder  INNER JOIN 
Manufacturing.dbo.ANETOriginalWorkOrderTrace as OWT WITH (NOLOCK) ON OWT.GlobalWorkOrderID =AWO.GlobalWorkOrderID AND OWT.MfgStyle  = 
ACD.MfgStyleCD AND convert(varchar(3),OWT.MfgColor)  = convert(varchar(3),ACD.MfgColorCD)  AND OWT.MfgSizeCD  = ACD.MfgSizeCD AND OWT.MfgAttributeCD  =ACD.MfgAttributeCD   
AND OWT.SewLoc = ACD.SewPlantCD  WHERE (CONVERT(date, ACD.ANETCreatedOnDate) BETWEEN '4/14/2024' AND '4/20/2024') AND (ACD.FromPlantCD IN('86')) AND (ACD.HSCD IS NOT NULL)

--Get Attribution Sew Merge Lots
IF OBJECT_ID('tempdb..#AttributionSewMergeLots') IS NOT NULL DROP TABLE #AttributionSewMergeLots
SELECT distinct ACD.*, OWT.ApparelNETWorkLot  into #AttributionSewMergeLots  FROM   #AttributionLotsWithWO AS ACD  INNER JOIN Manufacturing.dbo.ANETOriginalWorkOrderTrace 
as OWT WITH (NOLOCK) ON  OWT.OriginalWorkOrder  = ACD.OriginalWorkOrder AND  OWT.MfgStyle  = ACD.MfgStyleCD AND convert(varchar(3),OWT.MfgColor)  = convert(varchar(3),ACD.MfgColorCD) 
AND OWT.MfgSizeCD  = ACD.MfgSizeCD AND OWT.MfgAttributeCD  =ACD.MfgAttributeCD   AND ACD.SewPlantCD = OWT.SewLoc  WHERE OWT.AttrSewLoc = '   '

--Last Transaction
IF OBJECT_ID('tempdb..#lastTrans') IS NOT NULL DROP TABLE #lastTrans
SELECT ACD.Manifest, ACD.WorkOrder, MAX(ACD.ANETCreatedOnDate) AS lastUpdate  INTO #lastTrans  FROM Manufacturing.dbo.ANETCOODetails AS ACD  WITH (NOLOCK)  
INNER JOIN  #AttributionSewMergeLots AS M on ACD.WorkOrder = m.ApparelNETWorkLot  and ACD.FromPlantCD = m.SewPlantCD  WHERE (CONVERT(date, ANETCreatedOnDate) <= '4/20/2024')  
AND (HSCD IS NOT NULL) GROUP BY Manifest, ACD.WorkOrder

--Trasmitted Manifest
IF OBJECT_ID('tempdb..#TrasmittedManifest') IS NOT NULL DROP TABLE #TrasmittedManifest
SELECT CD.WorkOrder, CD.MfgStyleCD, CD.MfgColorCD, CD.MfgSizeCD, CD.MfgSizeDESC, CD.MfgAttributeCD, CD.MfgRevisionNO, CD.PkgStyleCD, CD.PkgColorCD, CD.PkgSizeCD, 
CD.PkgSizeDESC, CD.PkgAttributeCD, CD.PkgRevisionNO, CD.SelStyleCD, CD.SelColorCD, CD.SelSizeCD, CD.SelSizeDESC, CD.SelAttributeCD, CD.SelRevisionNO, CD.ParentWorkOrder, 
CD.AssortmentParentWO, CD.FromPlantCD, CD.CutPlantCD, CD.SewPlantCD, CD.Lot, CD.ParentLot,  lastTrans.lastUpdate, SUM(round(Convert(Decimal(10, 4), CD.PiecesApplied) / Convert(Decimal(10, 4), 12), 4)) 
As Doz  ,convert(varchar(3), CD.MfgColorCD) as MfgColorCDWithThreeCharacters  into #TrasmittedManifest  FROM #lastTrans AS lastTrans INNER JOIN  Manufacturing.dbo.ANETCOODetails AS CD WITH (NOLOCK) ON 
lastTrans.Manifest = CD.Manifest AND lastTrans.WorkOrder = CD.WorkOrder AND lastTrans.lastUpdate = CD.ANETCreatedOnDate  GROUP BY CD.WorkOrder, CD.MfgStyleCD, CD.MfgColorCD, CD.MfgSizeCD, CD.MfgSizeDESC, 
CD.MfgAttributeCD, CD.MfgRevisionNO, CD.PkgStyleCD, CD.PkgColorCD, CD.PkgSizeCD, CD.PkgSizeDESC, CD.PkgAttributeCD, CD.PkgRevisionNO, CD.SelStyleCD, CD.SelColorCD, CD.SelSizeCD, CD.SelSizeDESC, 
CD.SelAttributeCD, CD.SelRevisionNO, CD.ParentWorkOrder, CD.AssortmentParentWO, CD.FromPlantCD, CD.CutPlantCD, CD.SewPlantCD,CD.SewPlantCD, CD.Lot, CD.ParentLot,  lastTrans.lastUpdate

--Work Order Info
IF OBJECT_ID('tempdb..#WOInfo') IS NOT NULL DROP TABLE #WOInfo
select DISTINCT TM.WorkOrder ,OWT.GlobalWorkOrderID, CONVERT(INT,AWO.PriorityCD)  AS PriorityCD  INTO #WOInfo  from #TrasmittedManifest as TM  INNER 
JOIN Manufacturing.dbo.ANETWorkOrders AS AWO WITH (NOLOCK)  ON TM.WorkOrder = AWO.WorkOrder AND TM.MfgStyleCD = AWO.StyleCD AND convert(varchar(3),TM.MfgColorCD) = 
convert(varchar(3),AWO.ColorCD) AND TM.MfgSizeCD = AWO.SizeCD AND TM.MfgAttributeCD =AWO.AttributeCD  INNER JOIN Manufacturing.dbo.ANETOriginalWorkOrderTrace as 
OWT WITH (NOLOCK)  ON AWO.GlobalWorkOrderID  = OWT.GlobalWorkOrderID

--Transmitted WorkLots
IF OBJECT_ID('tempdb..#TransmittedWorkLots') IS NOT NULL DROP TABLE #TransmittedWorkLots
select TM.WorkOrder , TM.ParentWorkOrder, TM.AssortmentParentWO, TM.FromPlantCD, TM.CutPlantCD,   TM.SewPlantCD,SUM(TM.Doz) AS DOZ, OWT.OriginalWorkOrder, 
OWT.ApparelNETWorkLot,  TM.MfgStyleCD AS MfgStyle, TM.MfgColorCD AS MfgColor, TM.MfgSizeCD AS MfgSizeCD, TM.MfgSizeDesc AS MfgSizeDesc, TM.PkgStyleCD AS PkgStyle, 
TM.PkgColorCD AS PkgColor,  TM.PkgSizeCD , TM.PkgSizeDesc, TM.SelStyleCD AS SellStyle, TM.SelColorCD AS SellColor, TM.SelSizeCD AS SellSizeCD,  TM.SelSizeDESC AS SellSizeDesc, 
TM.CutPlantCD AS CutLoc, TM.SewPlantCD AS SewLoc,  ISNULL(TRY_CONVERT(int,OWT.OriginalWorkOrder),0) as WO_NUMBER  ,MAX(TM.lastUpdate) AS lastUpdate,  ISNULL(TM.AssortmentParentWO, TM.WorkOrder) 
As WorkLot  ,WO.PriorityCD  INTO #TransmittedWorkLots  from #TrasmittedManifest as TM  INNER JOIN #WOInfo AS WO on TM.WorkOrder =  wo.WorkOrder  LEFT JOIN Manufacturing .dbo.ANETOriginalWorkOrderTrace 
as OWT WITH (NOLOCK)  ON WO.GlobalWorkOrderID  = OWT.GlobalWorkOrderID  GROUP BY TM.WorkOrder, TM.ParentWorkOrder, TM.AssortmentParentWO, TM.FromPlantCD, TM.CutPlantCD,  TM.SewPlantCD, 
OWT.OriginalWorkOrder, OWT.ApparelNETWorkLot,  TM.MfgStyleCD, TM.MfgColorCD, TM.MfgSizeCD, TM.MfgSizeDesc, TM.PkgStyleCD, TM.PkgColorCD,  TM.PkgSizeCD , TM.PkgSizeDesc, TM.SelStyleCD, TM.SelColorCD, 
TM.SelSizeCD,  TM.SelSizeDESC , TM.CutPlantCD, TM.SewPlantCD  ,WO.PriorityCD  ORDER BY TM.WorkOrder 

--Transmitted WorklotsbyDate
IF OBJECT_ID('tempdb..#TransmittedWorkLotsByDate') IS NOT NULL DROP TABLE #TransmittedWorkLotsByDate
select TW.WorkOrder , TW.ParentWorkOrder, TW.AssortmentParentWO, TW.FromPlantCD, TW.CutPlantCD, TW.SewPlantCD,TW.DOZ, TW.OriginalWorkOrder, TW.ApparelNETWorkLot, 
TW.MfgStyle, TW.MfgColor, TW.MfgSizeCD, TW.MfgSizeDesc, TW.PkgStyle, TW.PkgColor,TW.PkgSizeCD, TW.PkgSizeDesc, TW.SellStyle, TW.SellColor, TW.SellSizeCD,TW.SellSizeDesc, 
TW.CutLoc, TW.SewLoc, TW.WO_NUMBER,TW.WorkLot, MAX(WOAD.ANETCreatedOnDate) as TrasmittedDate  INTO #TransmittedWorkLotsByDate  from #TransmittedWorkLots as TW left join 
Manufacturing.dbo.ANETWorkOrderActionDetails as WOAD on  TW.WorkLot = WOAD.WorkOrder  WHERE CONVERT(DATE,WOAD.ANETCreatedOnDate) <='4/20/2024'  AND WOAD.ActionCD = 'SH' AND 
WOAD.Quantity > 0  GROUP BY TW.WorkOrder , TW.ParentWorkOrder, TW.AssortmentParentWO, TW.FromPlantCD, TW.CutPlantCD,  TW.SewPlantCD,TW.DOZ, TW.OriginalWorkOrder, 
TW.ApparelNETWorkLot, TW.MfgStyle, TW.MfgColor, TW.MfgSizeCD, TW.MfgSizeDesc, TW.PkgStyle, TW.PkgColor,  TW.PkgSizeCD, TW.PkgSizeDesc, TW.SellStyle, TW.SellColor, TW.SellSizeCD,  
TW.SellSizeDesc, TW.CutLoc, TW.SewLoc, TW.WO_NUMBER,TW.WorkLot

--Initial Transaction
IF OBJECT_ID('tempdb..#InitialTrans') IS NOT NULL DROP TABLE #InitialTrans
SELECT *  into #InitialTrans FROM (  SELECT TWL.WorkOrder,ActionCD,Max(ANETCreatedOnDate) AS DT, 'wo' as obs  FROM #TransmittedWorkLotsByDate  AS TWL 
INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.WorkOrder = WOD.WorkOrder WHERE ActionCD ='IT' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate 
GROUP BY TWL.WorkOrder,ActionCD  Union  SELECT TWL.WorkOrder,ActionCD,Max(ANETCreatedOnDate) AS DT, 'assort' as obs  FROM #TransmittedWorkLotsByDate AS TWL 
INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.AssortmentParentWO = WOD.WorkOrder WHERE ActionCD ='IT' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate 
GROUP BY TWL.WorkOrder,ActionCD  Union  SELECT TWL.WorkOrder,ActionCD,Max(ANETCreatedOnDate) AS DT, 'Pare' as obs  FROM #TransmittedWorkLotsByDate AS TWL 
INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.ParentWorkOrder = WOD.WorkOrder WHERE ActionCD ='IT' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate 
GROUP BY TWL.WorkOrder,ActionCD)  AS InitialTrans 

--Final Transaction
IF OBJECT_ID('tempdb..#EndTrans') IS NOT NULL DROP TABLE #EndTrans
SELECT *  into #EndTrans FROM (  SELECT TWL.WorkOrder,ActionCD,Max(ANETCreatedOnDate) AS DT, 'wo' as obs  FROM #TransmittedWorkLotsByDate AS TWL 
INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.WorkOrder = WOD.WorkOrder WHERE ActionCD ='OS' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate 
GROUP BY TWL.WorkOrder,ActionCD  Union  SELECT TWL.WorkOrder,ActionCD,Max(ANETCreatedOnDate) AS DT, 'assort' as obs  FROM #TransmittedWorkLotsByDate AS TWL 
INNER JOIN dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.AssortmentParentWO = WOD.WorkOrder WHERE ActionCD ='OS' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate 
GROUP BY TWL.WorkOrder,ActionCD  Union  SELECT TWL.WorkOrder,ActionCD,Max(ANETCreatedOnDate) AS DT, 'Pare' as obs  FROM #TransmittedWorkLotsByDate AS TWL INNER JOIN 
dbo.ANETWorkOrderActionDetails as WOD WITH (NOLOCK) ON TWL.ParentWorkOrder = WOD.WorkOrder WHERE ActionCD ='OS' AND Quantity>0 AND ANETCreatedOnDate<= TWL.TrasmittedDate 
GROUP BY TWL.WorkOrder,ActionCD)  AS EndTrans

--Transaction Info
IF OBJECT_ID('tempdb..#TransactionInfo') IS NOT NULL DROP TABLE #TransactionInfo
select DISTINCT *,  'ID' AS InitialTrans,  ISNULL((SELECT dt FROM #InitialTrans WHERE WorkOrder = TWL.WorkOrder AND obs ='WO'),  ISNULL((SELECT dt FROM #InitialTrans 
WHERE WorkOrder = TWL.WorkOrder AND obs ='assort'),(SELECT dt FROM #InitialTrans WHERE WorkOrder = TWL.WorkOrder AND obs ='Pare'))) AS InitialDate,  'FQ' AS EndTrans,  
ISNULL((SELECT dt FROM #EndTrans  WHERE WorkOrder = TWL.WorkOrder AND obs ='WO'),  ISNULL((SELECT dt FROM #EndTrans WHERE WorkOrder = TWL.WorkOrder AND obs ='assort'),
(SELECT dt FROM #EndTrans WHERE WorkOrder = TWL.WorkOrder AND obs ='Pare'))) AS EndDate,  TWL.DOZ  AS Quantity  into #TransactionInfo  from #TransmittedWorkLots as TWL

--Get Current Work Order
IF OBJECT_ID('tempdb..#CurrentWO') IS NOT NULL DROP TABLE #CurrentWO
SELECT  Manufacturing.dbo.TTSCutOrderRoutingDetail.WorkOrderNumber, MAX(Manufacturing.dbo.TTSCutOrderRoutingDetail.CutOrderSequence) AS CutOrderSequence, 
RIGHT('000000' + CONVERT(varchar, Manufacturing.dbo.TTSCutOrderRoutingDetail.WorkOrderNumber), 6)AS CurrentWorkOrderNumber, SellStyle , MfgStyle, MfgColor, MfgSizeDesc, SizeDescription   
INTO #CurrentWO  FROM  Manufacturing.dbo.TTSCutOrderRoutingDetail WITH (NOLOCK) INNER JOIN #TransactionInfo AS TransmittesdOrder ON Manufacturing.dbo.TTSCutOrderRoutingDetail.WorkOrderNumber = 
TransmittesdOrder.WO_NUMBER AND Manufacturing.dbo.TTSCutOrderRoutingDetail.GarmentStyle = TransmittesdOrder.MfgStyle AND convert(varchar(3),Manufacturing.dbo.TTSCutOrderRoutingDetail.GarmentColor) = 
convert(varchar(3),TransmittesdOrder.MfgColor)  And TransmittesdOrder.MfgSizeCD = Manufacturing.dbo.TTSCutOrderRoutingDetail.SizeDescription 
GROUP BY Manufacturing.dbo.TTSCutOrderRoutingDetail.WorkOrderNumber, SellStyle,MfgStyle,MfgColor,MfgSizeDesc,SizeDescription

--Get LeadTime Info
IF OBJECT_ID('tempdb..#LeadTimeInfo') IS NOT NULL DROP TABLE #LeadTimeInfo
select DISTINCT TI.CutLoc as CutPlant,TI.SewPlantCD as SewPlant   , CONVERT(varchar, CONVERT(date, CONVERT(VARCHAR(8), CORD.CutDueDate)), 101) 
AS CutDueDate  ,CONVERT(varchar, CONVERT(date, CONVERT(VARCHAR(8), CORD.SeDueDate)), 101) AS SewDueDate  ,CONVERT(varchar, CONVERT(date, CONVERT(VARCHAR(8), CORD.DCDueDate)), 101) 
AS DCDueDate  ,TI.OriginalWorkOrder AS WorkOrder,  CORD.OriginalTTSWO AS OriginalWO,  TI.WorkOrder AS WorkLot,  ISNULL(TI.PriorityCD ,ISNULL(CORD.Priority,0)) AS [Priority],  TI.SellStyle 
AS SellingStyle  ,TI.MfgStyle AS MFGStyle  ,TI.MfgColor AS MFGColor  ,TI.MfgSizeDesc AS MFGSize,  'IT' AS InitialTransCode  ,TI.InitialDate, 'OS' AS EndTransCode,  TI.EndDate ,TI.Quantity 
AS Doz  ,round(CONVERT(decimal(30,2),DATEDIFF (Second, TI.InitialDate, TI.EndDate)) / CONVERT(decimal(30,2), 86400),2) AS LTDays  INTO #LeadTimeInfo  From  #CurrentWO as co INNER JOIN  
Manufacturing.dbo.TTSCutOrderRoutingDetail AS CORD WITH (NOLOCK) ON CO.WorkOrderNumber  = CORD.WorkOrderNumber  AND  CO.CutOrderSequence = CORD.CutOrderSequence  
RIGHT JOIN #TransactionInfo as TI ON CORD.WorkOrderNumber = TI.WO_NUMBER AND  CORD.GarmentStyle = TI.mfgStyle And Convert(varchar(3), CORD.GarmentColor) = Convert(varchar(3), TI.MfgColor)  
And CORD.SizeDescription = TI.MfgSizeCD  ORDER BY WorkLot

--FinalQuery
SELECT ISNULL(CutPlant,'') as CutPlant, SewPlant , ISNULL(CutDueDate,'09/09/1999') as CutDueDate, ISNULL(SewDueDate,'09/09/1999') as SewDueDate, ISNULL(DCDueDate,'09/09/1999') AS DCDueDate, 
ISNULL(TRY_CONVERT(BIGINT,WorkOrder),0) as WorkOrder, 
convert(BIGINT, isnull(OriginalWO,ISNULL(TRY_CONVERT(BIGINT,OriginalWO),0))) as OriginalWO, 
isnull(WorkLot,'') as WorkLot, [Priority], isnull(SellingStyle,'') as SellingStyle, isnull(MFGStyle,'') as MFGStyle, isnull(MFGColor,'') as MFGColor, isnull(MFGSize,'') as MFGSize, 
isnull(InitialTransCode,'') as InitialTransCode, ISNULL(InitialDate,'09/09/1999 15:46:30') AS  InitialDate 
,isnull(EndTransCode,'') as EndTransCode, ISNULL(EndDate,'09/09/1999 15:46:30') AS EndDate, Doz 
,ISNULL(LTDays,0) AS LTDays 
,CASE WHEN InitialDate is null OR EndDate is null THEN 'Exclude' ELSE CASE WHEN LTDays < 0 THEN 'Exclude' ELSE 'Include' END END AS [ToConsider?] 
,P.[PlantDESC] AS SewPlantName 
FROM #LeadTimeInfo as LT 
LEFT JOIN Manufacturing.dbo.ANETFacilities as P with (nolock)  on LT.SewPlant = P.PlantCD 
ORDER BY SellingStyle