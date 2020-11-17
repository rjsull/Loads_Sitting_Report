#Import Libraries
library(tidyverse)
library(RODBC)
library(openxlsx)

###Connect to TMW Suite Replication
dbhandle <- odbcDriverConnect('driver={SQL Server};server=NFIV-SQLTMW-04;database=TMWSuite;trusted_connection=true')

###Contains SQL for G&P Assigned Trailers Sitting on Yards
loadsSitting <- sqlQuery(dbhandle, "
select top 1000
[Trl] = trl_number,
--[Type1] = trl_type1,
[Type] = name,
[PC#] = trl_branch,
--trl_terminal,
[Sts] = trl_status,
[TrlSchDate] = CONVERT(VARCHAR(20),trl_sch_date,120),
[DestCmp] = trl_sch_cmp_id,
[DestCty] = c2.cty_nmstct,
[SchSts] = trl_sch_status,
[TrlAvailDate] = CONVERT(VARCHAR(20),trl_avail_date,120),
[NextAvailCmp] = trl_avail_cmp_id,
[NextAvailCity] = c3.cty_nmstct,
[LastEvent] = trl_prior_event,
[PriorCmp] = trl_prior_cmp_id,
[PriorCity] = c1.cty_nmstct,
[LegEndDate] = lh.LegEndDate,
[CurrentTime] = GETDATE(),

--[DwellDays] = DATEDIFF(day, lh.LegEndDate, GETDATE()) ,
--[DwellHours] = DATEDIFF(hour, lh.LegEndDate, GETDATE()),
[DwellMins] = DATEDIFF(minute, lh.LegEndDate, GETDATE())


FROM trailerprofile
LEFT JOIN (SELECT labeldefinition,
	abbr,
	name,
	userlabelname
	FROM labelfile 
	WHERE labeldefinition = 'TrlType1'
) l ON l.abbr = trailerprofile.trl_type1
LEFT JOIN city c1 on trl_prior_city = c1.cty_code
LEFT JOIN city c2 on trl_sch_city = c2.cty_code
LEFT JOIN city c3 on trl_avail_city = c3.cty_code
LEFT JOIN (select lgh_primary_trailer,
			MAX(lgh_enddate) AS LegEndDate
			FROM legheader LEFT JOIN trailerprofile ON trl_number = lgh_primary_trailer
			WHERE lgh_outstatus = 'CMP'
			/*
			AND (trailerprofile.trl_branch IN ('570','571','572','573','574','580','581','586')
			OR cmp_id_end IN ('GPCOSC',
								'GPATGA',
								'GPGRSC',
								'GPNASC',
								'GPANSC',
								'GPCHSC',
								'GPDYCH',
								'GPCLNC',
								'GPFLSC',
								'GPSAGA',
								'GPCCHA',
								'GPCHTN',
								'GPHALA',
								'GPLATX')
								)*/
			group by lgh_primary_trailer) lh ON lh.lgh_primary_trailer = trl_number
WHERE
trl_prior_event = 'DLT'
--and trl_prior_cmp_id = 'GPGRSC'
AND (trailerprofile.trl_terminal IN ('570','571','572','573','574','580','581','586')
OR trl_prior_cmp_id IN ('GPCOSC',
						'GPATGA',
						'GPGRSC',
						'GPNASC',
						'GPANSC',
						'GPCHSC',
						'GPDYCH',
						'GPCLNC',
						'GPFLSC',
						'GPSAGA',
						'GPCCHA',
						'GPCHTN',
						'GPHALA',
						'GPLATX'))
AND trl_status NOT IN ('OUT','AVL')
--and trl_number <> 'B5013'
AND trl_number NOT LIKE '%DUM%'
ORDER BY c1.cty_nmstct,trl_prior_cmp_id,trl_sch_date ASC
")

###Close DB connection
odbcClose(dbhandle)

###Create summary sheet
df <- data.frame(loadsSitting)
df <- df %>% 
  mutate(DwellTime = round(DwellMins/60/24,digits = 2))
df <- select(df,-DwellMins)
df
trls <- df %>% 
  group_by(PriorCmp,PriorCity) %>% 
  summarise("Trailers"= n(), AvgDwell = round(mean(DwellTime),2),MaxDwell = max(DwellTime))
trls

col_index <- which(data.frame(colnames(df)) == "LegEndDate")
col_index2 <- which(data.frame(colnames(df)) == "CurrentTime")


###Format and Export as Local Excel File 
wb <- createWorkbook(creator = ifelse(.Platform$OS.type == "windows", Sys.getenv("USERNAME"), Sys.getenv("USER")))
sheet1 <- "Loads Sitting Report"
sheet2 <- "Trailer Count by Terminal"
n <- ncol(df)
n2 <- ncol(trls)
addWorksheet(wb, sheet1)
addWorksheet(wb, sheet2)
writeData(wb, sheet1, df, startCol = 1, startRow = 1, colNames = TRUE, rowNames = FALSE)
writeData(wb, sheet2, trls, startCol = 1, startRow = 1, colNames = TRUE, rowNames = FALSE)
addFilter(wb, sheet1, row = 1, cols = 1:n)
#addFilter(wb, sheet2, row = 1, cols = 1:n2)
freezePane(wb, sheet1, firstRow = TRUE)
freezePane(wb, sheet2, firstRow = TRUE)
setColWidths(wb, sheet1, cols = 1:n, widths = "auto")
setColWidths(wb, sheet2, cols = 1:n2, widths = "auto")
setColWidths(wb, sheet1, cols = col_index, widths = 18)
setColWidths(wb, sheet1, cols = col_index2, widths = 18)
saveWorkbook(wb, file = "C:/Users/tollenaard/Documents/G_P_Loads_Sitting_Report.xlsx", overwrite = TRUE)