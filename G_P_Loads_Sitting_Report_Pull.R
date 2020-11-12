#Import Libraries
library(tidyverse)
library(RODBC)
library(openxlsx)

###Connect to TMW Suite Replication
dbhandle <- odbcDriverConnect('driver={SQL Server};
                               server=NFIV-SQLTMW-04;
                               database=TMWSuite;trusted_connection=true')

###Contains SQL for G&P Assigned Trailers Sitting on Yards
loadsSitting <- sqlQuery(dbhandle, 
                         "select top 1000
[Trl   ] = trl_number,
--[Type1   ] = trl_type1,
[Type   ] = name,
[PC#   ] = trl_branch,
--trl_terminal,
[Sts   ] = trl_status,
[TrlSchDate   ] = CONVERT(VARCHAR(20),trl_sch_date,120),
[DestCmp   ] = trl_sch_cmp_id,
[DestCty   ] = c2.cty_nmstct,
[SchSts   ] = trl_sch_status,
[TrlAvailDate   ] = CONVERT(VARCHAR(20),trl_avail_date,120),
[NextAvailCmp   ] = trl_avail_cmp_id,
[NextAvailCity   ] = c3.cty_nmstct,
[LastEvent   ] = trl_prior_event,
[PriorCmp] = trl_prior_cmp_id,
[PriorCity] = c1.cty_nmstct
from trailerprofile
left join (SELECT labeldefinition,
	abbr,
	name,
	userlabelname
	FROM labelfile 
	WHERE labeldefinition = 'TrlType1'
) l ON l.abbr = trailerprofile.trl_type1
left join city c1 on trl_prior_city = c1.cty_code
left join city c2 on trl_sch_city = c2.cty_code
left join city c3 on trl_avail_city = c3.cty_code
where
trl_prior_event = 'DLT'
--and trl_prior_cmp_id = 'GPGRSC'
and (trailerprofile.trl_terminal in ('570','571','572','573','574','580','581','586')
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
and trl_status NOT IN ('OUT','AVL')
--and trl_number <> 'B5013'
order by c1.cty_nmstct,trl_prior_cmp_id,trl_sch_date ASC")

###Close DB connection
odbcClose(dbhandle)

###Create summary sheet
df <- data.frame(loadsSitting)
trls <- df %>% 
  group_by(PriorCmp,PriorCity) %>% 
  summarise("Trailers  "= n())

###Format and Export as Local Excel File 
wb <- createWorkbook(creator = ifelse(.Platform$OS.type == "windows", Sys.getenv("USERNAME"), Sys.getenv("USER")))
sheet1 <- "Loads Sitting Report"
sheet2 <- "Trailer Count by Terminal"
n <- ncol(loadsSitting)
n2 <- ncol(trls)
addWorksheet(wb, sheet1)
addWorksheet(wb, sheet2)
writeData(wb, sheet1, loadsSitting, startCol = 1, startRow = 1, colNames = TRUE, rowNames = FALSE)
writeData(wb, sheet2, trls, startCol = 1, startRow = 1, colNames = TRUE, rowNames = FALSE)
addFilter(wb, sheet1, row = 1, cols = 1:n)
addFilter(wb, sheet2, row = 1, cols = 1:n2)
freezePane(wb, sheet1, firstRow = TRUE)
freezePane(wb, sheet2, firstRow = TRUE)
setColWidths(wb, sheet1, cols = 1:n, widths = "auto")
setColWidths(wb, sheet2, cols = 1:n2, widths = "auto")
saveWorkbook(wb, file = "C:/Users/sullivanry/Documents/Loads-Sitting-Report/G_P_Loads_Sitting_Report.xlsx", overwrite = TRUE)


