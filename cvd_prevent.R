################################################################################
# CVD Prevent data extraction
# Using this to test out connection to Powerpoint through OfficeR package

# Chris Mainey - c.mainey1@nhs.net
# 17/04/2024
################################################################################

# Connect to SQL Server.
library(DBI)
con <- dbConnect(odbc::odbc(), .connection_string = "Driver={SQL Server};server=MLCSU-BI-SQL;database=EAT_Reporting_BSOL",
                 timeout = 10)

dt <- dbGetQuery(con,
                 "Select a.*, b.Abbreviated
                  from EAT_Reporting_BSOL.[Development].[BSOL_1255_CVDP_Data] a inner join
                  EAT_Reporting_BSOL.[Development].[BSOL_1255_CVDP_Data_PCNs] b on a.AreaCode = b.AreaCode"
                 )

# Check
library(tidyverse)

# 26 indicators, all, male and female
dt %>%
  distinct(IndicatorName)


dt %>%
  filter(IndicatorCode == "CVDP002AF") %>%
  ggplot(aes(x=AreaName, y= Value, fill=factor(MetricCategoryName)))+
  geom_col()+
  scale_fill_brewer(palette = "Dark2")+
  theme_minimal() +
  facet_grid(cols = vars("MetricCategoryTypeName"))


# Powerpoint import
library(officer)

ppt_tmp <- read_pptx("docs/BSOL_CVD_PREVENT.pptx")

