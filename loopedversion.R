################################################################################
# CVD Prevent data extraction
# Using this to test out connection to Powerpoint through OfficeR package

# Chris Mainey - c.mainey1@nhs.net
# 19/04/2024
################################################################################

# Connect to SQL Server.
library(tidyverse)
library(DBI)
library(officer)
library(rvg)

con <- dbConnect(odbc::odbc(), .connection_string = "Driver={SQL Server};server=MLCSU-BI-SQL;database=EAT_Reporting_BSOL",
                 timeout = 10)

dt_PCN <- dbGetQuery(con,
                     "Select a.*, b.Abbreviated
                  from EAT_Reporting_BSOL.[Development].[BSOL_1255_CVDP_Data] a inner join
                  EAT_Reporting_BSOL.[Development].[BSOL_1255_CVDP_Data_PCNs] b on a.AreaCode = b.AreaCode
                  WHERE TimePeriodName = 'To June 2023' and MetricCategoryName = 'Persons'"
)

dt_ICB <- dbGetQuery(con,
                     "Select a.*, b.Abbreviated
                  from EAT_Reporting_BSOL.[Development].[BSOL_1255_CVDP_Data] a inner join
                  EAT_Reporting_BSOL.[Development].[BSOL_1255_CVDP_Data_ICBs] b on a.AreaCode = b.AreaCode
                  WHERE TimePeriodName = 'To June 2023' and MetricCategoryName = 'Persons'
                  and a.AreaCode = 'E54000055'"
)

# Check

# Indicators
inds <- unique(dt_PCN$IndicatorCode)

# Helpful functions
#function to wordwrap
wrapper <- function(x, ...)
{
  paste(strwrap(x, ...), collapse = "\n")
}

# Rounding function
roundUp <- function(x) 10^ceiling(log10(x))
roundDown <- function(x) 10^floor(log10(-x))

# Create presentation

my_pres2 <- read_pptx("docs/BSOL_CVD_PREVENT 2.pptx")



for(i in inds){
  # scale calc
  sc_calc<-
    dt_PCN %>%
    filter(IndicatorCode == i) %>%
    summarise(sc_min = max(0, round(min(Value)-15,-1)),
              sc_max = min(100, round(max(Value)+15,-1)),
    ) %>%
    unlist()
  print(i)

  # Title calc
  #
  sh_title <-
    dt_PCN %>%
    filter(IndicatorCode == i) %>%
    select(IndicatorShortName) %>%
    distinct() %>%
    pull()

  lng_title <-
    dt_PCN %>%
    filter(IndicatorCode == i) %>%
    select(IndicatorName) %>%
    distinct() %>%
    pull() %>%
    wrapper(width =220)

  a<-
    dt_PCN %>%
    filter(IndicatorCode == i) %>%
    arrange(Value) %>%
    mutate(AreaName=factor(AreaName, levels=AreaName)) %>%
    ggplot(aes(x=AreaName, y= Value))+
    geom_col(position = position_identity(), fill="#4fbff0")+
    geom_hline(data=filter(dt_ICB, IndicatorCode == i)
               , aes(yintercept=Value), col="red", linewidth = 2)+
    # geom_hline(data=filter(dt_ICB, IndicatorCode == "CVDP002AF")
    #            , aes(yintercept=UpperConfidenceLimit), col="red", linetype="dashed")+
    # geom_hline(data=filter(dt_ICB, IndicatorCode == "CVDP002AF")
    #            , aes(yintercept=LowerConfidenceLimit), col="red", linetype="dashed")+
    scale_y_continuous("Percentage", limits = sc_calc, na.value = 0)+
    scale_x_discrete("PCN")+
    labs(subtitle = lng_title) +
    theme_minimal() +
    theme(axis.text.x=element_blank()
          , plot.subtitle = element_text(face = "italic", size = 16)
          , axis.text = element_text(size = 14)
          #, plot.subtitle = element_text(face = "italic", size = 12)
          #, plot.subtitle = element_text(face = "italic", size = 12)
    )

  a <- rvg::dml(ggobj=a)


  # Add a slide
  my_pres2 <-add_slide(my_pres2, layout = "1_Normal Slide Picture", master="21_BasicWhite")

  # Add title
  my_pres2 <- ph_with(my_pres2, value = sh_title, location =  ph_location_label("Slide Title"))

  # Add plot
  my_pres2 <- ph_with(my_pres2, value = a, location = ph_location_type("pic"))

  # Text
  min_PCN <- filter(dt_ICB, IndicatorCode==i) %>% summarise(min(Value)) %>%  pull()
  max_PCN <- filter(dt_ICB, IndicatorCode==i) %>% summarise(min(Value))  %>%  pull()
  BSOL_val <- filter(dt_ICB, IndicatorCode==i) %>% select(Value) %>%  pull()

  txt_val <- paste0("PCNs in BSOL range from ",
                   min_PCN,
                   " to ",
                   max_PCN,
                   ", with a BSOL average value of ",
                   BSOL_val,
                   ".")
  # Add commentary
  my_pres2 <- ph_with(my_pres2, value = txt_val
                     , location = ph_location_label("Commentary"))
}




print(my_pres2, target = "output/second_example.pptx")
