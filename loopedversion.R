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

dt_ICB_all <- dbGetQuery(con,
                     "Select a.*, b.Abbreviated
                  from EAT_Reporting_BSOL.[Development].[BSOL_1255_CVDP_Data] a left join
                  EAT_Reporting_BSOL.[Development].[BSOL_1255_CVDP_Data_ICBs] b on a.AreaCode = b.AreaCode
                  WHERE TimePeriodName = 'To June 2023' and MetricCategoryName = 'Persons'"

)

# Check

# Indicators
inds <- sort(unique(dt_PCN$IndicatorCode) )

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

  # ICB
  sc_calc_ICB<-
    dt_ICB_all %>%
    filter(IndicatorCode == i & AreaType == 'ICB') %>%
    summarise(sc_min = max(0, round((0.8*min(Value, na.rm = TRUE))-5,-1)),
              sc_max = min(100, round((1.1*max(Value, na.rm = TRUE))+5,-1))
    ) %>%
    unlist()
  print(i)

  a<- dt_ICB_all %>%
    filter(IndicatorCode == i & AreaType == 'ICB') %>%
    arrange(Value) %>%
    mutate(AreaName=factor(AreaName, levels=AreaName),
           BSOL = ifelse(AreaCode == 'E54000055',1,0)) %>%
    ggplot(aes(x=AreaName, y= Value))+
    geom_col(position = position_identity(), aes(fill=factor(BSOL)), show.legend = FALSE)+
    geom_hline(data=filter(dt_ICB_all, IndicatorCode == i, AreaType == 'CTRY')
               , aes(yintercept=Value), col="red", linewidth = 2)+
    # geom_hline(data=filter(dt_ICB, IndicatorCode == "CVDP002AF")
    #            , aes(yintercept=UpperConfidenceLimit), col="red", linetype="dashed")+
    # geom_hline(data=filter(dt_ICB, IndicatorCode == "CVDP002AF")
    #            , aes(yintercept=LowerConfidenceLimit), col="red", linetype="dashed")+
    scale_fill_manual(values = c("#4fbff0", "#fc8700"))+
    scale_y_continuous("Percentage"
                       , limits = sc_calc_ICB
                       , na.value = 0)+
    scale_x_discrete("ICB")+
    labs(subtitle = lng_title) +
    theme_minimal() +
    theme(axis.text.x=element_blank()
          , plot.subtitle = element_text(face = "italic", size = 18)
          , axis.text = element_text(size = 18)
          , axis.title = element_text(size = 18)
          #, plot.subtitle = element_text(face = "italic", size = 12)
          #, plot.subtitle = element_text(face = "italic", size = 12)
    )

  a <- rvg::dml(ggobj=a)


  # PCN
  # scale calc
  sc_calc<-
    dt_PCN %>%
    filter(IndicatorCode == "CVDP006HYP") %>%
    summarise(sc_min = max(0, round((0.8*min(Value, na.rm = TRUE))-5,-1)),
              sc_max = min(100, round((1.1*max(Value, na.rm = TRUE))+5,-1)),
    ) %>%
    unlist()

  print(i)

  b<-
    dt_PCN %>%
    filter(IndicatorCode == i) %>%
    arrange(Value) %>%
    mutate(AreaName=factor(AreaName, levels=AreaName)) %>%
    ggplot(aes(x=AreaName, y= Value))+
    geom_col(position = position_identity(), fill="#8cedab")+
    geom_hline(data=filter(dt_ICB, IndicatorCode == i)
               , aes(yintercept=Value), col="red", linewidth = 2)+
    # geom_hline(data=filter(dt_ICB, IndicatorCode == "CVDP002AF")
    #            , aes(yintercept=UpperConfidenceLimit), col="red", linetype="dashed")+
    # geom_hline(data=filter(dt_ICB, IndicatorCode == "CVDP002AF")
    #            , aes(yintercept=LowerConfidenceLimit), col="red", linetype="dashed")+
    scale_y_continuous("Percentage"
                       , limits = sc_calc
                       , na.value = 0)+
    scale_x_discrete("PCN")+
    labs(subtitle = lng_title) +
    theme_minimal() +
    theme(axis.text.x=element_blank()
          , plot.subtitle = element_text(face = "italic", size = 18)
          , axis.text = element_text(size = 18)
          , axis.title = element_text(size = 18)
          #, plot.subtitle = element_text(face = "italic", size = 12)
          #, plot.subtitle = element_text(face = "italic", size = 12)
    )

  b <- rvg::dml(ggobj=b)




  # Add a slide
  my_pres2 <-add_slide(my_pres2, layout = "1_Normal Slide Picture", master="21_BasicWhite")

  # Add title
  my_pres2 <- ph_with(my_pres2, value = paste("ICB vs. England -", sh_title), location =  ph_location_label("Slide Title"))

  # Add plot
  my_pres2 <- ph_with(my_pres2, value = a, location = ph_location_type("pic"))

  # Text
  min_ICB <- filter(dt_ICB_all, IndicatorCode==i & AreaType == 'CTRY') %>% summarise(min(Value)) %>%  pull()
  max_ICB <- filter(dt_ICB_all, IndicatorCode==i & AreaType == 'CTRY') %>% summarise(min(Value))  %>%  pull()
  ENG_val <- filter(dt_ICB_all, IndicatorCode==i & AreaType == 'CTRY') %>% select(Value) %>%  pull()
  BSOL_val <- filter(dt_ICB, IndicatorCode==i) %>% select(Value) %>%  pull()
  ICB_status <- ifelse(BSOL_val > ENG_val, "higher than", ifelse(BSOL_val < ENG_val, "lower than", "the same as"))

  txt_val_ICB <- paste0("ICBs range from ",
                   min_ICB,
                   " to ",
                   max_ICB,
                   ", with a BSOL value of ",
                   ENG_val,
                   ".  BSOL is ",
                   ICB_status,
                   " the England average.")
  # Add commentary
  my_pres2 <- ph_with(my_pres2, value = txt_val_ICB
                     , location = ph_location_label("Commentary"))

 # PCN
  # Add a slide
  my_pres2 <-add_slide(my_pres2, layout = "1_Normal Slide Picture", master="21_BasicWhite")

  # Add title
  my_pres2 <- ph_with(my_pres2, value = paste("PCNs vs. BSOL -", sh_title), location =  ph_location_label("Slide Title"))

  # Add plot
  my_pres2 <- ph_with(my_pres2, value = b, location = ph_location_type("pic"))

  # Text
  min_PCN <- filter(dt_ICB, IndicatorCode==i) %>% summarise(min(Value)) %>%  pull()
  max_PCN <- filter(dt_ICB, IndicatorCode==i) %>% summarise(min(Value))  %>%  pull()
  #BSOL_val <- filter(dt_ICB, IndicatorCode==i) %>% select(Value) %>%  pull()

  txt_val_PCN <- paste0("PCNs in BSOL range from ",
                    min_PCN,
                    " to ",
                    max_PCN,
                    ", with a BSOL value of ",
                    BSOL_val,
                    ".")
  # Add commentary
  my_pres2 <- ph_with(my_pres2, value = txt_val_PCN
                      , location = ph_location_label("Commentary"))
}




print(my_pres2, target = "output/second_example.pptx")
