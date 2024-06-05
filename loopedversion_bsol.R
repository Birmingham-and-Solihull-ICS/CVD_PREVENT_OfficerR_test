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
                     "Select a.*, b.Abbreviated, c.Locality
                      from EAT_Reporting_BSOL.[Development].[BSOL_1255_CVDP_Data] a inner join
                      EAT_Reporting_BSOL.[Development].[BSOL_1255_CVDP_Data_PCNs] b on a.AreaCode = b.AreaCode
                      left join (select distinct PCN, [PCN code], Locality from EAT_Reporting_BSOL.Reference.BSOL_ICS_PracticeMapped
                      WHERE Is_Current_Practice = 1) c ON a.AreaCode = c.[PCN code]
                      WHERE TimePeriodName = 'To September 2023' and MetricCategoryName = 'Persons'"
)

dt_PCN_ethn <- dbGetQuery(con,
                     "Select a.*, b.Abbreviated, c.Locality
                      from EAT_Reporting_BSOL.[Development].[BSOL_1255_CVDP_Data] a inner join
                      EAT_Reporting_BSOL.[Development].[BSOL_1255_CVDP_Data_PCNs] b on a.AreaCode = b.AreaCode
                      left join (select distinct PCN, [PCN code], Locality from EAT_Reporting_BSOL.Reference.BSOL_ICS_PracticeMapped) c ON a.AreaCode = c.[PCN code]
                      WHERE TimePeriodName = 'To September 2023' and MetricCategoryTypeName = 'Ethnicity'"
)

dt_ICB <- dbGetQuery(con,
                     "Select a.*, b.Abbreviated
                  from EAT_Reporting_BSOL.[Development].[BSOL_1255_CVDP_Data] a inner join
                  EAT_Reporting_BSOL.[Development].[BSOL_1255_CVDP_Data_ICBs] b on a.AreaCode = b.AreaCode
                  WHERE TimePeriodName = 'To September 2023' and MetricCategoryName = 'Persons'
                  and a.AreaCode = 'E54000055'"
)



dt_ICB_all <- dbGetQuery(con,
                     "Select a.*, b.Abbreviated
                  from EAT_Reporting_BSOL.[Development].[BSOL_1255_CVDP_Data] a left join
                  EAT_Reporting_BSOL.[Development].[BSOL_1255_CVDP_Data_ICBs] b on a.AreaCode = b.AreaCode
                  WHERE TimePeriodName = 'To September 2023' and MetricCategoryName = 'Persons'"

)

dt_ICB_bsol_dep <- dbGetQuery(con,
                              "Select a.*, b.Abbreviated
                              from EAT_Reporting_BSOL.[Development].[BSOL_1255_CVDP_Data] a inner join
                              EAT_Reporting_BSOL.[Development].[BSOL_1255_CVDP_Data_ICBs] b on a.AreaCode = b.AreaCode
                              WHERE TimePeriodName = 'To September 2023' and MetricCategoryTypeName = 'Deprivation quintile'
                              and a.AreaCode = 'E54000055'"
)

# Check

# Indicators
# Manual fix to put NHD with diabetes
inds <-
  dt_PCN %>%
  distinct(IndicatorCode) %>%
  mutate(IndicatorCode2 = ifelse(IndicatorCode=="CVDP002NDH", "CVDP002DM", IndicatorCode)) %>%
    arrange(substring(IndicatorCode2, str_length(IndicatorCode2)-2, str_length(IndicatorCode2)), IndicatorCode2) %>%
  pull(IndicatorCode)




#select(sort(unique(dt_PCN$IndicatorCode))

# Helpful functions
#function to wordwrap
wrapper <- function(x, ...)
{
  paste(strwrap(x, ...), collapse = "\n")
}


# Rounding function
roundUp <- function(x) 10^ceiling(log10(x))
roundDown <- function(x) 10^floor(log10(-x))


# Set ggplot theme general
theme_set(
      theme_minimal()+
        theme(
          text = element_text(size = 16)
      , axis.text.x=element_blank()
      , plot.subtitle = element_text(face = "italic", size = 20)
      , axis.text = element_text(size = 18)
      , axis.title = element_text(size = 18)
      , legend.text = element_text(size = 20)
      , legend.title = element_text(size = 22)
        )
)

#  geom_text ratio - bit weird: https://stackoverflow.com/questions/25061822/ggplot-geom-text-font-size-control
gg_ratio <- 8/5 # 14/5 was too big
# for some reason it started passing the size properly, so parked this.




# Create presentation
# Start building the slides from here.

my_pres2 <- read_pptx("docs/BSOL_CVD_PREVENT 2.pptx")

# Iterate through each indicator, building graphs and then inserting into slides, adding text.
for(i in inds){


  # Title calc
  #
  sh_title <-
    dt_PCN %>%
    filter(IndicatorCode == i) %>%
    mutate(sh_title = paste0(IndicatorCode,": ",IndicatorShortName)) %>%
    distinct(sh_title) %>%
    pull()

  lng_title <-
    dt_PCN %>%
    filter(IndicatorCode == i) %>%
    select(IndicatorName) %>%
    distinct() %>%
    pull() %>%
    wrapper(width =175)


  # ICB
  sc_calc_ICB<-
    dt_ICB_all %>%
    filter(IndicatorCode == i & AreaType == 'ICB') %>%
    summarise(sc_min = max(0, round((0.8*min(Value, na.rm = TRUE)))),
              sc_max = min(100, round((1.1*max(Value, na.rm = TRUE))))
    ) %>%
    unlist()

  # correct really small values to be at least 1.
  if(sc_calc_ICB[2]==0){ sc_calc_ICB[2] <- 0.5}
  print(i)

  # England value - - used in plots and logic for text
  ENG_val <- filter(dt_ICB_all, IndicatorCode==i & AreaType == 'CTRY') %>% select(Value) %>%  pull()
  # Bsol value - used in plots and logic for text
  BSOL_val <- filter(dt_ICB, IndicatorCode==i) %>% select(Value) %>%  pull()


  # ICBs nationally
  a<- dt_ICB_all %>%
    filter(IndicatorCode == i & AreaType == 'ICB') %>%
    arrange(Value) %>%
    mutate(AreaName=factor(AreaName, levels=AreaName),
           BSOL = ifelse(AreaCode == 'E54000055',TRUE,FALSE)) %>%
    ggplot(aes(x=AreaName, y= Value))+
    geom_col(position = position_identity(), aes(fill=BSOL))+
    geom_hline(yintercept=ENG_val, col="blue", linewidth = 2)+
    annotate("text", 0,ENG_val,label = "England", vjust = 1.5, hjust=0, col="blue", size= 8)+

    scale_fill_manual(values = c("#4fbff0", "#fc8700"))+
    scale_y_continuous("Percentage"
                       , limits = sc_calc_ICB
                       , na.value = 0)+
    scale_x_discrete("ICB")+
    labs(subtitle = lng_title) +
    theme(legend.position = "bottom")


  a <- rvg::dml(ggobj=a)




  # Deprivation at ICB level - data not provided lower

  sc_calc_ICB<-
    dt_ICB_bsol_dep %>%
    filter(IndicatorCode == i) %>%
    summarise(sc_min = max(0, round((0.8*min(Value, na.rm = TRUE)))),
              sc_max = min(100, round((1.1*max(Value, na.rm = TRUE))))
    ) %>%
    unlist()

  # correct really small values to be at least 1.
  if(sc_calc_ICB[2]==0){ sc_calc_ICB[2] <- 0.5}
  print(i)

  b<-
    dt_ICB_bsol_dep  %>%
    filter(IndicatorCode == i) %>%
    ggplot()+
    geom_col(aes(x=MetricCategoryName, y= Value,  fill = MetricCategoryName), position = position_dodge(), show.legend = FALSE)+

    geom_hline(yintercept=BSOL_val, col="red", linewidth = 2)+
    annotate("text", 0,BSOL_val,label = "BSOL", vjust = -0.5, hjust=0, col="red", size= 8)+

    geom_hline(yintercept=ENG_val, col="blue", linewidth = 2)+
    annotate("text", 0,ENG_val,label = "England", vjust = 1.5, hjust=0, col="blue", size= 8)+

    scale_fill_manual("Deprivation", values = c("#8cedab","#4fbff0","#fc8700", "#031d44", "#b88ce3", "#005EB8", "#b2b7b9"))+
    scale_y_continuous("Percentage"
                       #, limits = sc_calc_ICB
                       , na.value = 0)+
    scale_x_discrete("Deprivation Quintile")+
    labs(subtitle = lng_title) +
    #facet_grid(cols = vars(MetricCategoryTypeName), scales = "free_x")+
    theme(axis.text.x = element_text(size = 18))

  b <- rvg::dml(ggobj=b)



  # PCN
  # scale calc
  sc_calc<-
    dt_PCN %>%
    filter(IndicatorCode == i) %>%
    summarise(sc_min = max(0, round((0.8*min(Value, na.rm = TRUE)))),
              sc_max = min(100, round((1.1*max(Value, na.rm = TRUE)))),
    ) %>%
    unlist()

  # correct really small values to be at least 1.
  if(sc_calc[2]==0){ sc_calc[2] <- 0.5}

  print(i)

  c<-
    dt_PCN %>%
    filter(IndicatorCode == i) %>%
    arrange(Value) %>%
    mutate(AreaName=factor(AreaName, levels=AreaName),
           `Central Locality` = ifelse(Locality == "Central", TRUE,FALSE)) %>%
    ggplot()+
    geom_col(aes(x=AreaName, y= Value, fill = Locality), position = position_identity())+

    geom_hline(yintercept=BSOL_val, col="red", linewidth = 2)+
    annotate("text", 0,BSOL_val,label = "BSOL", vjust = -0.5, hjust=0, col="red", size= 8)+

    scale_fill_manual(values = c("#8cedab","#4fbff0", "#fc8700", "#031d44", "#b88ce3", "#b2b7b9"))+
    scale_y_continuous("Percentage"
                       , limits = sc_calc
                       , na.value = 0)+
    scale_x_discrete("PCN")+
    labs(subtitle = lng_title)

  c <- rvg::dml(ggobj=c)

  #     green  light_blue      orange   deep_navy      purple    nhs_blue light_slate    charcoal       white
  # "#8cedab"   "#4fbff0"   "#fc8700"   "#031d44"   "#b88ce3"   "#005EB8"   "#b2b7b9"   "#2c2825"   "#ffffff"
  #unique(dt_PCN_ethn_central$AreaName)

  # Added error handling if Ethnicity data is missing at PCN level.
  if(
    dt_PCN_ethn  %>%
    filter(IndicatorCode == i) %>%
    nrow() == 0
  ){
    d <- NA
  } else {

    # PCN
    # scale calc
    sc_calc<-
      dt_PCN_ethn %>%
      filter(IndicatorCode == i) %>%
      summarise(sc_min = max(0, round((0.8*min(Value, na.rm = TRUE)))),
                sc_max = min(100, round((1.1*max(Value, na.rm = TRUE)))),
      ) %>%
      unlist()

    # correct really small values to be at least 1.
    if(sc_calc[2]==0){ sc_calc[2] <- 0.5}

    dt_PCN_ethn %>%  distinct(AreaName)
  d<-
    dt_PCN_ethn %>%
    filter(IndicatorCode == i) %>%
    mutate(PCN_mask = case_when(
      # Central
       AreaName == "Moseley, Billesley & Yardley Wood PCN" ~ "A"
      , AreaName == "Community Care Hall Green PCN" ~ "B"
      , AreaName == "Pershore PCN" ~ "C"
      , AreaName == "Smartcare Central PCN" ~ "D"
      , AreaName == "Balsall Heath, Sparkhill & Moseley PCN" ~ "E"
      # North
      , AreaName == "North Birmingham PCN" ~ "F"
      , AreaName == "SUTTON GROUP PRACTICE PCN" ~ "G"
      , AreaName == "MMP CENTRAL AND NORTH PCN" ~ "H"
      , AreaName == "Gosk PCN" ~ "I"
      , AreaName == "Kingstanding, Erdington & Nechells PCN" ~ "J"
      , AreaName == "Alliance Of Sutton Practices PCN" ~ "K"
      # West
      , AreaName == "West Birmingham PCN" ~ "L"
      , AreaName == "People's Health Partnership PCN" ~ "M"
      , AreaName == "Swb I3 PCN" ~ "N"
      , AreaName == "Swb Modality PCN" ~ "O"
      , AreaName == "Swb Urban Health PCN" ~ "P"
      , AreaName == "PIONEERS INTEGRATED PARTNERSHIP PCN" ~ "Q"
      # East
      , AreaName == "Birmingham East Central PCN" ~ "R"
      , AreaName == "Bordesley East PCN" ~ "S"
      , AreaName == "Shard End And Kitts Green PCN" ~ "T"
      , AreaName == "Washwood Heath PCN" ~ "U"
      , AreaName == "Small Heath PCN" ~ "V"
      , AreaName == "Nechells, Saltley & Alum Rock PCN" ~ "W"
      # South
      , AreaName == "Harborne PCN" ~ "X"
      , AreaName == "Weoley And Rubery PCN" ~ "Y"
      , AreaName == "Bournville And Northfield PCN" ~ "Z"
      , AreaName == "South West Birmingham PCN" ~ "AA"
      , AreaName == "Edgbaston PCN" ~ "AB"
      , AreaName == "Quinton And Harborne PCN" ~ "AC"
      , AreaName == "South Birmingham Alliance PCN" ~ "AD"
      # Solihull
      , AreaName == "GPS HEALTHCARE PCN" ~ "AE"
      , AreaName == "Solihull Healthcare Partnership PCN" ~ "AF"
      , AreaName == "Solihull Rural PCN" ~ "AG"
      , AreaName == "North Solihull PCN" ~ "AH"
      , AreaName == "Solihull South Central PCN" ~ "AI"

    )) %>%
    ggplot()+
    geom_col(aes(x=PCN_mask, y= Value,  fill = MetricCategoryName), position = position_dodge())+
    geom_hline(yintercept=BSOL_val, col="red", linewidth = 2)+
    annotate("text", 0,BSOL_val,label = "BSOL", vjust = -0.5, hjust=0, col = "red", size= 8)+
    scale_fill_manual("Ethnicity", values = c("#8cedab","#4fbff0","#fc8700", "#031d44", "#b88ce3", "#005EB8", "#b2b7b9"))+

    scale_y_continuous("Percentage"
                       , limits = sc_calc
                       , na.value = 0)+
    scale_x_discrete("PCN")+
    labs(subtitle = lng_title) +
    facet_grid(cols = vars(PCN_mask), scales = "free_x")

   d <- rvg::dml(ggobj=d)
  }


###############################################
   # Build slides
##############################################



  # Transition
   my_pres2 <-add_slide(my_pres2, layout = "Transition", master="21_BasicWhite")
   # Add title
   my_pres2 <- ph_with(my_pres2, value = paste(sh_title), location =  ph_location_label("Title 1"))

  # Add a slide
  my_pres2 <-add_slide(my_pres2, layout = "1_Normal Slide Picture", master="21_BasicWhite")

  # Add title
  my_pres2 <- ph_with(my_pres2, value = paste("ICB vs. England -", sh_title), location =  ph_location_label("Slide Title"))

  # Add plot
  my_pres2 <- ph_with(my_pres2, value = a, location = ph_location_type("pic"))

  # Text
  min_ICB <- filter(dt_ICB_all, IndicatorCode==i & AreaType == 'ICB') %>% summarise(min(Value)) %>%  pull()
  max_ICB <- filter(dt_ICB_all, IndicatorCode==i & AreaType == 'ICB') %>% summarise(max(Value))  %>%  pull()
  # Moved to earlier for plotting
  #ENG_val <- filter(dt_ICB_all, IndicatorCode==i & AreaType == 'CTRY') %>% select(Value) %>%  pull()
  #Moved to earlier for plotting
  #BSOL_val <- filter(dt_ICB, IndicatorCode==i) %>% select(Value) %>%  pull()
  ICB_status <- ifelse(BSOL_val > ENG_val, "higher than", ifelse(BSOL_val < ENG_val, "lower than", "the same as"))

  txt_val_ICB <- paste0("ICBs range from ",
                   min_ICB,
                   " to ",
                   max_ICB,
                   ", with a BSOL value of ",
                   BSOL_val,
                   ".  BSOL is ",
                   ICB_status,
                   " the England average.")

  # Add commentary
  my_pres2 <- ph_with(my_pres2, value = txt_val_ICB
                     , location = ph_location_label("Commentary"))

  # BSOL deprivation
  # Add a slide
  my_pres2 <-add_slide(my_pres2, layout = "1_Normal Slide Picture", master="21_BasicWhite")

  # Add title
  my_pres2 <- ph_with(my_pres2, value = paste("ICB - Deprivation -", sh_title), location =  ph_location_label("Slide Title"))

  # Add plot
  my_pres2 <- ph_with(my_pres2, value = b, location = ph_location_type("pic"))

 # PCN
  # Add a slide
  my_pres2 <-add_slide(my_pres2, layout = "1_Normal Slide Picture", master="21_BasicWhite")

  # Add title
  my_pres2 <- ph_with(my_pres2, value = paste("PCNs vs. BSOL -", sh_title), location =  ph_location_label("Slide Title"))

  # Add plot
  my_pres2 <- ph_with(my_pres2, value = c, location = ph_location_type("pic"))

  # Text
  min_PCN <- filter(dt_PCN, IndicatorCode==i) %>% summarise(min(Value)) %>%  pull()
  max_PCN <- filter(dt_PCN, IndicatorCode==i) %>% summarise(max(Value))  %>%  pull()
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

  # PCN - Ethnicity
  # Add a slide
  my_pres2 <-add_slide(my_pres2, layout = "1_Normal Slide Picture", master="21_BasicWhite")

  # Add title
  my_pres2 <- ph_with(my_pres2, value = paste("Central Locality PCNs: Ethnicity -", sh_title), location =  ph_location_label("Slide Title"))

  # Added error handling stage where ethnicity is missing at PCN
  if( is.na(d[1])){
    my_pres2 <- ph_with(my_pres2, value = "Ethnicity Data not available at PCN for this indicator"
                        , location = ph_location_label("Commentary"))
  } else {
  # Add plot
  my_pres2 <- ph_with(my_pres2, value = d, location = ph_location_type("pic"))
  }

}




print(my_pres2, target = "output/bsol_example.pptx")
