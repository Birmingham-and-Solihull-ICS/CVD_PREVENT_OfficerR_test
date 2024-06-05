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
library(tidyverse)

# 26 indicators, all, male and female
ind_full <- dt_PCN %>%
            group_by(IndicatorName, TimePeriodName) %>%
            summarise(ct = n()) %>%
            pivot_wider(names_from = TimePeriodName , values_from = ct)

write_csv(ind_full, "ind_full.csv")

inds <- ind_full %>%
    left_join(ind_sept23, keep = TRUE)

write_tsv(ind_full, file = "inds_check.tsv")

#function to wordwrap
wrapper <- function(x, ...)
{
  paste(strwrap(x, ...), collapse = "\n")
}

# Rounding function
roundUp <- function(x) 10^ceiling(log10(x))
roundDown <- function(x) 10^floor(log10(-x))
# scale calc
sc_calc<-
  dt_PCN %>%
  filter(IndicatorCode == "CVDP002AF") %>%
  summarise(sc_min = max(0, round(min(Value)-15,-1)),
            sc_max = min(100, round(max(Value)+15,-1)),
            ) %>%
  unlist()


#
sh_title <-
  dt_PCN %>%
  filter(IndicatorCode == "CVDP002AF") %>%
  select(IndicatorShortName) %>%
  distinct() %>%
  pull()

lng_title <-
  dt_PCN %>%
  filter(IndicatorCode == "CVDP002AF") %>%
  select(IndicatorName) %>%
  distinct() %>%
  pull() %>%
  wrapper(width =250)

a<-
  dt_PCN %>%
  filter(IndicatorCode == "CVDP002AF") %>%
  arrange(Value) %>%
  mutate(AreaName=factor(AreaName, levels=AreaName)) %>%
  ggplot(aes(x=AreaName, y= Value))+
  geom_col(position = position_identity(), fill="#4fbff0")+
  geom_hline(data=filter(dt_ICB, IndicatorCode == "CVDP002AF")
             , aes(yintercept=Value), col="red", linewidth = 2)+
  # geom_hline(data=filter(dt_ICB, IndicatorCode == "CVDP002AF")
  #            , aes(yintercept=UpperConfidenceLimit), col="red", linetype="dashed")+
  # geom_hline(data=filter(dt_ICB, IndicatorCode == "CVDP002AF")
  #            , aes(yintercept=LowerConfidenceLimit), col="red", linetype="dashed")+
  scale_fill_brewer(palette = "Dark2")+
  scale_y_continuous("Percentage", limits = sc_calc, na.value = 0)+
  scale_x_discrete("PCN")+
  labs(subtitle = lng_title) +
  theme_minimal() +
  theme(axis.text.x=element_blank()
        , plot.subtitle = element_text(face = "italic", size = 12)
        , axis.text = element_text(size = 12)
        #, plot.subtitle = element_text(face = "italic", size = 12)
        #, plot.subtitle = element_text(face = "italic", size = 12)
        )

a <- rvg::dml(ggobj=a)

# Powerpoint import
library(officer)

my_pres <- read_pptx("docs/BSOL_CVD_PREVENT 2.pptx")
#rm(my_pres)

# What are the layouts and masters available
layout_summary(my_pres)

# title location
#rm(loc_title)
#loc_title <- ph_location_type(type = "title", newlabel = "title")

my_pres <-add_slide(my_pres, layout = "1_Normal Slide Picture", master="21_BasicWhite")
#my_pres <-add_slide(ppt_tmp, layout = "Normal Slide Blank", master="21_BasicWhite")

# my_pres <-add_slide(my_pres, layout = "1_Normal Slide Picture", master="21_BasicWhite")

my_pres <- ph_with(my_pres, value = sh_title, location =  ph_location_label("Slide Title"))

my_pres <- ph_with(my_pres, value = "Test"
                   , location = ph_location_label("Commentary"))

my_pres <- ph_with(my_pres, value = a, location = ph_location_type("pic"))



get_ph_loc()


ph_location_type(my_pres)

pptx_summary(my_pres)

plot_layout_properties(ppt_tmp, layout = "Normal Slide Blank", master="21_BasicWhite")
slide_summary(ppt_tmp,2)

ph_location_label(title)
# export and try
print(my_pres, target = "test.pptx")


plot2 <-
  iris %>%
  ggplot(aes(x=Petal.Length,y=Petal.Width, col=Species))+
  geom_point(size=2)+
  scale_colour_viridis_d()

plot2 <- rvg::dml(ggobj = plot2)


my_pres <-add_slide(my_pres, layout = "1_Normal Slide Picture", master="21_BasicWhite")

my_pres <- ph_with(my_pres, value = "Iris Data", location =  ph_location_label("Slide Title"))

my_pres <- ph_with(my_pres, value = plot2, location = ph_location_type("pic"))

# export and try
print(my_pres, target = "test2.pptx")
