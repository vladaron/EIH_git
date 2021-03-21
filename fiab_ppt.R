##import

round_df <- function(x, digits) {
  # round all numeric variables
  # x: data frame 
  # digits: number of digits to round
  numeric_columns <- sapply(x, mode) == 'numeric'
  x[numeric_columns] <-  round(x[numeric_columns], digits)
  x
}

library(dplyr)
library(tidyverse)
system("java -version")
library(rJava)
library(xlsx)
library(readxl)
install.packages("flextable")
library(flextable)
install.packages("formattable")
library(formattable)
library(officer)
setwd("~/Documents/UCL/UCL-M1/Approf/eih_fiab/exctracted_data")
dat_ppt_fiab <- read_excel("ppt_fiab.xlsx")
dat_mdc_ppt_fiab <- dat_ppt_fiab %>% 
  mutate ( mdc95 = sem*1.96*sqrt(2)) %>% mutate(mdc95_mean = mdc95/((mean_1+mean_2)/2)) %>%
  round_df(2)



## table ppt protocol
#protocol used
protocol <- dat_mdc_ppt_fiab %>% select(Reference,Gender, Population, n_population, Reliab_type, n_Measure, rate, Break, Site, Interval, Familiarization) %>%
  unite(Pop, n_population, Population, sep = " ") 

#table protocol
protocol <- as_grouped_data(x= protocol, groups = c("Reference"))
table_protocol <-flextable::as_flextable(protocol) %>%
  bold(j = 1, i = ~ !is.na(Reference), bold = TRUE, part = "body") %>%
  bold(part = "header", bold = TRUE) %>%
  colformat_double(i = ~ is.na(Reference), j = "Pop", digits = 0, big.mark = "") %>%
  autofit()
table_protocol<- add_header_row(table_protocol, values = " PPT reliability protocols", colwidths =9) %>%
  bold(i = 1, bold = TRUE, part = "header")
table_protocol <- align(table_protocol, i= 1, part = "header", align = "center")
table_protocol <- set_header_labels(table_protocol,Pop = "Pop", Gender = "Gender", Reliab_type = "Type of reliability", n_Measure = "Number PPT", rate = "Rate PPT (kPa/s)", Break = "Break (min)" ,Site = "Site", Interval = "Interval", Familiarization = "Familiarization" ) 
small_border <- fp_border(color = "black", width = 2)
table_protocol <- border_outer(table_protocol, part="all", border = small_border )
table_protocol <- align(table_protocol, j = c( "Reliab_type", "n_Measure", "rate", "Break", "Site", "Interval", "Familiarization"), part = "all", align="center")
table_protocol <- table_protocol %>%
  bg(i = ~ Familiarization %in% "yes",
     j = c( "Familiarization"), 
     bg ="lawngreen", part = "body") %>%
  bg(i = ~ Familiarization %in% "no",
     j = c( "Familiarization"), 
     bg = "red2", part = "body")



#table values #grouped by ref

ppt_fiab <- dat_mdc_ppt_fiab %>% select(Reference, Site, Gender, mean_1, sd_1, mean_2, sd_2, ICC, ICC_CI_low, ICC_CI_up, sem, mdc95, mdc95_mean) %>% 
  unite(ICC_CI95, ICC_CI_low, ICC_CI_up, sep = " - ")

ppt_fiab <- as_grouped_data(x= ppt_fiab, groups = c("Reference"))
table_ppt_fiab <-flextable::as_flextable(ppt_fiab) %>%
  bold(j = 1, i = ~ !is.na(Reference), bold = TRUE, part = "body") %>%
  bold(part = "header", bold = TRUE) %>%
  colformat_double(i = ~ is.na(Reference), j = "Site", digits = 0, big.mark = "") %>%
  autofit()
table_ppt_fiab<- add_header_row(table_ppt_fiab, values = " Reliability PPT ", colwidths = 11) %>%
  bold(i = 1, bold = TRUE, part = "header")
table_ppt_fiab <- align(table_ppt_fiab, i= 1, part = "header", align = "center") 
table_ppt_fiab <- set_header_labels(table_ppt_fiab, Site = "Site", Gender = "Gender", mean_1 = "Mean 1", sd_1 = "SD 1", mean_2= "Mean 2", sd_2= "SD 2", ICC = "ICC", ICC_CI95 = "95% CI", sem = "SEM", mdc95= "MDC95", mdc95_mean = "MDC95 / mean" )
table_ppt_fiab <- align(table_ppt_fiab, j = c("Gender","mean_1", "sd_1", "mean_2", "sd_2","ICC", "ICC_CI95", "sem", "mdc95", "mdc95_mean"), part = "all", align = "center")
small_border <- fp_border(color = "black", width = 2)
table_ppt_fiab_ref <- border_outer(table_ppt_fiab, part="all", border = small_border )
table_ppt_fiab_ref <- table_ppt_fiab_ref %>%
  bg(i = ~ ICC <= 0.5 ,
     j = c("ICC"), 
     bg ="red", part = "body") %>%
  bg(i = ~ ICC > 0.5 & ICC <= 0.75,
     j = c("ICC"), 
     bg ="darkorange", part = "body") %>%
  bg(i = ~ ICC > 0.75 & ICC <=0.9,
     j = c("ICC"),
     bg = "deepskyblue", part= "body") %>%
  bg(i = ~ ICC> 0.9,
     j= c("ICC"),
     bg = "green", part = "body")

#table value #groupe by site


ppt_fiab <- dat_mdc_ppt_fiab %>% select(Reference, Site, Gender, mean_1, sd_1, mean_2, sd_2, ICC, ICC_CI_low, ICC_CI_up, sem, mdc95, mdc95_mean) %>% 
  unite(ICC_CI95, ICC_CI_low, ICC_CI_up, sep = " - ") %>% arrange(factor(Site, levels = c(" quad", "trap", "masseter", "temporal", "back", "wrist", "forearm", "hand" )))

ppt_fiab <- as_grouped_data(x= ppt_fiab, groups = c("Site"))
table_ppt_fiab <-flextable::as_flextable(ppt_fiab) %>%
  bold(j = 1, i = ~ !is.na(Site), bold = TRUE, part = "body") %>%
  bold(part = "header", bold = TRUE) %>%
  colformat_double(i = ~ is.na(Site), j = "Reference", digits = 0, big.mark = "") %>%
  autofit()
table_ppt_fiab<- add_header_row(table_ppt_fiab, values = " Reliability PPT ", colwidths = 11) %>%
  bold(i = 1, bold = TRUE, part = "header")
table_ppt_fiab <- align(table_ppt_fiab, i= 1, part = "header", align = "center") 
table_ppt_fiab <- set_header_labels(table_ppt_fiab, Reference = "Reference", Gender = "Gender", mean_1 = "Mean 1", sd_1 = "SD 1", mean_2= "Mean 2", sd_2= "SD 2", ICC = "ICC", ICC_CI95 = "95% CI", sem = "SEM", mdc95= "MDC95", mdc95_mean = "MDC95 / mean" )
table_ppt_fiab <- align(table_ppt_fiab, j = c("Gender", "mean_1", "sd_1", "mean_2", "sd_2","ICC", "ICC_CI95", "sem", "mdc95", "mdc95_mean"), part = "all", align = "center")
small_border <- fp_border(color = "black", width = 2)
table_ppt_fiab_site <- border_outer(table_ppt_fiab, part="all", border = small_border )
table_ppt_fiab_site <- table_ppt_fiab_site %>%
  bg(i = ~ ICC <= 0.5 ,
     j = c("ICC"), 
     bg ="red", part = "body") %>%
  bg(i = ~ ICC > 0.5 & ICC <=0.75,
     j = c("ICC"), 
     bg ="darkorange", part = "body") %>%
  bg(i = ~ ICC > 0.75 & ICC <= 0.9,
     j = c("ICC"),
     bg = "deepskyblue", part= "body") %>%
  bg(i = ~ ICC> 0.9,
     j= c("ICC"),
     bg = "green", part = "body")


##legend
qualif <- c("poor", "moderate", "good", "excellent")
ICC_value <- c(" ICC < 0.5", "0.5 < ICC < 0.75", "0.75 < ICC < 0.9", "ICC > 0.9" )

legend <- data.frame(qualif, ICC_value)
table_legend <- flextable(legend) %>%
  autofit()
table_legend <- table_legend %>%
  bg(i = 1 ,  j = 1:2, bg ="red", part = "body") %>%
  bg(i = 2, j = 1:2, bg ="darkorange", part = "body") %>%
  bg(i = 3, j = 1:2,  bg = "deepskyblue", part= "body") %>%
  bg(i =4, j= 1:2, bg = "green", part = "body") 
table_legend <- add_header_row(table_legend, values = "Reliability (Koo 2016)", colwidths = 2 ) %>%
  bold(i = 1, part = "header", bold = TRUE) %>%
  align(i=1, part = "header", align = "center")
table_legend <- set_header_labels(table_legend, qualif = "Qualification", ICC_value = "ICC value") %>%
  bold(i=1:2, part = "header", bold = TRUE)
table_legend <- align(table_legend, j = c("ICC_value"), part = "all", align = "center")

##extract power point

ppt <- read_pptx()
ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
ppt <- ph_with(ppt, value = table_protocol, location = ph_location_left())
ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
ppt <- ph_with(ppt, value = table_ppt_fiab_ref, location = ph_location_left()) 
ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
ppt <- ph_with(ppt, value = table_ppt_fiab_site, location = ph_location_left()) 
ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
ppt <- ph_with(ppt, value = table_legend, location = ph_location_left()) 
print(ppt, target = "ppt_fiab_r.pptx") 




