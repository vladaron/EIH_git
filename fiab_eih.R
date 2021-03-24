#import
library(dplyr)
library(tidyverse)
system("java -version")
library(rJava)
library(xlsx)
library(readxl)
install.packages("flextable")
library(flextable)
library(googlesheets4)

dat_eih <- read_sheet("https://docs.google.com/spreadsheets/d/1W9XiSLdRMn9wmZ7vLMckzIeHSUxis5JfnHaW_Qj0F3Y/edit#gid=0")
dat_mdc_eih <- dat_eih %>% mutate (mdc95_abs = SEM_abs*1.96*sqrt(2), mdc95_rel = SEM_rel*1.96*sqrt(2)) %>% mutate(mdc95_abs_mean = mdc95_abs/((mean_pré_1+mean_pré_2)/2))
str(dat_mdc_eih)
library(officer)

round_df <- function(x, digits) {
  # round all numeric variables
  # x: data frame 
  # digits: number of digits to round
  numeric_columns <- sapply(x, mode) == 'numeric'
  x[numeric_columns] <-  round(x[numeric_columns], digits)
  x
}
###protocol used
protocol <- dat_mdc_eih %>% filter (Modality =="cycling") %>% select(Reference, Population, n_population, Modality, Intensity, Duration, Control, Measure, rate_ppt, number_ppt, wash_out) %>% 
  unite(Pop, n_population, Population, sep = " ") %>%
  separate(Measure, c("Modality_test", "Site"), "_")
#table protocol
protocol <- as_grouped_data(x= protocol, groups = c("Reference"))
table_protocol <-flextable::as_flextable(protocol) %>%
  bold(j = 1, i = ~ !is.na(Reference), bold = TRUE, part = "body") %>%
  bold(part = "header", bold = TRUE) %>%
  colformat_double(i = ~ is.na(Reference), j = "Pop", digits = 0, big.mark = "") %>%
  autofit()
table_protocol<- add_header_row(table_protocol, values = " EIH & Aerobic exercise protocols", colwidths = 10) %>%
  bold(i = 1, bold = TRUE, part = "header")
table_protocol <- align(table_protocol, i= 1, part = "header", align = "center")
table_protocol <- set_header_labels(table_protocol,Pop = "Pop & Gender", Modality = "Ex type", Intensity = "Intensity", Duration = "Ex time (min)", Control = "Control", Modality_test = "Test mod", Site = "Site",  rate_ppt = "Rate PPT", number_ppt = "Number PPT", wash_out = "Wash out (wk)") 
small_border <- fp_border(color = "black", width = 2)
table_protocol <- border_outer(table_protocol, part="all", border = small_border )


## ppt quad
dat_eih_quad <- dat_mdc_eih %>% 
  select (Reference, Modality, Measure, mean_pré_1, sd_pré_1,abs_change_1,sd_abs_change_1,rel_change_1, sd_rel_change_1, mean_pré_2, sd_pré_2, abs_change_2,sd_abs_change_2,rel_change_2,sd_rel_change_2, ICC_abs, ICC_abs_CI_inf, ICC_abs_CI_sup, ICC_rel, ICC_rel_CI_inf, ICC_rel_CI_sup, SEM_abs, SEM_rel,mdc95_abs, mdc95_abs_mean, mdc95_rel) %>% 
  filter(Measure == "ppt_quad") %>%
  mutate(mean_pré = ((mean_pré_1 +mean_pré_2)/2), sd_pré = (sd_pré_1 + sd_pré_2)/2,abs_change = (abs_change_1+abs_change_2)/2, sd_abs_change= (sd_abs_change_1+sd_abs_change_2)/2, rel_change = (rel_change_1+rel_change_2)/2, sd_rel_change= (sd_rel_change_1 + sd_rel_change_2)/2) %>%
  select(Reference,Modality, Measure,mean_pré, sd_pré, abs_change, sd_abs_change, rel_change, sd_rel_change, ICC_abs, ICC_abs_CI_inf, ICC_abs_CI_sup, ICC_rel, ICC_rel_CI_inf, ICC_rel_CI_sup, SEM_abs, SEM_rel,mdc95_abs, mdc95_abs_mean, mdc95_rel) %>%
  unite(ICC_abs_CI, ICC_abs_CI_inf, ICC_abs_CI_sup, sep = " ; ") %>% unite(ICC_rel_CI, ICC_rel_CI_inf, ICC_rel_CI_sup, sep=" ; ") %>%
  unite(pré_mean_sd, mean_pré, sd_pré, sep=" ± ")%>%
  unite(mean_abs_change_sd, abs_change, sd_abs_change, sep=" ± ") %>%
  unite(mean_rel_change_sd , rel_change, sd_rel_change,sep=" ± ") 
dat_eih_quad_aer <- dat_eih_quad %>% filter(Modality == "cycling")
#table_mean : quad ppt
quad_ppt <- dat_eih_quad_aer %>% select(Reference, Modality,  pré_mean_sd, mean_abs_change_sd, mean_rel_change_sd )
quad_ppt <- as_grouped_data(x= quad_ppt, groups = c("Reference"))
table_quad_ppt <-flextable::as_flextable(quad_ppt) %>%
  bold(j = 1, i = ~ !is.na(Reference), bold = TRUE, part = "body") %>%
  bold(part = "header", bold = TRUE) %>%
  colformat_double(i = ~ is.na(Reference), j = "Modality", digits = 0, big.mark = "") %>%
  autofit()
table_quad_ppt<- add_header_row(table_quad_ppt, values = " PPT at Quadriceps & Aerobic exercise ", colwidths = 4) %>%
  bold(i = 1, bold = TRUE, part = "header")
table_quad_ppt <- align(table_quad_ppt, i= 1, part = "header", align = "center")
table_quad_ppt <- set_header_labels(table_quad_ppt, Modality = "Exercise type", pré_mean_sd = "Mean pré ± SD", mean_abs_change_sd = "Mean abs change ± SD", mean_rel_change_sd = "Mean rel change ± SD (%)") 

#table_reliab : quad ppt
quad_fiab <- dat_eih_quad_aer %>% select(Reference, Modality, ICC_abs, ICC_abs_CI, ICC_rel, ICC_rel_CI, SEM_abs, SEM_rel, mdc95_abs, mdc95_abs_mean, mdc95_rel) %>% round_df(2)
quad_fiab <- as_grouped_data(x= quad_fiab, groups = c("Reference"))
table_quad_fiab <-flextable::as_flextable(quad_fiab) %>%
  bold(j = 1, i = ~ !is.na(Reference), bold = TRUE, part = "body") %>%
  bold(part = "header", bold = TRUE) %>%
  colformat_double(i = ~ is.na(Reference), j = "Modality", digits = 0, big.mark = "") %>%
  autofit()
table_quad_fiab<- add_header_row(table_quad_fiab, values = " Reliability at quadriceps & Aerobic exercise (cycling) ", colwidths = 10) %>%
  bold(i = 1, bold = TRUE, part = "header")
table_quad_fiab <- align(table_quad_fiab, i= 1, part = "header", align = "center") 
table_quad_fiab <- set_header_labels(table_quad_fiab, Modality = "Exercise type", ICC_abs = "ICC abs ", ICC_abs_CI ="CI 95%" , ICC_rel = "ICC rel ", ICC_rel_CI ="CI 95%", SEM_abs = "SEM abs", SEM_rel = "SEM rel", mdc95_abs = "mdc95 abs", mdc95_abs_mean = "mdc95 abs / mean", mdc95_rel = "mdc95 rel")
table_quad_fiab <- align(table_quad_fiab, j = c("ICC_abs", "ICC_abs_CI", "ICC_rel", "ICC_rel_CI", "SEM_abs","SEM_rel", "mdc95_abs", "mdc95_abs_mean", "mdc95_rel"), part = "all", align = "center")
table_quad_fiab <- table_quad_fiab %>% 
  bg(i = ~ ICC_abs < 0.5 ,
     j = c("ICC_abs"), 
     bg ="red", part = "body")
table_quad_fiab <- table_quad_fiab %>%
  bg(i = ~ ICC_abs > 0.5 & ICC_abs < 0.75,
     j = c("ICC_abs"), 
     bg ="darkorange", part = "body") %>%
  bg(i = ~ ICC_abs > 0.75 & ICC_abs < 0.9,
     j = c("ICC_abs"),
     bg = "deepskyblue", part= "body") %>%
  bg(i = ~ ICC_abs > 0.9,
     j= c("ICC_abs"),
     bg = "green", part = "body") %>%
  bg(i = ~ ICC_rel < 0.5 ,
     j = c("ICC_rel"), 
     bg ="red", part = "body") %>%
  bg(i = ~ ICC_rel > 0.5 & ICC_rel < 0.75,
     j = c("ICC_rel"), 
     bg ="darkorange", part = "body") %>%
  bg(i = ~ ICC_rel > 0.75 & ICC_rel <0.9,
     j = c("ICC_rel"),
     bg = "deepskyblue", part= "body") %>%
  bg(i = ~ ICC_rel > 0.9,
     j= c("ICC_rel"),
     bg = "green", part = "body")

##PPT remote aerobic
dat_eih_remote_aer <- dat_mdc_eih %>% 
  select (Reference, Modality, Measure, mean_pré_1, sd_pré_1,abs_change_1,sd_abs_change_1,rel_change_1, sd_rel_change_1, mean_pré_2, sd_pré_2, abs_change_2,sd_abs_change_2,rel_change_2,sd_rel_change_2, ICC_abs, ICC_abs_CI_inf, ICC_abs_CI_sup, ICC_rel, ICC_rel_CI_inf, ICC_rel_CI_sup, SEM_abs, SEM_rel,mdc95_abs, mdc95_abs_mean, mdc95_rel) %>% 
  filter(Measure !="ppt_quad") %>%
  mutate(mean_pré = ((mean_pré_1 +mean_pré_2)/2), sd_pré = (sd_pré_1 + sd_pré_2)/2,abs_change = (abs_change_1+abs_change_2)/2, sd_abs_change= (sd_abs_change_1+sd_abs_change_2)/2, rel_change = (rel_change_1+rel_change_2)/2, sd_rel_change= (sd_rel_change_1 + sd_rel_change_2)/2) %>%
  select(Reference,Modality, Measure,mean_pré, sd_pré, abs_change, sd_abs_change, rel_change, sd_rel_change, ICC_abs, ICC_abs_CI_inf, ICC_abs_CI_sup, ICC_rel, ICC_rel_CI_inf, ICC_rel_CI_sup, SEM_abs, SEM_rel,mdc95_abs, mdc95_abs_mean, mdc95_rel) %>%
  unite(ICC_abs_CI, ICC_abs_CI_inf, ICC_abs_CI_sup, sep = " ; ") %>% unite(ICC_rel_CI, ICC_rel_CI_inf, ICC_rel_CI_sup, sep=" ; ") %>%
  unite(pré_mean_sd, mean_pré, sd_pré, sep=" ± ")%>%
  unite(mean_abs_change_sd, abs_change, sd_abs_change, sep=" ± ") %>%
  unite(mean_rel_change_sd , rel_change, sd_rel_change,sep=" ± ") %>%
  filter(Modality=="cycling")
dat_eih_remote_aer <- dat_eih_remote_aer %>% separate(Measure, c("Modality_test", "Site"), "_")
#table mean remote
remote_ppt <- dat_eih_remote_aer %>% select(Reference, Modality, Site , pré_mean_sd, mean_abs_change_sd, mean_rel_change_sd )
remote_ppt <- as_grouped_data(x= remote_ppt, groups = c("Reference"))
table_remote_ppt <-flextable::as_flextable(remote_ppt) %>%
  bold(j = 1, i = ~ !is.na(Reference), bold = TRUE, part = "body") %>%
  bold(part = "header", bold = TRUE) %>%
  colformat_double(i = ~ is.na(Reference), j = "Modality", digits = 0, big.mark = "") %>%
  autofit()
table_remote_ppt<- add_header_row(table_remote_ppt, values = " PPT at remote sites & Aerobic exercise (cycling) ", colwidths = 5) %>%
  bold(i = 1, bold = TRUE, part = "header")
table_remote_ppt <- align(table_remote_ppt, i= 1, part = "header", align = "center")
table_remote_ppt <- set_header_labels(table_remote_ppt, Modality = "Exercise type", Site = "Site", pré_mean_sd = "Mean pré ± SD", mean_abs_change_sd = "Mean abs change ± SD", mean_rel_change_sd = "Mean rel change ± SD (%)") 

#table remote reliability
remote_fiab <- dat_eih_remote_aer %>% select(Reference, Modality, ICC_abs, ICC_abs_CI, ICC_rel, ICC_rel_CI, SEM_abs, SEM_rel, mdc95_abs, mdc95_abs_mean, mdc95_rel) %>% round_df(2)
remote_fiab <- as_grouped_data(x= remote_fiab, groups = c("Reference"))
table_remote_fiab <-flextable::as_flextable(remote_fiab) %>%
  bold(j = 1, i = ~ !is.na(Reference), bold = TRUE, part = "body") %>%
  bold(part = "header", bold = TRUE) %>%
  colformat_double(i = ~ is.na(Reference), j = "Modality", digits = 0, big.mark = "") %>%
  autofit()
table_remote_fiab<- add_header_row(table_remote_fiab, values = " Reliability at remote sites & Aerobic exercise (cycling) ", colwidths = 10) %>%
  bold(i = 1, bold = TRUE, part = "header")
table_remote_fiab <- align(table_remote_fiab, i= 1, part = "header", align = "center") 
table_remote_fiab <- set_header_labels(table_remote_fiab, Modality = "Exercise type", ICC_abs = "ICC abs ", ICC_abs_CI ="CI 95%" , ICC_rel = "ICC rel ", ICC_rel_CI ="CI 95%", SEM_abs = "SEM abs", SEM_rel = "SEM rel", mdc95_abs = "mdc95 abs", mdc95_abs_mean = "mdc95 abs / mean", mdc95_rel = "mdc95 rel")
table_remote_fiab <- align(table_remote_fiab, j = c("ICC_abs", "ICC_abs_CI", "ICC_rel", "ICC_rel_CI", "SEM_abs","SEM_rel", "mdc95_abs", "mdc95_abs_mean", "mdc95_rel"), part = "all", align = "center")
table_remote_fiab <- table_remote_fiab %>% 
  bg(i = ~ ICC_abs < 0.5 ,
     j = c("ICC_abs"), 
     bg ="red", part = "body") %>%
  bg(i = ~ ICC_abs > 0.5 & ICC_abs < 0.75,
     j = c("ICC_abs"), 
     bg ="darkorange", part = "body") %>%
  bg(i = ~ ICC_abs > 0.75 & ICC_abs <0.9,
     j = c("ICC_abs"),
     bg = "deepskyblue", part= "body") %>%
  bg(i = ~ ICC_abs > 0.9,
     j= c("ICC_abs"),
     bg = "green", part = "body") %>%
  bg(i = ~ ICC_rel < 0.5 ,
     j = c("ICC_rel"), 
     bg ="red", part = "body") %>%
  bg(i = ~ ICC_rel > 0.5 & ICC_rel < 0.75,
     j = c("ICC_rel"), 
     bg ="darkorange", part = "body") %>%
  bg(i = ~ ICC_rel > 0.75 & ICC_rel <0.9,
     j = c("ICC_rel"),
     bg = "deepskyblue", part= "body") %>%
  bg(i = ~ ICC_rel > 0.9,
     j= c("ICC_rel"),
     bg = "green", part = "body")


## kappa
kappa_eih <- dat_mdc_eih %>% select(Reference, Modality, Measure, agr_k_resp, agr_k_resp_sign) %>%
  filter(Modality=="cycling") %>% separate(Measure, c("Modality_test", "Site"), "_") %>% select(Reference,Modality, Site, agr_k_resp, agr_k_resp_sign)
# colors
customGreen0 = "#DeF7E9"
customGreen = "#71CA97"
customRed = "#ff7f7f"
#table kappa
kappa_eih <- as_grouped_data(x= kappa_eih, groups = c("Reference"))
table_kappa <-flextable::as_flextable(kappa_eih) %>%
  bold(j = 1, i = ~ !is.na(Reference), bold = TRUE, part = "body") %>%
  bold(part = "header", bold = TRUE) %>%
  colformat_double(i = ~ is.na(Reference), j = "Modality", digits = 0, big.mark = "") %>%
  autofit()
table_kappa<- add_header_row(table_kappa, values = " Agreement for all sites & Aerobic exercise (cyclin & ppt) ", colwidths =4) %>%
  bold(i = 1, bold = TRUE, part = "header")
table_kappa <- align(table_kappa, i= 1, part = "header", align = "center")
table_kappa <- set_header_labels(table_kappa, Modality = "Exercise type", Site = "Site", agr_k_resp = "Agreement EIH (k)",agr_k_resp_sign = "k significativity")
table_kappa <- align(table_kappa, j= c("Site", "agr_k_resp", "agr_k_resp_sign"), part = "all", align = "center")
table_kappa <- table_kappa %>% 
  bold(part = "header", bold = TRUE) %>%
  bg(i = ~ agr_k_resp_sign %in% "yes",
     j = c( "agr_k_resp", "agr_k_resp_sign"), 
     bg =customGreen, part = "body") %>%
  bg(i = ~ agr_k_resp_sign %in% "no",
     j = c( "agr_k_resp", "agr_k_resp_sign"), 
     bg =customRed, part = "body")


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

## Power Point extraction 
ppt <- read_pptx()
ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
ppt <- ph_with(ppt, value = table_protocol, location = ph_location_left())
ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
ppt <- ph_with(ppt, value = table_quad_ppt, location = ph_location_left()) 
ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
ppt <- ph_with(ppt, value = table_quad_fiab, location = ph_location_left()) 
ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
ppt <- ph_with(ppt, value = table_remote_ppt, location = ph_location_left()) 
ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
ppt <- ph_with(ppt, value = table_remote_fiab, location = ph_location_left()) 
ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
ppt <- ph_with(ppt, value = table_kappa, location = ph_location_left()) 
ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
ppt <- ph_with(ppt, value = table_legend, location = ph_location_left()) 
print(ppt, target = "eih_fiab_r.pptx") 




#################################################################################################################
##ppt trap/back/thenar
dat_eih_trap <- dat_mdc_eih %>% 
  select (Reference, Modality, Measure, mean_pré_1,sd_pré_1, mean_pré_2, sd_pré_2, ICC_abs, ICC_rel, SEM_abs, SEM_rel,mdc95_abs, mdc95_abs_mean, mdc95_rel) %>% 
  filter(Measure == "ppt_trap") %>%
  mutate(mean_pré = ((mean_pré_1 +mean_pré_2)/2), sd_pré = ((sd_pré_1 + sd_pré_2)/2)) %>%
  select(Reference, Modality, mean_pré, sd_pré, ICC_abs, ICC_rel, SEM_abs, SEM_rel,mdc95_abs, mdc95_abs_mean, mdc95_rel)
## ppt back
dat_eih_back <- dat_mdc_eih %>% 
  select (Reference, Modality, Measure, mean_pré_1, sd_pré_1, mean_pré_2, sd_pré_2, ICC_abs, ICC_rel, SEM_abs, SEM_rel,mdc95_abs, mdc95_abs_mean, mdc95_rel) %>% 
  filter(Measure == "ppt_back") %>%
  mutate(mean_pré = ((mean_pré_1 +mean_pré_2)/2), sd_pré = ((sd_pré_1 + sd_pré_2)/2)) %>%
  select(Reference, Modality, mean_pré, sd_pré, ICC_abs, ICC_rel, SEM_abs, SEM_rel,mdc95_abs, mdc95_abs_mean, mdc95_rel)
dat_eih_thenar <- dat_mdc_eih %>% 
  select (Reference, Modality, Measure, mean_pré_1, sd_pré_1, mean_pré_2, sd_pré_2, ICC_abs, ICC_rel, SEM_abs, SEM_rel,mdc95_abs, mdc95_abs_mean, mdc95_rel) %>% 
  filter(Measure == "ppt_thenar") %>%
  mutate(mean_pré = ((mean_pré_1 +mean_pré_2)/2), sd_pré = ((sd_pré_1 +sd_pré_2)/2)) %>%
  select(Reference, Modality, mean_pré, sd_pré, ICC_abs, ICC_rel, SEM_abs, SEM_rel,mdc95_abs, mdc95_abs_mean, mdc95_rel)


## extraction excel
dat_eih_quad
write.xlsx(dat_eih_quad, "eih_quad.xlsx")
dat_eih_trap
dat_eih_back
dat_eih_thenar
##excel extraction protocol
write.xlsx(protocol, "eih_fiab_protocol.xlsx")
#extract excel quad 
head(dat_eih_quad_aer)
write.xlsx(dat_eih_quad_aer, "dat_eih_quad_aer.xlsx")
# extract excel remote
write.xlsx(dat_eih_remote_aer, "dat_eih_remote_aer.xlsx")
#extract excel kappa
write.xlsx(kappa_eih, "kappa_eih.xlsx")