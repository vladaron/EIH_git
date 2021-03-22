##import
library(dplyr)
library(tidyverse)
system("java -version")
library(rJava)
library(xlsx)
library(readxl)
library(data.table)
library(dplyr)
install.packages("formattable")
library(formattable)
library(tidyr)
install.packages("gdtools", type = "source")
install.packages("flextable")
library(flextable)
library(officer)
library(data.table)

setwd("~/Documents/UCL/UCL-M1/Approf/eih_fiab/eih_fiab/exctracted_data")
var_eih <- read_excel("var_eih.xlsx")


##gender
gender_var_eih <- var_eih %>% filter( Variable == "gender") %>% filter(Modality_test=="ppt") %>% select(!Variable)

#aer_ex
aer_gender_var <- gender_var_eih %>% filter(Modality_ex %in% c("cycling", "treadmill")) 
#other_ex
ex_gender_var <- gender_var_eih %>% filter(Modality_ex !="cycling", Modality_ex != "treadmill") 
##fitness
fit_var_eih <- var_eih %>% filter(Variable=="fitness")%>% filter(Modality_test=="ppt")%>% select(!Variable)
#aer_ex
aer_fit_var <- fit_var_eih %>%filter(Modality_ex %in% c("cycling", "treadmill")) 
#other_ex
ex_fit_var <- fit_var_eih %>% filter(Modality_ex !="cycling", Modality_ex != "treadmill") 
##age
age_var_eih <- var_eih %>% filter(Variable=="age")%>% filter(Modality_test=="ppt")%>% select(!Variable)
#aer_ex
aer_age_var <- age_var_eih %>%filter(Modality_ex %in% c("cycling", "treadmill")) 
#other_ex
ex_age_var <- age_var_eih %>% filter(Modality_ex !="cycling", Modality_ex != "treadmill") 
##psysoc
psy_var_eih <- var_eih %>% filter(Variable == "psysoc")%>% filter(Modality_test=="ppt")%>% select(!Variable)
#aer_ex
aer_psy_var <- psy_var_eih %>%filter(Modality_ex %in% c("cycling", "treadmill")) 
#other_ex
ex_psy_var <- psy_var_eih %>% filter(Modality_ex !="cycling", Modality_ex != "treadmill") 
##menstruation
menstr_var_eih <- var_eih %>% filter(Variable == "menstrual_cycle")%>% filter(Modality_test=="ppt")%>% select(!Variable)
#aer_ex
aer_menstr_var <- menstr_var_eih %>%filter(Modality_ex %in% c("cycling", "treadmill"))
#other_ex
ex_menstr_var <- menstr_var_eih %>% filter(Modality_ex !="cycling", Modality_ex != "treadmill") 

## extraction
write.xlsx(aer_psy_var, "var_eih.xlsx", sheetName = "aer_psy_var", append=TRUE)
write.xlsx(ex_psy_var, "var_eih.xlsx", sheetName = "ex_psy_var", append=TRUE)
write.xlsx(aer_fit_var, "var_eih.xlsx", sheetName = "aer_fit_var", append=TRUE)
write.xlsx(ex_fit_var, "var_eih.xlsx", sheetName = "ex_fit_var", append=TRUE)
write.xlsx(aer_age_var, "var_eih.xlsx", sheetName = "aer_age_var", append=TRUE)
write.xlsx(ex_age_var, "var_eih.xlsx", sheetName = "ex_age_var", append=TRUE)
write.xlsx(aer_gender_var, "var_eih.xlsx", sheetName = "aer_gender_var", append=TRUE)
write.xlsx(ex_gender_var, "var_eih.xlsx", sheetName = "ex_gender_var", append=TRUE)
write.xlsx(aer_menstr_var, "var_eih.xlsx", sheetName = "aer_menstr_var", append=TRUE)
write.xlsx(ex_menstr_var, "var_eih.xlsx", sheetName = "ex_menstr_var", append=TRUE)



##flextable
# colors
customGreen0 = "#DeF7E9"
customGreen = "#71CA97"
customRed = "#ff7f7f"

# aer_gender


aer_gender_var <- as_grouped_data(x =aer_gender_var, groups = c("Reference"))
table_aer_gender <- flextable :: as_flextable(aer_gender_var) %>%
  bold(j=1, i= ~ !is.na(Reference), bold =TRUE, part="body") %>%
  italic(j=1, i= ~ !is.na(Reference),part="body") %>%
  bold(part = "header", bold = TRUE) %>%
  colformat_double(i = ~ is.na(Reference), j = "n", digits = 0, big.mark = "") %>%
  autofit()

table_aer_gender <- add_header_row(table_aer_gender, values = "Gender : Aerobic exercise & ppt ", colwidths = 9)
table_aer_gender <- table_aer_gender %>% 
  bold(part = "header", bold = TRUE) %>%
  bg(i = ~ Impact %in% "yes",
     j = c('Impact', 'Rem', "Site", "Modality_test", "Modality_ex", "Duration_ex", "Intensity_ex"), 
     bg =customGreen, part = "body") %>%
  bg(i = ~ Impact %in% "no",
     j = c('Impact', 'Rem', "Site", "Modality_test", "Modality_ex", "Duration_ex", "Intensity_ex"), 
     bg =customRed, part = "body")
table_aer_gender<- align(table_aer_gender, i= 1, part = "header", align = "center")
table_aer_gender <- set_header_labels(table_aer_gender, n= "N", Condition = "Pop", Modality_test = "Test_mod", Modality_ex = "Ex_mod", Duration_ex = "Ex_time (min)", Impact = "Impact", Rem = "Rem")
table_aer_gender <- set_table_properties(table_aer_gender, layout = "autofit", width = .5)


#ex_gender
ex_gender_var <- as_grouped_data(x =ex_gender_var, groups = c("Reference"))
table_ex_gender <- flextable :: as_flextable(ex_gender_var) %>%
  bold(j=1, i= ~ !is.na(Reference), bold =TRUE, part="body") %>%
  italic(j=1, i= ~ !is.na(Reference),part="body") %>%
  bold(part = "header", bold = TRUE) %>%
  colformat_double(i = ~ is.na(Reference), j = "n", digits = 0, big.mark = "") %>%
  autofit()

table_ex_gender <- add_header_row(table_ex_gender, values = "Gender : Other exercise & ppt ", colwidths = 9)
table_ex_gender <- table_ex_gender %>% 
  bold(part = "header", bold = TRUE) %>%
  bg(i = ~ Impact %in% "yes",
     j = c('Impact', 'Rem', "Site", "Modality_test", "Modality_ex", "Duration_ex", "Intensity_ex"), 
     bg =customGreen, part = "body") %>%
  bg(i = ~ Impact %in% "no",
     j = c('Impact', 'Rem', "Site", "Modality_test", "Modality_ex", "Duration_ex", "Intensity_ex"), 
     bg =customRed, part = "body")
table_ex_gender<- align(table_ex_gender, i= 1, part = "header", align = "center")
table_ex_gender <- set_header_labels(table_ex_gender, n= "N", Condition = "Pop", Modality_test = "Test_mod", Modality_ex = "Ex_mod", Duration_ex = "Ex_time (min)", Impact = "Impact", Rem = "Rem")
table_ex_gender <- set_table_properties(table_ex_gender, layout = "autofit", width = .5)
table_ex_gender

##fitness
#aer_ex

aer_fit_var <- as_grouped_data(x =aer_fit_var, groups = c("Reference"))
table_aer_fit <- flextable :: as_flextable(aer_fit_var) %>%
  bold(j=1, i= ~ !is.na(Reference), bold =TRUE, part="body") %>%
  italic(j=1, i= ~ !is.na(Reference),part="body") %>%
  bold(part = "header", bold = TRUE) %>%
  colformat_double(i = ~ is.na(Reference), j = "n", digits = 0, big.mark = "") %>%
  autofit()

table_aer_fit <- add_header_row(table_aer_fit, values = "Fitness : Aerobic exercise & ppt ", colwidths = 9)
table_aer_fit <- table_aer_fit %>% 
  bold(part = "header", bold = TRUE) %>%
  bg(i = ~ Impact %in% "yes",
     j = c('Impact', 'Rem', "Site", "Modality_test", "Modality_ex", "Duration_ex", "Intensity_ex"), 
     bg =customGreen, part = "body") %>%
  bg(i = ~ Impact %in% "no",
     j = c('Impact', 'Rem', "Site", "Modality_test", "Modality_ex", "Duration_ex", "Intensity_ex"), 
     bg =customRed, part = "body")
table_aer_fit<- align(table_aer_fit, i= 1, part = "header", align = "center")
table_aer_fit <- set_header_labels(table_aer_fit, n= "N", Condition = "Pop", Modality_test = "Test_mod", Modality_ex = "Ex_mod", Duration_ex = "Ex_time (min)", Impact = "Impact", Rem = "Rem")
table_aer_fit <- set_table_properties(table_aer_fit, layout = "autofit", width = .5)

#other_ex

ex_fit_var <- as_grouped_data(x =ex_fit_var, groups = c("Reference"))
table_ex_fit <- flextable :: as_flextable(ex_fit_var) %>%
  bold(j=1, i= ~ !is.na(Reference), bold =TRUE, part="body") %>%
  italic(j=1, i= ~ !is.na(Reference),part="body") %>%
  bold(part = "header", bold = TRUE) %>%
  colformat_double(i = ~ is.na(Reference), j = "n", digits = 0, big.mark = "") %>%
  autofit()

table_ex_fit <- add_header_row(table_ex_fit, values = "Fitness : Other exercise & ppt ", colwidths = 9)
table_ex_fit <- table_ex_fit %>% 
  bold(part = "header", bold = TRUE) %>%
  bg(i = ~ Impact %in% "yes",
     j = c('Impact', 'Rem', "Site", "Modality_test", "Modality_ex", "Duration_ex", "Intensity_ex"), 
     bg =customGreen, part = "body") %>%
  bg(i = ~ Impact %in% "no",
     j = c('Impact', 'Rem', "Site", "Modality_test", "Modality_ex", "Duration_ex", "Intensity_ex"), 
     bg =customRed, part = "body")
table_ex_fit<- align(table_ex_fit, i= 1, part = "header", align = "center")
table_ex_fit <- set_header_labels(table_ex_fit, n= "N", Condition = "Pop", Modality_test = "Test_mod", Modality_ex = "Ex_mod", Duration_ex = "Ex_time (min)", Impact = "Impact", Rem = "Rem")
table_ex_fit <- set_table_properties(table_ex_fit, layout = "autofit", width = .5)


##age

#aer_ex

aer_age_var <- as_grouped_data(x =aer_age_var, groups = c("Reference"))
table_aer_age <- flextable :: as_flextable(aer_age_var) %>%
  bold(j=1, i= ~ !is.na(Reference), bold =TRUE, part="body") %>%
  italic(j=1, i= ~ !is.na(Reference),part="body") %>%
  bold(part = "header", bold = TRUE) %>%
  colformat_double(i = ~ is.na(Reference), j = "n", digits = 0, big.mark = "") %>%
  autofit()

table_aer_age <- add_header_row(table_aer_age, values = "Age : Aerobic exercise & ppt ", colwidths = 9)
table_aer_age <- table_aer_age %>% 
  bold(part = "header", bold = TRUE) %>%
  bg(i = ~ Impact %in% "yes",
     j = c('Impact', 'Rem', "Site", "Modality_test", "Modality_ex", "Duration_ex", "Intensity_ex"), 
     bg =customGreen, part = "body") %>%
  bg(i = ~ Impact %in% "no",
     j = c('Impact', 'Rem', "Site", "Modality_test", "Modality_ex", "Duration_ex", "Intensity_ex"), 
     bg =customRed, part = "body")
table_aer_age<- align(table_aer_age, i= 1, part = "header", align = "center")
table_aer_age <- set_header_labels(table_aer_age, n= "N", Condition = "Pop", Modality_test = "Test_mod", Modality_ex = "Ex_mod", Duration_ex = "Ex_time (min)", Impact = "Impact", Rem = "Rem")
table_aer_age <- set_table_properties(table_aer_age, layout = "autofit", width = .5)


#other_ex

ex_age_var <- as_grouped_data(x =ex_age_var, groups = c("Reference"))
table_ex_age <- flextable :: as_flextable(ex_age_var) %>%
  bold(j=1, i= ~ !is.na(Reference), bold =TRUE, part="body") %>%
  italic(j=1, i= ~ !is.na(Reference),part="body") %>%
  bold(part = "header", bold = TRUE) %>%
  colformat_double(i = ~ is.na(Reference), j = "n", digits = 0, big.mark = "") %>%
  autofit()

table_ex_age <- add_header_row(table_ex_age, values = "Age : Other exercise & ppt ", colwidths = 9)
table_ex_age <- table_ex_age %>% 
  bold(part = "header", bold = TRUE) %>%
  bg(i = ~ Impact %in% "yes",
     j = c('Impact', 'Rem', "Site", "Modality_test", "Modality_ex", "Duration_ex", "Intensity_ex"), 
     bg =customGreen, part = "body") %>%
  bg(i = ~ Impact %in% "no",
     j = c('Impact', 'Rem', "Site", "Modality_test", "Modality_ex", "Duration_ex", "Intensity_ex"), 
     bg =customRed, part = "body")
table_ex_age<- align(table_ex_age, i= 1, part = "header", align = "center")
table_ex_age <- set_header_labels(table_ex_age, n= "N", Condition = "Pop", Modality_test = "Test_mod", Modality_ex = "Ex_mod", Duration_ex = "Ex_time (min)", Impact = "Impact", Rem = "Rem")
table_ex_age <- set_table_properties(table_ex_age, layout = "autofit", width = .5)


##psysoc

#aer_ex
aer_psysoc_var <- as_grouped_data(x =aer_psy_var, groups = c("Reference"))
table_aer_psysoc <- flextable :: as_flextable(aer_psysoc_var) %>%
  bold(j=1, i= ~ !is.na(Reference), bold =TRUE, part="body") %>%
  italic(j=1, i= ~ !is.na(Reference),part="body") %>%
  bold(part = "header", bold = TRUE) %>%
  colformat_double(i = ~ is.na(Reference), j = "n", digits = 0, big.mark = "") %>%
  autofit()

table_aer_psysoc <- add_header_row(table_aer_psysoc, values = "Psysoc : Aerobic exercise & ppt ", colwidths = 9)
table_aer_psysoc <- table_aer_psysoc %>% 
  bold(part = "header", bold = TRUE) %>%
  bg(i = ~ Impact %in% "yes",
     j = c('Impact', 'Rem', "Site", "Modality_test", "Modality_ex", "Duration_ex", "Intensity_ex"), 
     bg =customGreen, part = "body") %>%
  bg(i = ~ Impact %in% "no",
     j = c('Impact', 'Rem', "Site", "Modality_test", "Modality_ex", "Duration_ex", "Intensity_ex"), 
     bg =customRed, part = "body")
table_aer_psysoc<- align(table_aer_psysoc, i= 1, part = "header", align = "center")
table_aer_psysoc <- set_header_labels(table_aer_psysoc, n= "N", Condition = "Pop", Modality_test = "Test_mod", Modality_ex = "Ex_mod", Duration_ex = "Ex_time (min)", Impact = "Impact", Rem = "Rem")
table_aer_psysoc <- set_table_properties(table_aer_psysoc, layout = "autofit", width = .5)

#other_ex

ex_psysoc_var <- as_grouped_data(x =ex_psy_var, groups = c("Reference"))
table_ex_psysoc <- flextable :: as_flextable(ex_psysoc_var) %>%
  bold(j=1, i= ~ !is.na(Reference), bold =TRUE, part="body") %>%
  italic(j=1, i= ~ !is.na(Reference),part="body") %>%
  bold(part = "header", bold = TRUE) %>%
  colformat_double(i = ~ is.na(Reference), j = "n", digits = 0, big.mark = "") %>%
  autofit()

table_ex_psysoc <- add_header_row(table_ex_psysoc, values = "Psysoc : Other exercise & ppt ", colwidths = 9)
table_ex_psysoc <- table_ex_psysoc %>% 
  bold(part = "header", bold = TRUE) %>%
  bg(i = ~ Impact %in% "yes",
     j = c('Impact', 'Rem', "Site", "Modality_test", "Modality_ex", "Duration_ex", "Intensity_ex"), 
     bg =customGreen, part = "body") %>%
  bg(i = ~ Impact %in% "no",
     j = c('Impact', 'Rem', "Site", "Modality_test", "Modality_ex", "Duration_ex", "Intensity_ex"), 
     bg =customRed, part = "body")
table_ex_psysoc<- align(table_ex_psysoc, i= 1, part = "header", align = "center")
table_ex_psysoc <- set_header_labels(table_ex_psysoc, n= "N", Condition = "Pop", Modality_test = "Test_mod", Modality_ex = "Ex_mod", Duration_ex = "Ex_time (min)", Impact = "Impact", Rem = "Rem")
table_ex_psysoc <- set_table_properties(table_ex_psysoc, layout = "autofit", width = .5)


##menstruation

#aer_ex

table_aer_menstr_var <- as_grouped_data(x =aer_menstr_var, groups = c("Reference"))
table_aer_menstr <- flextable :: as_flextable(aer_menstr_var) %>%
  bold(j=1, i= ~ !is.na(Reference), bold =TRUE, part="body") %>%
  italic(j=1, i= ~ !is.na(Reference),part="body") %>%
  bold(part = "header", bold = TRUE) %>%
  colformat_double(i = ~ is.na(Reference), j = "n", digits = 0, big.mark = "") %>%
  autofit()

table_aer_menstr <- add_header_row(table_aer_menstr, values = "Menstruation cycle : Aerobic exercise & ppt ", colwidths = 9)
table_aer_menstr <- table_aer_menstr %>% 
  bold(part = "header", bold = TRUE) %>%
  bg(i = ~ Impact %in% "yes",
     j = c('Impact', 'Rem', "Site", "Modality_test", "Modality_ex", "Duration_ex", "Intensity_ex"), 
     bg =customGreen, part = "body") %>%
  bg(i = ~ Impact %in% "no",
     j = c('Impact', 'Rem', "Site", "Modality_test", "Modality_ex", "Duration_ex", "Intensity_ex"), 
     bg =customRed, part = "body")
table_aer_menstr<- align(table_aer_menstr, i= 1, part = "header", align = "center")
table_aer_menstr <- set_header_labels(table_aer_menstr, n= "N", Condition = "Pop", Modality_test = "Test_mod", Modality_ex = "Ex_mod", Duration_ex = "Ex_time (min)", Impact = "Impact", Rem = "Rem")
table_aer_menstr <- set_table_properties(table_aer_menstr, layout = "autofit", width = .5)


#other_ex
ex_menstr_var <- as_grouped_data(x =ex_menstr_var, groups = c("Reference"))
table_ex_menstr <- flextable :: as_flextable(ex_menstr_var) %>%
  bold(j=1, i= ~ !is.na(Reference), bold =TRUE, part="body") %>%
  italic(j=1, i= ~ !is.na(Reference),part="body") %>%
  bold(part = "header", bold = TRUE) %>%
  colformat_double(i = ~ is.na(Reference), j = "n", digits = 0, big.mark = "") %>%
  autofit()

table_ex_menstr <- add_header_row(table_ex_menstr, values = "Menstrual cycle : Other exercise & ppt ", colwidths = 9)
table_ex_menstr <- table_ex_menstr %>% 
  bold(part = "header", bold = TRUE) %>%
  bg(i = ~ Impact %in% "yes",
     j = c('Impact', 'Rem', "Site", "Modality_test", "Modality_ex", "Duration_ex", "Intensity_ex"), 
     bg =customGreen, part = "body") %>%
  bg(i = ~ Impact %in% "no",
     j = c('Impact', 'Rem', "Site", "Modality_test", "Modality_ex", "Duration_ex", "Intensity_ex"), 
     bg =customRed, part = "body")
table_ex_menstr<- align(table_ex_menstr, i= 1, part = "header", align = "center")
table_ex_menstr <- set_header_labels(table_ex_menstr, n= "N", Condition = "Pop", Modality_test = "Test_mod", Modality_ex = "Ex_mod", Duration_ex = "Ex_time (min)", Impact = "Impact", Rem = "Rem")
table_ex_menstr <- set_table_properties(table_ex_menstr, layout = "autofit", width = .5)



## Formattable

table_ex_gender_var <- formattable(ex_gender_var, 
                                   align = c("l", "c", "c", "c", "c", "c", "c", "c", "c", "c", "c", "c", "r"), 
                                   list(
                                     "Reference" = formatter("span", style = ~ style(color = "grey", font.weight = "bold")),
                                     "Impact" = formatter("span", style = x ~ style(color = ifelse(x=="yes", "green","red")))                                  ))
## other options

table_ex_gender <- table_ex_gender %>%
  color(i = ~ Impact %in% "yes",
        j = ~ Impact + Rem , color = customGreen) %>%
  color(i = ~ Impact %in% "no",
        j = ~ Impact + Rem, color = customRed)
## Power point

library(officer)
ppt <- read_pptx()
ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
ppt <- ph_with(ppt, value = table_aer_gender, location = ph_location_left()) 
ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
ppt <- ph_with(ppt, value = table_ex_gender, location = ph_location_left()) 
ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
ppt <- ph_with(ppt, value = table_aer_fit, location = ph_location_left()) 
ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
ppt <- ph_with(ppt, value = table_ex_fit, location = ph_location_left()) 
ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
ppt <- ph_with(ppt, value = table_aer_age, location = ph_location_left()) 
ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
ppt <- ph_with(ppt, value = table_ex_age, location = ph_location_left()) 
ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
ppt <- ph_with(ppt, value = table_aer_psysoc, location = ph_location_left()) 
ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
ppt <- ph_with(ppt, value = table_ex_psysoc, location = ph_location_left()) 
ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")
ppt <- ph_with(ppt, value = table_ex_menstr, location = ph_location_left()) 


print(ppt, target = "eih_var_r.pptx") 




## test


ex_psysoc_var <- as_grouped_data(x =ex_psy_var, groups = c("Reference"))
table_ex_psysoc <- flextable :: as_flextable(ex_psysoc_var) %>%
  bold(j=1, i= ~ !is.na(Reference), bold =TRUE, part="body") %>%
  italic(j=1, i= ~ !is.na(Reference),part="body") %>%
  bold(part = "header", bold = TRUE) %>%
  colformat_double(i = ~ is.na(Reference), j = "n", digits = 0, big.mark = "") %>%
  autofit()

table_ex_psysoc <- add_header_row(table_ex_psysoc, values = "Psysoc : Other exercise & ppt ", colwidths = 9)
table_ex_psysoc <- table_ex_psysoc %>% 
  bold(part = "header", bold = TRUE) %>%
  bg(i = ~ Impact %in% "yes",
     j = c('Impact', 'Rem', "Site", "Modality_test", "Modality_ex", "Duration_ex", "Intensity_ex"), 
     bg =customGreen, part = "body") %>%
  bg(i = ~ Impact %in% "no",
     j = c('Impact', 'Rem', "Site", "Modality_test", "Modality_ex", "Duration_ex", "Intensity_ex"), 
     bg =customRed, part = "body")
table_ex_psysoc<- align(table_ex_psysoc, i= 1, part = "header", align = "center")
table_ex_psysoc <- set_header_labels(table_ex_psysoc, n= "N", Condition = "Pop", Modality_test = "Test_mod", Modality_ex = "Ex_mod", Duration_ex = "Ex_time (min)", Impact = "Impact", Rem = "Rem")
table_ex_psysoc <- set_table_properties(table_ex_psysoc, layout = "autofit", width = .5)





