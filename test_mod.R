library(pdftools)
library(tidyverse)
library(tabulizer)
library(shiny)
library(miniUI)
library(googlesheets4)
library(tidyr)
library(officer)
library(data.table)
library(rJava)
library(xlsx)
library(readxl)
library(data.table)
library(dplyr)
library(tidyverse)
library(skimr)
library(ggplot2)
library(RColorBrewer)
library(ggthemes)
library(patchwork)
sem_to_sd <- function(sem, n){
  sd <- sem * sqrt(n)
  round(sd,2)
}    

test_mod <- read_sheet("https://docs.google.com/spreadsheets/d/1uR60vInUAPX4LhcyF60HPfFelOHdvPMuEJpx34SVaKA/edit#gid=0")

yes
0

str(test_mod)
test_mod %>% count(test_site) %>% view()
test_mod %>% select(reference, test_mod, ex_type, test_site) %>% filter(ex_type == "cycling_arm")  %>% view()
test_mod %>% select(reference, n, age_mean, stimu_type, test_mod) %>% filter(reference== "Vaegter_2017") %>% view()

#local vs remote

loc <- test_mod %>% mutate(local2 =if_else((ex_type == "back_ext" & test_site %in% c("back","biceps_fem", "sacrum")) |
                                                  (ex_type == "elbow_flex/ext" & test_site %in% c("biceps", "elbow", "forearm", "upper_limb")) |
                                                  (ex_type == "core_training" & test_site %in% c( "abdomen", "back", "glut_med", "sacrum")) |
                                                  (ex_type == "elbow_flex" & test_site %in% c("biceps", "elbow", "forearm", "upper_limb")) |
                                                  (ex_type == "full_body_circuit" & test_site %in% c("abdomen", "back", "biceps", "biceps_fem") )|
                                                  (ex_type == "full_body_stretching" & test_site %in% c("back", "glut_med")) |
                                                  (ex_type == "grip" & test_site %in% c("forearm", "finger_dominant", "hand", "brachioradialis", "upper_limb")) |
                                                  (ex_type == "cycling" & test_site %in% c("quadriceps", "biceps_fem", "calf", "knee", "lower_limb",
                                                                                           "patellar_tendon_sympt", "patellar_tendon_asympt", "tib_ant")) |
                                                  ( ex_type == "knee_ext" & test_site == "quadriceps") |
                                                  (ex_type %in% c("leg_press", "leg_press_bfr40", "leg_press_bfr80", "squat", "running/squat") & test_site == "quadriceps") |
                                                  (ex_type == "lifting" & test_site == "back") |
                                                  (ex_type =="lower_limb_circuit" & test_site == "lower_limb" ) |
                                                  (ex_type =="upper_limb_circuit" & test_site == "upper_limb" ) |
                                                  (ex_type == "shoulder_abd" & test_site == "deltoid") |
                                                  (ex_type == "tooth_clenching" & test_site == "masseter") |
                                                  (ex_type == "trunk_flexion" & test_site == "abdomen") |
                                                  (test_site %in% c("quadriceps", "biceps_fem", "calf", "knee", "lower_limb",
                                                   "patellar_tendon_sympt", "patellar_tendon_asympt", "tib_ant") & 
                                                   ex_type %in% c("cycling", "walking", "runinng", "cycling_interval", "stepper")) | 
                                                   (test_site == "trapezius" & ex_type == "cycling_arm"),"yes", "no"))

loc %>% relocate(local2, local) %>% head(20) %>% view()

# number of reference
count(test_mod, reference)%>% select(reference) %>% pull() %>% length()

#view ref
count(test_mod, reference)%>% select(reference)  %>% arrange() %>% view()

## number of test_mod
test_mod %>% count(test_mod) %>% view()
test_mod %>% select(reference, test_mod, condition) %>% filter(test_mod == "cptt") %>% view()

loc %>% relocate(local2) %>% select(local2, reference, ex_mode, ex_type, test_mod, test_site) %>%filter(local2 == "yes", ex_mode == "aerobic") %>%view()

#visualisation test_mod with eih

## local vs remote effect aerobic
ploc1_aer <- loc %>% filter(local2 == "yes", ex_mode == "aerobic") %>% count(test_mod,eih) %>% 
  mutate(proportion = 100*(n/sum(n)))  %>%
  filter(proportion >10) %>% mutate(proportion = round(proportion,2))%>% 
  ggplot(aes(test_mod,proportion, fill=eih, label=proportion)) +
  geom_bar(stat="identity")+
  scale_fill_brewer(palette="Pastel1") +
  scale_y_continuous(limits = c(0, 100))+
  geom_text(position = position_stack(vjust = 0.5), size = 3, colour="white")+
  labs(
    subtitle = "Over 10%",
    x = "", y="Proportion (%)",
    fill = "EIH effect :") +
  theme_stata() +
  theme(axis.text.x = element_text(angle = 60, hjust = 1))
  
ploc2_aer <- loc %>% filter(local2 == "yes", ex_mode == "aerobic") %>% count(test_mod,eih) %>% 
  mutate(proportion = 100*(n/sum(n)))  %>% 
  filter(proportion <=10 ) %>% mutate(proportion = round(proportion,2)) %>%
  ggplot(aes(fct_reorder(test_mod,proportion), proportion, fill=eih, label=proportion)) +
  geom_bar(stat="identity")+
  scale_fill_brewer(palette="Pastel1") +
  scale_y_continuous(limits = c(0, 10))+
  geom_text(position = position_stack(vjust = 0.5), size = 3, colour="white")+
  labs(
    subtitle = "Under 10% ",
    x = "Test modalities", y="Proportion (%)",
    fill = "EIH effect :") +
  theme_stata() +
  theme(axis.text.x = element_text(angle = 60, hjust = 1)) 

(ploc1_aer/ploc2_aer) + plot_layout(nrow=2, guides="collect") + plot_annotation(title = "Test modalities across studies (local effects, aerobic exercise) & EIH impact")


prem1_aer <- loc %>% filter(local2 == "no", ex_mode == "aerobic") %>% count(test_mod,eih) %>% 
  mutate(proportion = 100*(n/sum(n)))  %>%
  filter(proportion >10) %>% mutate(proportion = round(proportion,2))%>%
  ggplot(aes(test_mod,proportion, fill=eih, label=proportion)) +
  geom_bar(stat="identity")+
  scale_fill_brewer(palette="Pastel1") +
  scale_y_continuous(limits = c(0, 100))+
  geom_text(position = position_stack(vjust = 0.5), size = 3, colour="white")+
  labs(
    subtitle = "Over 10%",
    x = "", y="Proportion (%)",
    fill = "EIH effect :") +
  theme_stata() +
  theme(axis.text.x = element_text(angle = 60, hjust = 1))

prem2_aer <-loc %>% filter(local2 == "no", ex_mode == "aerobic") %>% count(test_mod,eih) %>%
  mutate(proportion = 100*(n/sum(n)))  %>% 
  filter(proportion <=10 ) %>% mutate(proportion = round(proportion,2)) %>%
  ggplot(aes(fct_reorder(test_mod,proportion), proportion, fill=eih, label=proportion)) +
  geom_bar(stat="identity")+
  scale_fill_brewer(palette="Pastel1") +
  scale_y_continuous(limits = c(0, 10))+
  geom_text(position = position_stack(vjust = 0.5), size = 3, colour="white")+
  labs(
    subtitle = "Under 10% ",
    x = "Test modalities", y="Proportion (%)",
    fill = "EIH effect :") +
  theme_stata() +
  theme(axis.text.x = element_text(angle = 60, hjust = 1)) 

(prem1_aer/prem2_aer) + plot_layout(nrow=2, guides="collect") + plot_annotation(title = "Test modalities across studies (remote effects, aerobic exercise) & EIH impact")
## Loca vs remote non aerobic ex
ploc1_other <- loc %>% filter(local2 == "yes", ex_mode != "aerobic") %>% count(test_mod,eih) %>% 
  mutate(proportion = 100*(n/sum(n)))  %>%
  filter(proportion >10) %>% mutate(proportion = round(proportion,2))%>% 
  ggplot(aes(test_mod,proportion, fill=eih, label=proportion)) +
  geom_bar(stat="identity")+
  scale_fill_brewer(palette="Pastel1") +
  scale_y_continuous(limits = c(0, 100))+
  geom_text(position = position_stack(vjust = 0.5), size = 3, colour="white")+
  labs(
    subtitle = "Over 10%",
    x = "", y="Proportion (%)",
    fill = "EIH effect :") +
  theme_stata() +
  theme(axis.text.x = element_text(angle = 60, hjust = 1))

ploc2_other <- loc %>% filter(local2 == "yes", ex_mode != "aerobic") %>% count(test_mod,eih) %>% 
  mutate(proportion = 100*(n/sum(n)))  %>% 
  filter(proportion <=10 ) %>% mutate(proportion = round(proportion,2)) %>%
  ggplot(aes(fct_reorder(test_mod,proportion), proportion, fill=eih, label=proportion)) +
  geom_bar(stat="identity")+
  scale_fill_brewer(palette="Pastel1") +
  scale_y_continuous(limits = c(0, 10))+
  geom_text(position = position_stack(vjust = 0.5), size = 3, colour="white")+
  labs(
    subtitle = "Under 10% ",
    x = "Test modalities", y="Proportion (%)",
    fill = "EIH effect :") +
  theme_stata() +
  theme(axis.text.x = element_text(angle = 60, hjust = 1)) 

(ploc1_other/ploc2_other) + plot_layout(nrow=2, guides="collect") + plot_annotation(title = "Test modalities across studies (local effects, other exercise) & EIH impact")


prem1_other <- loc %>% filter(local2 == "no", ex_mode != "aerobic") %>% count(test_mod,eih) %>% 
  mutate(proportion = 100*(n/sum(n)))  %>%
  filter(proportion >10) %>% mutate(proportion = round(proportion,2))%>%
  ggplot(aes(test_mod,proportion, fill=eih, label=proportion)) +
  geom_bar(stat="identity")+
  scale_fill_brewer(palette="Pastel1") +
  scale_y_continuous(limits = c(0, 100))+
  geom_text(position = position_stack(vjust = 0.5), size = 3, colour="white")+
  labs(
    subtitle = "Over 10%",
    x = "", y="Proportion (%)",
    fill = "EIH effect :") +
  theme_stata() +
  theme(axis.text.x = element_text(angle = 60, hjust = 1))

prem2_other <-loc %>% filter(local2 == "no", ex_mode != "aerobic") %>% count(test_mod,eih) %>%
  mutate(proportion = 100*(n/sum(n)))  %>% 
  filter(proportion <=10 ) %>% mutate(proportion = round(proportion,2)) %>%
  ggplot(aes(fct_reorder(test_mod,proportion), proportion, fill=eih, label=proportion)) +
  geom_bar(stat="identity")+
  scale_fill_brewer(palette="Pastel1") +
  scale_y_continuous(limits = c(0, 10))+
  geom_text(position = position_stack(vjust = 0.5), size = 3, colour="white")+
  labs(
    subtitle = "Under 10% ",
    x = "Test modalities", y="Proportion (%)",
    fill = "EIH effect :") +
  theme_stata() +
  theme(axis.text.x = element_text(angle = 60, hjust = 1)) 

(prem1_other/prem2_other) + plot_layout(nrow=2, guides="collect") + plot_annotation(title = "Test modalities across studies (remote effects, other exercise) & EIH impact")
# visualisation test_mod with ex_type
# visulation test_mod with stim type
#visualisation test_mod

p1 <- test_mod %>% count(test_mod) %>% mutate(proportion = 100*(n/sum(n))) %>% select(test_mod,proportion) %>% full_join(test_mod) %>%
  filter(proportion >10) %>%
  ggplot(aes(test_mod,proportion, fill=test_mod)) +
  geom_bar(stat="identity")+
  scale_fill_brewer(palette="Pastel1") +
  scale_y_continuous(limits = c(0, 100))+
  scale_x_discrete(labels = NULL) +
  labs(
    subtitle = "Over 10%",
    x = "", y="Proportion (%)",
    fill = "Test modalities") +
  theme_stata()
p2 <- test_mod %>% count(test_mod) %>% mutate(proportion = 100*(n/sum(n)))  %>% 
  filter(proportion <=10 & proportion > 1.5) %>%
  ggplot(aes(test_mod,proportion, fill=test_mod)) +
  geom_bar(stat="identity")+
  scale_fill_brewer(palette="Paired") +
  scale_y_continuous(limits = c(0, 10))+
  scale_x_discrete(labels = NULL) +
  labs(
    subtitle = "Between 1.5 & 10% ",
    x = "", y="Proportion (%)",
    fill = "Test modalities") +
  theme_stata()


p4 <- test_mod %>% count(test_mod) %>% mutate(proportion = 100*(n/sum(n))) %>%
  filter(proportion <= 1.5 & proportion > 0.5) %>%
  ggplot(aes(test_mod_h,proportion, fill=test_mod)) +
  geom_bar(stat="identity")+
  scale_fill_brewer(palette="Set3") +
  scale_y_continuous(limits = c(0, 5))+
  labs(
    subtitle = "Between 0.5 & 1.5% ",
    x = "", y="Proportion (%)",
    fill = "Test modalities") +
  theme_stata()

p5 <- test_mod %>% count(test_mod) %>% mutate(proportion = 100*(n/sum(n))) %>%
  filter(proportion <= .5) %>%
  ggplot(aes(test_mod_h,proportion, fill=test_mod)) +
  geom_bar(stat="identity")+
  scale_fill_brewer(palette="Set3") +
  scale_y_continuous(limits = c(0, 5))+
  scale_x_discrete(labels = NULL) +
  labs(
    subtitle = "Under 0.5% ",
    x = "", y="Proportion (%)",
    fill = "Test modalities") +
  theme_stata()
p2
plot_test_mod <- (p1|p2)/(p4|p5)
plot_test_mod
plot_test_mod + plot_annotation(title = "Test modalities across studies")
#tryning to represent test_site as ell
a <- test_mod %>% select(test_mod) %>% count(test_mod) %>% mutate(proportion = 100*(n/sum(n)))%>% filter(proportion > 3) 
b <- test_mod %>% select(test_mod, test_site)
c <- left_join(a,b)
c %>%
  ggplot(aes(test_mod,proportion, fill=test_site)) +
  geom_bar(stat="identity") 

  scale_fill_brewer(palette="Spectral") +
  theme_stata()

test_mod %>% filter(reference == "Bartholomew_1996") %>% select(ex_type, ex_mode, test_mod, test_site, eih) %>% view()

display.brewer.all()


#eih & clinical pop
test_mod %>% filter(condition != "healthy") %>% select(reference, condition,test_mod, test_site, ex_type,eih, ) %>% arrange(eih) %>% view()








n <- 36
sem <- c(23.4, 26.7, 25.2, 24.8)
                                                                                                              
sem_to_sd(sem,n)

m <- c(20.75, )
s <- c(3.8, 2.9,4.7, 4.6)
mean(m)
mean(s)







