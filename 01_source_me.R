#' Prepares the files for the BC data catalog

library(tidyverse)
library(conflicted)
library(here)
library(vroom)
library(janitor)
library(lubridate)
library(readxl)
library(openxlsx)
conflicts_prefer(dplyr::filter)
pct = createStyle(numFmt="0.0%")
current_year <- year(today())
fyfn <- current_year+5
tyfn <- current_year+10
source("hoo_text.R")
#functions---------------------------
last_three_columns <- function(tbbl){
  (ncol(tbbl)-2): ncol(tbbl)
}

write_last3_percent <- function(lst_of_tbbls, file_name){
  tbbl <- lst_of_tbbls|>
    enframe()|>
    mutate(rows=map_dbl(value, nrow),
           rows=rows+1,#need an extra row???
           rows=map(rows, seq),
           cols=map(value, last_three_columns)
    )
  wb = createWorkbook()
  sht = map(tbbl$name, addWorksheet, wb=wb)
  map2(sht, tbbl$value, writeData, wb=wb)
  pmap(list(sheet=sht, rows= tbbl$rows, cols= tbbl$cols), addStyle, wb=wb, style=pct, gridExpand=TRUE)
  saveWorkbook(wb, here("out", file_name), overwrite = TRUE)
}

fix_names <- function(tbbl){
  colnames(tbbl) <- str_replace_all(colnames(tbbl),"\r\n"," ")
  tbbl
}

cagrs <- function(tbbl){
  current <- tbbl$value[tbbl$year==current_year]
  fyfn <- tbbl$value[tbbl$year==fyfn]
  tyfn <- tbbl$value[tbbl$year==tyfn]
  ffy_cagr <- ((fyfn/current)^(.2)-1)
  sfy_cagr <- ((tyfn/fyfn)^(.2)-1)
  ty_cagr <- ((tyfn/current)^(.1)-1)
  tibble(`1st 5-year CAGR`=ffy_cagr,
         `2nd 5-year CAGR`=sfy_cagr,
         `10-year CAGR`=ty_cagr)
}

sums <- function(tbbl){
  ffy_sum <- sum(tbbl$value[tbbl$year %in% (current_year+1):(current_year+5)])
  sfy_sum <- sum(tbbl$value[tbbl$year %in% (current_year+6):(current_year+10)])
  ty_sum <- sum(tbbl$value[tbbl$year %in% (current_year+1):(current_year+10)])

  tibble(`1st 5-year Sum`=ffy_sum,
         `2nd 5-year Sum`=sfy_sum,
         `10-year Sum`=ty_sum)
}

load_sheet <- function(sht){
  temp <- read_excel(here("raw_data",
                          list.files(here("raw_data"), pattern = "HOO BC and Regions")),
                     sheet=sht,
                     skip=4)
  temp[,1:3]%>%
    mutate(TEER=str_sub(`#NOC (2021)`,3,3))
}

join_income <- function(tbbl){
  inner_join(tbbl, income, by=c("#NOC (2021)"="NOC"))
}

get_noc <- function(str){
  temp <- as.data.frame(str_split_fixed(str, " ", 2))
  colnames(temp) <- c("#NOC (2021)","description")
  temp <- temp%>%
    mutate(`#NOC (2021)`=paste0("#", `#NOC (2021)`))%>%
    select(-description)
}

#read in data--------------------

employment <- vroom(here("raw_data","employment.csv"), skip = 3)%>%
  janitor::remove_empty() |>
  mutate(Description=if_else(Description=="Seniors managers - public and private sector",
                             "Senior managers - public and private sector",
                             Description))

jo <- vroom(here("raw_data","job_openings.csv"), skip = 3)%>%
  janitor::remove_empty()|>
  mutate(Description=if_else(Description=="Seniors managers - public and private sector",
                             "Senior managers - public and private sector",
                             Description))

income <- read_excel(here("raw_data","Census 2021 Median Employment Income.xlsx"))|>
  select(NOC, `Median Income (Census 2021)`= contains("Median employment income"))|>
  mutate(`Median Income (Census 2021)`=as.numeric(`Median Income (Census 2021)`))

sd <- vroom(here("raw_data","supply_demand5.csv"), skip = 3)%>%
  janitor::remove_empty()|>
  mutate(Description=if_else(Description=="Seniors managers - public and private sector",
                             "Senior managers - public and private sector",
                             Description))

#employment by industry and occupation for bc-------------------------------

tbbl1 <- employment%>%
  pivot_longer(cols=starts_with("2"), names_to = "year", values_to = "value")%>%
  clean_names()%>%
  filter(geographic_area=="British Columbia")%>%
  group_by(noc, description, industry, variable, geographic_area)%>%
  nest()%>%
  mutate(cagrs=map(data, cagrs))%>%
  unnest(data)%>%
  pivot_wider(names_from = year, values_from = value)%>%
  relocate(cagrs, .after=everything())%>%
  unnest(cagrs)
colnames(tbbl1) <- str_to_title(str_replace_all(colnames(tbbl1), "_", " "))
colnames(tbbl1)[1] <- "NOC"

tbbl1 <- list("data" = tbbl1)
write_last3_percent(tbbl1, "Employment by Industry and Occupation for BC.xlsx")

#employment by industry for bc and regions-----------------------------------

tbbl2 <- employment%>%
  pivot_longer(cols=starts_with("2"), names_to = "year", values_to = "value")%>%
  clean_names()%>%
  filter(noc=="#T",
         !geographic_area %in% c("North","South East"))%>%
  group_by(noc, description, industry, variable, geographic_area)%>%
  nest()%>%
  mutate(cagrs=map(data, cagrs))%>%
  unnest(data)%>%
  pivot_wider(names_from = year, values_from = value)%>%
  relocate(cagrs, .after=everything())%>%
  unnest(cagrs)
colnames(tbbl2) <- str_to_title(str_replace_all(colnames(tbbl2), "_", " "))
colnames(tbbl2)[1] <- "NOC"

tbbl2_data <- list("data" = tbbl2)
tbbl2_by_region <- tbbl2|>
  ungroup()|>
  select(-NOC, -Description, -Variable)|>
  split(tbbl2$`Geographic Area`)


tbbl2 <- c(tbbl2_data, tbbl2_by_region)
write_last3_percent(tbbl2, "Employment by Industry for BC and Regions.xlsx")

#job openings by industry and occupation for bc------------------------

tbbl3 <- jo%>%
  pivot_longer(cols=starts_with("2"), names_to = "year", values_to = "value")%>%
  clean_names()%>%
  filter(geographic_area=="British Columbia",
         variable=="Job Openings")%>%
  group_by(noc, description, industry, variable, geographic_area)%>%
  nest()%>%
  mutate(sums=map(data, sums))%>%
  unnest(data)%>%
  pivot_wider(names_from = year, values_from = value)%>%
  relocate(sums, .after=everything())%>%
  unnest(sums)
colnames(tbbl3) <- str_to_title(str_replace_all(colnames(tbbl3), "_", " "))
colnames(tbbl3)[1] <- "NOC"

write.xlsx(tbbl3, here("out", "Job Openings by Industry and Occupation for BC.xlsx"))

# High_Opportunity_Occupations_BC_and_regions------------------------------
hoo_sheets <- excel_sheets(here("raw_data",
                                "LMO 2023E HOO BC and Regions 2023-08-23.xlsx"))
hoo <- tibble(sheet=hoo_sheets[-length(hoo_sheets)])%>%
  mutate(data=map(sheet, load_sheet),
         data=map(data, join_income),
         data=map(data, fix_names))%>% #weirdness with job openings column name
  deframe()
hoo <- c(data_dictionary, hoo)

write.xlsx(hoo, file = here("out",
                         "High Opportunity Occupations BC and Regions.xlsx"))

#JO_by_Type,_Ind_and_Occ_for_BC_and_Regions_xxxx.xlsx---------

tbbl4 <- jo%>%
  pivot_longer(cols=starts_with("2"), names_to = "year", values_to = "value")%>%
  clean_names()%>%
  filter(variable %in% c("Job Openings", "Expansion Demand", "Replacement Demand"),
         !geographic_area %in% c("North","South East")
         )%>%
  group_by(noc, description, industry, variable, geographic_area)%>%
  nest()%>%
  mutate(sums=map(data, sums))%>%
  unnest(data)%>%
  pivot_wider(names_from = year, values_from = value)%>%
  relocate(sums, .after=everything())%>%
  unnest(sums)

colnames(tbbl4) <- str_to_title(str_replace_all(colnames(tbbl4), "_", " "))
colnames(tbbl4)[1] <- "NOC"

tbbl4|>
  write.xlsx(here("out",
                  "JO by Type, Ind and Occ for BC and Regions.xlsx"),
             overwrite = TRUE,
             )

#Employment_by_Ind_and_Occ_for_BC_and_Regions_xxxx.xlsx------------------

tbbl5 <- employment|>
  filter(!`Geographic Area` %in% c("North","South East"))|>
  pivot_longer(cols = starts_with("2"),
               names_to = "Date",
               values_to = "Value")
  write.xlsx(tbbl5,
             here("out",
                  "Employment by Ind and Occ for BC and Regions.xlsx"))

#Employment_by_Occupation_for_BC_and_Regions_xxxx.xlsx------------------

tbbl6 <- employment%>%
  pivot_longer(cols=starts_with("2"), names_to = "year", values_to = "value")%>%
  clean_names()%>%
  filter(industry=="All industries",
         !geographic_area %in% c("North","South East"))%>%
  group_by(noc, description, industry, variable, geographic_area)%>%
  nest()%>%
  mutate(cagrs=map(data, cagrs))%>%
  unnest(data)%>%
  pivot_wider(names_from = year, values_from = value)%>%
  relocate(cagrs, .after=everything())%>%
  unnest(cagrs)
colnames(tbbl6) <- str_to_title(str_replace_all(colnames(tbbl6), "_", " "))
colnames(tbbl6)[1] <- "NOC"

tbbl6_data <- list("data"=tbbl6)
tbbl6_by_region <- tbbl6|>
  ungroup()|>
  select(-Industry, -Variable)|>
  split(tbbl6$`Geographic Area`)

tbbl6 <- c(tbbl6_data, tbbl6_by_region)
write_last3_percent(tbbl6, "Employment by Occupation for BC and Regions.xlsx")

#Job_Openings_by_Type_and_Occ_for_BC_and_Regions_xxxx.xlsx

tbbl7 <- jo%>%
  pivot_longer(cols=starts_with("2"), names_to = "year", values_to = "value")%>%
  clean_names()%>%
  filter(industry=="All industries",
         variable %in% c("Job Openings","Expansion Demand","Replacement Demand"),
         !geographic_area %in% c("North","South East"))%>%
  group_by(noc, description, industry, variable, geographic_area)%>%
  nest()%>%
  mutate(sums=map(data, sums))%>%
  unnest(data)%>%
  pivot_wider(names_from = year, values_from = value)%>%
  relocate(sums, .after=everything())%>%
  unnest(sums)
colnames(tbbl7) <- str_to_title(str_replace_all(colnames(tbbl7), "_", " "))
colnames(tbbl7)[1] <- "NOC"

tbbl7_data <- list("data"=tbbl7)
tbbl7_by_region <- tbbl7|>
  ungroup()|>
  select(-Industry)|>
  split(tbbl7$`Geographic Area`)

tbbl7 <- c(tbbl7_data, tbbl7_by_region)

write.xlsx(tbbl7, here("out",
                       "Job Openings by Type and Occ for BC and Regions.xlsx"),
           asTable = TRUE)

#JO_by_Type,_Ind_and_Occ_for_BC_and_Regions_long---------------------------

tbbl8 <- jo%>%
  pivot_longer(cols=starts_with("2"), names_to = "year", values_to = "value")%>%
  filter(!`Geographic Area` %in% c("North","South East"))

write_csv(tbbl8, here("out",
                      "JO by Type, Ind and Occ for BC and Regions (long).csv"))
#huge file so zip it---------------------
zip(zipfile=here("out",
                 "JO by Type, Ind and Occ for BC and Regions (long)"),
    files=here("out","JO by Type, Ind and Occ for BC and Regions (long).csv"))
#get rid of csv file
file.remove(here("out",
                  "JO by Type, Ind and Occ for BC and Regions (long).csv"))
# Supply Composition for BC (10-Year Total) from Feng-------------
# Supply Composition for BC (Annual) from Feng--------------------
# Definitions: from report----------------

#Job Openings by Skill Cluster-------------------------------

tbbl9 <- jo |>
  filter(Industry=="All industries",
         `Geographic Area`=="British Columbia",
         Variable=="Job Openings")|>
  select(-Industry, -`Geographic Area`, -Variable)|>
  pivot_longer(cols=starts_with("2"), names_to = "year", values_to = "jo")|>
  group_by(NOC, Description)|>
  summarize(jo=sum(jo))

clusters <- read_csv(here("raw_data","clusters.csv"))|>
  select(NOC, new_cluster)|>
  separate(NOC, into=c("NOC", "Description"), sep=": ")|>
  mutate(NOC=paste0("#", NOC))

inner_join(tbbl9, clusters)|>
  select(NOC,
         Description,
         `Occ Group: Skills Cluster`=new_cluster,
         "LMO Job Openings {current_year}-{tyfn}":=jo
         )|>
  write.xlsx(here("out",
                  "Job Openings by NOC and Skill Cluster.xlsx"))
