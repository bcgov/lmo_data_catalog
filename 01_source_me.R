#' This script prepares 10 files for the BC data catalog: Requires inputs:
#'
#' "employment.csv" (4castviewer)
#' "job_openings.csv"  (4castviewer)
#' "Occupational Characteristics..." (Feng)
#' "clusters.csv" (https://rpubs.com/rpmartin/1058369)

#' To Run: Source me!

#libraries----------------------------
library(tidyverse)
library(here)
library(vroom)
library(janitor)
library(readxl)
library(openxlsx)
library(conflicted)
#constants------------------------------------
conflicts_prefer(dplyr::filter)
pct = createStyle(numFmt="0.0%")
fyod <- 2024 #first year of data
fyfn <- fyod+5
tyfn <- fyod+10
source("hoo_text.R") #data dictionary for hoo file
#functions---------------------------
last_three_columns <- function(tbbl){
  #' takes a tibble as an input, and returns an integer vector with the positions of the last 3 columns.
  #' called by function write_last3_percent
  (ncol(tbbl)-2): ncol(tbbl)
}
write_last3_percent <- function(lst_of_tbbls, file_name){
  #' takes as an input a list of tibbles and a file name, and returns nothing:
  #' side effect is to create excel file. 
  tbbl <- lst_of_tbbls|>
    enframe()|>
    mutate(rows=map_dbl(value, nrow),
           rows=rows+1, #need an extra row???
           rows=map(rows, seq),
           cols=map(value, last_three_columns)
    )
  wb = createWorkbook()
  sht = map(tbbl$name, addWorksheet, wb=wb)
  map2(sht, tbbl$value, writeData, wb=wb)
  pmap(list(sheet=sht, rows= tbbl$rows, cols= tbbl$cols), addStyle, wb=wb, style=pct, gridExpand=TRUE)
  saveWorkbook(wb, here("out", file_name), overwrite = TRUE)
}

cagrs <- function(tbbl){
  #' takes as an input a tbbl with 2 columns year and value and returns
  #' a tibble with CAGRs: NOT multiplied by 100: formatted as % when written to excel.
  current <- tbbl$value[tbbl$year==fyod]
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
  #'takes a tibble with columns year and value, and returns a tibble with sums 
  ffy_sum <- sum(tbbl$value[tbbl$year %in% (fyod+1):(fyod+5)])
  sfy_sum <- sum(tbbl$value[tbbl$year %in% (fyod+6):(fyod+10)])
  ty_sum <- sum(tbbl$value[tbbl$year %in% (fyod+1):(fyod+10)])
  tibble(`1st 5-year Sum`=ffy_sum,
         `2nd 5-year Sum`=sfy_sum,
         `10-year Sum`=ty_sum)
}

keep_only_hoo <- function(column, tbbl){
  tbbl |>
    filter(!grepl("Non", get(column)))|>
    select(NOC, Description, `2021 Census Median Employment Income (Employed)`)|>
    mutate(TEER=str_sub(NOC, 3, 3), .after="Description")
}

add_jo <- function(tbbl, region){
  jos_for_region <- regional_jo_by_occ|>
    filter(`Geographic Area`==region)
  left_join(tbbl, jos_for_region)|>
    relocate(contains("Job Openings"), .after="Description")|>
    select(-`Geographic Area`)
}

#read in data--------------------
employment <- vroom(here("raw_data","employment.csv"), skip = 3)%>%
  janitor::remove_empty()

jo <- vroom(here("raw_data","job_openings.csv"), skip = 3)%>%
  janitor::remove_empty()

occ_char <- readxl::read_excel(here("raw_data", 
                                    list.files(here("raw_data"), 
                                               pattern="Occupational Characteristics")), 
                               skip = 3, 
                               na = "x")

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
hoo_cols <- colnames(occ_char)[str_detect(colnames(occ_char),"Group: HOO")]
hoo_sheet_names <- str_remove_all(hoo_cols, "Occ Group: ")|>
  str_remove_all(" 2024E")
occ_char_hoo <- occ_char|>
  select(NOC, 
         Description, 
         all_of(hoo_cols), 
         `2021 Census Median Employment Income (Employed)`
         )

regional_jo_by_occ <- jo|>
  filter(Industry=="All industries",
         Variable=="Job Openings")|>
  pivot_longer(cols = starts_with("2"), names_to = "year", values_to = "value")|>
  group_by(NOC, `Geographic Area`)|>
  summarize("LMO Job Openings {fyod}-{tyfn}":=sum(value))
regions <- sort(unique(regional_jo_by_occ$`Geographic Area`))

hoo_tbbl <- tibble(hoo_cols=hoo_cols,
                   hoo_sheet_names=hoo_sheet_names,
                   occ_char_hoo=list(occ_char_hoo))|>
  mutate(hoo=map2(hoo_cols, occ_char_hoo, keep_only_hoo))|>
  select(hoo_sheet_names, hoo)|>
  arrange(hoo_sheet_names)|>
  mutate(full_region_name=regions,
         hoo=map2(hoo, full_region_name, add_jo)
         )|>
  select(-full_region_name)

c(data_dictionary, setNames(hoo_tbbl$hoo, hoo_tbbl$hoo_sheet_names))|>
  write.xlsx(file = here("out", "High Opportunity Occupations BC and Regions.xlsx"))

#JO_by_Type,_Ind_and_Occ_for_BC_and_Regions_xxxx.xlsx---------

tbbl5 <- jo%>%
  pivot_longer(cols=starts_with("2"), names_to = "year", values_to = "value")%>%
  clean_names()%>%
  filter(variable %in% c("Job Openings", "Expansion Demand", "Replacement Demand"),
       #  !geographic_area %in% c("North","South East")
         )%>%
  group_by(noc, description, industry, variable, geographic_area)%>%
  nest()%>%
  mutate(sums=map(data, sums))%>%
  unnest(data)%>%
  pivot_wider(names_from = year, values_from = value)%>%
  relocate(sums, .after=everything())%>%
  unnest(sums)

colnames(tbbl5) <- str_to_title(str_replace_all(colnames(tbbl5), "_", " "))
colnames(tbbl5)[1] <- "NOC"

tbbl5|>
  write.xlsx(here("out",
                  "JO by Type, Ind and Occ for BC and Regions.xlsx"),
             overwrite = TRUE,
             )

#Employment_by_Ind_and_Occ_for_BC_and_Regions_xxxx.xlsx------------------

tbbl6 <- employment|>
  filter(!`Geographic Area` %in% c("North", "South East"))|>
  pivot_longer(cols = starts_with("2"),
               names_to = "Date",
               values_to = "Value")
  write.xlsx(tbbl6,
             here("out",
                  "Employment by Ind and Occ for BC and Regions.xlsx"))

#Employment_by_Occupation_for_BC_and_Regions_xxxx.xlsx------------------

tbbl7 <- employment%>%
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
colnames(tbbl7) <- str_to_title(str_replace_all(colnames(tbbl7), "_", " "))
colnames(tbbl7)[1] <- "NOC"

tbbl7_data <- list("data"=tbbl7)
tbbl7_by_region <- tbbl7|>
  ungroup()|>
  select(-Industry, -Variable)|>
  split(tbbl7$`Geographic Area`)

tbbl7 <- c(tbbl7_data, tbbl7_by_region)
write_last3_percent(tbbl7, "Employment by Occupation for BC and Regions.xlsx")

#Job_Openings_by_Type_and_Occ_for_BC_and_Regions_xxxx.xlsx

tbbl8 <- jo%>%
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
colnames(tbbl8) <- str_to_title(str_replace_all(colnames(tbbl8), "_", " "))
colnames(tbbl8)[1] <- "NOC"

tbbl8_data <- list("data"=tbbl8)
tbbl8_by_region <- tbbl8|>
  ungroup()|>
  select(-Industry)|>
  split(tbbl8$`Geographic Area`)

tbbl8 <- c(tbbl8_data, tbbl8_by_region)

write.xlsx(tbbl8, here("out",
                       "Job Openings by Type and Occ for BC and Regions.xlsx"),
           asTable = TRUE)

#JO_by_Type,_Ind_and_Occ_for_BC_and_Regions_long---------------------------

tbbl9 <- jo%>%
  pivot_longer(cols=starts_with("2"), names_to = "year", values_to = "value")%>%
  filter(!`Geographic Area` %in% c("North","South East"))

write_csv(tbbl9, here("out",
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

tbbl10 <- jo |>
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

inner_join(tbbl10, clusters)|>
  select(NOC,
         Description,
         `Occ Group: Skills Cluster`=new_cluster,
         "LMO Job Openings {fyod}-{tyfn}":=jo
         )|>
  write.xlsx(here("out",
                  "Job Openings by NOC and Skill Cluster.xlsx"))
