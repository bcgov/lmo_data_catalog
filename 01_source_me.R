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
current_year <- year(today())
fyfn <- current_year+5
tyfn <- current_year+10
#functions---------------------------

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
  inner_join(tbbl, income)
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
  janitor::remove_empty()

jo <- vroom(here("raw_data","job_openings.csv"), skip = 3)%>%
  janitor::remove_empty()

income <- read_excel(here("raw_data","Census 2021 Median Employment Income.xlsx"))|>
  select(NOC, `Median Income (Census 2021)`= contains("Median employment income"))|>
  mutate(`Median Income (Census 2021)`=as.numeric(`Median Income (Census 2021)`))

sd <- vroom(here("raw_data","supply_demand5.csv"), skip = 3)%>%
  janitor::remove_empty()

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

write.xlsx(tbbl1, here("out",
                       paste0("employment_by_industry_and_occupation_for_bc_lmo_",
                              current_year,
                              ".xlsx")))

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

write.xlsx(tbbl2, here("out",
                       paste0("employment_by_industry_for_bc_and_regions_lmo_",
                              current_year,
                              ".xlsx")))

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

write.xlsx(tbbl3, here("out",
                       paste0("job_openings_by_industry_and_occupation_for_bc_lmo_",
                              current_year,
                              ".xlsx")))

# hoo bc and regions by TEER------------------------------

hoo_sheets <- excel_sheets(here("raw_data",
                                        "LMO 2023E HOO BC and Regions 2023-08-23.xlsx"))
tibble(sheet=hoo_sheets[-length(hoo_sheets)])%>%
  mutate(data=map(sheet, load_sheet),
         data=map(data, join_income))%>%
  deframe()%>%
  write.xlsx(file = here("out",
                         paste0("hoo_bc_and_regions_by_TEER_lmo_",
                                current_year,
                                ".xlsx")))

#job-openings-expansion-replacement-by-occupation-for-bc-and-regions-lmo-YEARe.xlsx---------

tbbl4 <- jo%>%
  pivot_longer(cols=starts_with("2"), names_to = "year", values_to = "value")%>%
  clean_names()%>%
  filter(variable %in% c("Job Openings", "Expansion Demand", "Replacement Demand"),
         industry=="All industries",
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
  split(tbbl4$`Geographic Area`)|>
  write.xlsx(here("out",
                  paste0("job-openings-expansion-replacement-by-occupation-for-bc-and-regions-lmo-",
                         current_year,
                         ".xlsx")),
             overwrite = TRUE,
             asTable = TRUE
             )


#Employment_by_Ind_and_Occ_for_BC_and_Regions_xxxx.xlsx------------------

tbbl5 <- employment|>
  filter(!`Geographic Area` %in% c("North","South East"))|>
  pivot_longer(cols = starts_with("2"),
               names_to = "Date",
               values_to = "Value")
  write.xlsx(tbbl5,
             here("out",
                  paste0("Employment_by_Ind_and_Occ_for_BC_and_Regions_",
                         current_year,
                         ".xlsx")))

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

write.xlsx(tbbl6, here("out",
                       paste0("Employment_by_Occupation_for_BC_and_Regions_",
                              current_year,
                              ".xlsx")))

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

write.xlsx(tbbl7, here("out",
                       paste0("Job_Openings_by_Type_and_Occ_for_BC_and_Regions_",
                              current_year,
                              ".xlsx")))

#JO_by_Type,_Ind_and_Occ_for_BC_and_Regions_xxxx.xlsx---------------------------------

tbbl8 <- jo%>%
  pivot_longer(cols=starts_with("2"), names_to = "year", values_to = "value")%>%
  clean_names()%>%
  filter(variable %in% c("Job Openings","Expansion Demand","Replacement Demand"),
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

write.xlsx(tbbl8, here("out",
                       paste0("JO_by_Type,_Ind_and_Occ_for_BC_and_Regions_",
                              current_year,
                              ".xlsx")))
# LMO2023E-supply-composition-output-total10yr: from feng-----------------------
# LMO2023E-supply-composition-output-annual: from feng--------------------------
# Definitions: from report



