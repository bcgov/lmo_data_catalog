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
                          "LMO 2023E HOO BC and Regions 2023-08-23.xlsx"),
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

income <- read_excel(here("raw_data","median_income_jackie.xlsx"),
                     skip=6,
                     col_names = FALSE,
                     na = "x")
colnames(income) <- c("noc","everybody", "Median Income (Census 2021)")
income <- income%>%
  mutate(nocdf=map(noc, get_noc), .before = noc)%>%
  select(-noc, -everybody)%>%
  unnest(nocdf)

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

write.xlsx(tbbl1, here("out","employment_by_industry_and_occupation_for_bc_lmo_2023.xlsx"))

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

write.xlsx(tbbl2, here("out","employment_by_industry_for_bc_and_regions_lmo_2023.xlsx"))

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

write.xlsx(tbbl3, here("out","job_openings_by_industry_and_occupation_for_bc_lmo_2023.xlsx"))

# hoo bc and regions by TEER------------------------------

hoo_sheets <- excel_sheets(here("raw_data",
                                        "LMO 2023E HOO BC and Regions 2023-08-23.xlsx"))
hoo_data <- tibble(sheet=hoo_sheets[-length(hoo_sheets)])%>%
  mutate(data=map(sheet, load_sheet),
         data=map(data, join_income))%>%
  deframe()%>%
  write.xlsx(file = here("out","hoo_bc_and_regions_by_TEER_lmo_2023.xlsx"))

#job-openings-expansion-replacement-by-occupation-for-bc-and-regions-lmo-2022e.xlsx---------

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
                  "job-openings-expansion-replacement-by-occupation-for-bc-and-regions-lmo-2022e.xlsx"),
             overwrite = TRUE,
             asTable = TRUE
             )











