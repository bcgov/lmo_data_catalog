one <- "Data Dictionary"
two <- paste0("These data sets contain lists of occupations that are deemed to be high opportunity occupations (HOO) over the 10 year forecast period (",fyod,"-",tyfn,")")
three <- "Lists are provided for the 7 economic regions. Additionally, the lists
 provide estimates for Job Openings (for the 10 year forecast period), as well as
 the most recent Income data provided by Census 2021."
four <- "#NOC (2021): Denotes a 5-digit code according to the National Occupation Classification
 2021 system from Statistics Canada."
five <- "NOC (2021) Occupation Title:  Denotes the occupation title according to the National
 Occupation Classification 2021 system from Statisitics Canada."
six <- paste0("LMO Job Openings ",fyod,"-",tyfn,": The sum of expansion and replacement job openings. A job opening
 is the addition of a new job position through economic growth or a position that needs
 to be filled due  to someone exiting the labour force permanently.")
seven <- "TEER:  the type and/or amount of training, education, experience and
 responsibility typically required to work in an occupation. The NOC consists of
 six TEER categories, identified 0 through 5, which represent the second digit of the NOC code."
eight <- "Income: Median Income for full year full time employees."

data_dictionary <- list(`Data Dictionary`=tibble(" "=c(one, two, three, four, five, six, seven, eight)))

