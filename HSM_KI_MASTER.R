library(dplyr)
library(readxl)
library(openxlsx)
source("utils/create_formatted_excel.R")

# Load the data
# data <- read_excel("C:/Users/test/Downloads/ET_2025HSM_v2024-12-18_sample (2).xlsx", sheet = "sample")
# data <- read_excel("C:/Users/test/Downloads/ET_2025HSM_v2024-12-18_sample (1).xlsx", sheet = "sample")

data <- read_excel("C:/Users/test/Downloads/ET_2025HSM_v2024-12-18_sample_down_12_26_2024.xlsx", sheet = "sample")
View(data)

data <- data %>%
  mutate(kiib = ifelse(`number_of_ki_in_KI-BANK` < 5, 5, `number_of_ki_in_KI-BANK`))

dataaa <- data[rep(seq_len(nrow(data)), data$kiib), ]
View(dataaa)

dataaa$Ki_Identification <- ave(dataaa$kiib, dataaa$admin4Pcod, FUN = function(x) seq_along(x))
View(dataaa)

# create new column admin4Pcod_Ki_Identification by combining admin4Pcod and Ki_Identification separated by underscore
dataaa$admin4Pcod_Ki_Identification <- paste(dataaa$admin4Pcod, dataaa$Ki_Identification, sep = "_ki")
View(dataaa)

ki_data <- read_excel("C:/Users/test/Downloads/KI_BANK_2024-12-25.xlsx", sheet = "Clean_ki_Bank")

# select rows from ki_data where admin4Pcod is not NA
ki_data <- ki_data[!is.na(ki_data$admin4Pcod),] 
View(ki_data)

ki_dataa <- ki_data %>%
  group_by(admin4Pcod) %>%
  mutate(admin4Pcod_Ki_Identification = paste0(admin4Pcod, "_ki", row_number())) %>%
  ungroup() %>% 
  select(admin4Pcod_Ki_Identification,phone_num)

View(ki_dataa)

# left join ki_dataa and dataaa based on admin4Pcod_Ki_Identification
joined_data <- dataaa  %>%
  left_join(ki_dataa, by = "admin4Pcod_Ki_Identification")

View(joined_data)



###########################################HSM 2nd round KI sample#############################################
# filter the data where 2nd_round_HSM_sampling is sample, backup or Inaccessible
r2_data <- joined_data[joined_data$`2nd_round_HSM_sampling` %in% c("sample", "backup", "Inaccessible"),]

View(r2_data)
# select and order columns "admin1name", "admin1Pcod","admin2name", "admin2Pcod","admin3name","admin3Pcod","admin4name","admin4Pcod","admin4name_local","2nd_round_HSM_sampling","priority_for_Ki_Identification","admin4Pcod_Ki_Identification", "phone_num" 
r2_data <- r2_data[,c("admin1name", "admin1Pcod","admin2name", "admin2Pcod","admin3name","admin3Pcod","admin4name","admin4Pcod","admin4name_local","2nd_round_HSM_sampling","priority_for_Ki_Identification","number_of_ki_identification_needed","admin4Pcod_Ki_Identification", "phone_num")]
# add "Kebele_" to the admin4name column where admin4name is number only
r2_data$admin4name <- ifelse(grepl("^[0-9]+$", r2_data$admin4name), paste("Kebele ", r2_data$admin4name, sep = ""), r2_data$admin4name)


View(r2_data)



###########################################KI not in 2nd round#####################################################

# filter the data where from joined_data where not in r2_data
ki_not_in_r2 <- joined_data[!joined_data$admin4Pcod_Ki_Identification %in% r2_data$admin4Pcod_Ki_Identification,]

View(ki_not_in_r2)




# rename column names
data_rearranged <- r2_data %>%
  rename(
    Region = admin1name,
    Zone = admin2name,
    Woreda = admin3name,
    Kebele = admin4name,
    kebele_local=admin4name_local,
    Sampling = `2nd_round_HSM_sampling`,
    Priority = "priority_for_Ki_Identification",
    `KI to collect`="number_of_ki_identification_needed",
    `KI Identification` = admin4Pcod_Ki_Identification,
    `KI Phone number` = phone_num
  ) %>%
  select(Region, admin1Pcod, Zone, admin2Pcod, Woreda, 
         admin3Pcod, Kebele, admin4Pcod, Sampling, 
         Priority,`KI to collect`,`KI Identification`,`KI Phone number`)

View(data_rearranged)

# add new columns KI community role,	KI Phone number,	Comment, Partner organization Name and	Focal person E-mail Address from Partner organization
data_rearranged$`KI community role` <- NA
# data_rearranged$`KI Phone number` <- NA
data_rearranged$Comment <- NA
data_rearranged$`Partner organization Name` <- NA
data_rearranged$`Focal person E-mail Address from Partner organization` <- NA

# View(data_rearranged)
# 
# # make value in Region,Zone, Woreda and Kebele NA where  different from KI 01
# data_rearranged$Region[data_rearranged$`KI Order` != "KI 01"] <- NA
# data_rearranged$Zone[data_rearranged$`KI Order` != "KI 01"] <- NA
# data_rearranged$Woreda[data_rearranged$`KI Order` != "KI 01"] <- NA
# data_rearranged$Kebele[data_rearranged$`KI Order` != "KI 01"] <- NA
# data_rearranged$`KI to collect`[data_rearranged$`KI Order` != "KI 01"] <- NA


# View(data_rearranged)


############################################################################################################################################################################
#################################################Format the excel ###########################################################################################################################


create_formatted_excel(
  data = data_rearranged, 
  file_name = "HSM_2nd_Round_Sample.xlsx", 
  sheet_name = "HSM 2nd Round Sample_&_backup"
)

create_formatted_excel(
  data = ki_not_in_r2, 
  file_name = "HSM_Unsampled.xlsx", 
  sheet_name = "HSM_Unsampled.xlsx"
)
