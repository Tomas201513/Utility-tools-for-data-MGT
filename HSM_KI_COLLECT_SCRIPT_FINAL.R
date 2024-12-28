library(dplyr)
library(readxl)
library(openxlsx)


# Load the data
# data <- read_excel("C:/Users/test/Downloads/ET_2025HSM_v2024-12-18_sample (2).xlsx", sheet = "sample")
# data <- read_excel("C:/Users/test/Downloads/ET_2025HSM_v2024-12-18_sample (1).xlsx", sheet = "sample")

data <- read_excel("C:/Users/test/Downloads/ET_2025HSM_v2024-12-18_sample_down_12_26_2024.xlsx", sheet = "sample")
View(data)

# filter the data where 1st_round_HSM_sampling is sample or backup
data <- data[data$`2nd_round_HSM_sampling` %in% c("sample", "backup"),]

# filter columns admin4name	admin4Pcod	admin3name	admin3Pcod	admin2name	admin2Pcod	admin1name	admin1Pcod number_of_ki_identification_needed priority_for_Ki_Identification
dataa <- data[,c("admin4name", "admin4Pcod", "admin3name", "admin3Pcod", "admin2name", "admin2Pcod", "admin1name", "admin1Pcod", "2nd_round_HSM_sampling","number_of_ki_identification_needed", "priority_for_Ki_Identification")]
# View(dataa)

# filter the data where number_of_ki_identification_needed is less than 5
dataa <- dataa[dataa$number_of_ki_identification_needed > 0,]
# View(dataa)

# add "Kebele_" to the admin4name column where admin4name is number only
dataa$admin4name <- ifelse(grepl("^[0-9]+$", dataa$admin4name), paste("Kebele ", dataa$admin4name, sep = ""), dataa$admin4name)

# Duplicate each row based on the 'number_of_ki_identification_needed' column
dataaa <- dataa[rep(seq_len(nrow(dataa)), dataa$number_of_ki_identification_needed), ]

# Create a unique count column for each duplicate
dataaa$Ki_Identification <- ave(dataaa$number_of_ki_identification_needed, dataaa$admin4Pcod, FUN = function(x) seq_along(x))

# add "KI" before Ki_Identification Values
dataaa$Ki_Identification <- paste("KI", dataaa$Ki_Identification, sep = " 0")

# # Arrange specific columns in order and include the rest
# data_rearranged <- dataaa %>%
#   select(admin1name, admin1Pcod, admin2name, admin2Pcod, admin3name, 
#          admin3Pcod, admin4name, admin4Pcod,`2nd_round_HSM_sampling`, 
#          priority_for_Ki_Identification,number_of_ki_identification_needed,Ki_Identification)

# rename column names
data_rearranged <- dataaa %>%
  rename(
    Region = admin1name,
    Zone = admin2name,
    Woreda = admin3name,
    Kebele = admin4name,
    Sampling = `2nd_round_HSM_sampling`,
    # priority_for_Ki_Identification = priority_for_Ki_Identification,
    `KI to collect` = number_of_ki_identification_needed,
    `KI Order` = Ki_Identification
  ) %>%
  select(Region, admin1Pcod, Zone, admin2Pcod, Woreda, 
         admin3Pcod, Kebele, admin4Pcod, Sampling, 
         # priority_for_Ki_Identification,
         `KI to collect`, `KI Order`)

# add new columns KI community role,	KI Phone number,	Comment, Partner organization Name and	Focal person E-mail Address from Partner organization
data_rearranged$`KI community role` <- NA
data_rearranged$`KI Phone number` <- NA
data_rearranged$Comment <- NA
data_rearranged$`Partner organization Name` <- NA
data_rearranged$`Focal person E-mail Address from Partner organization` <- NA

# View(data_rearranged)

# make value in Region,Zone, Woreda and Kebele NA where  different from KI 01
data_rearranged$Region[data_rearranged$`KI Order` != "KI 01"] <- NA
data_rearranged$Zone[data_rearranged$`KI Order` != "KI 01"] <- NA
data_rearranged$Woreda[data_rearranged$`KI Order` != "KI 01"] <- NA
data_rearranged$Kebele[data_rearranged$`KI Order` != "KI 01"] <- NA
data_rearranged$`KI to collect`[data_rearranged$`KI Order` != "KI 01"] <- NA



# View(data_rearranged)


############################################################################################################################################################################
#################################################Format the excel ###########################################################################################################################


# Define base colors
base_colors <- c("#E4DFD4", "#F6F4F0", "#DDDEDF", "#BCBCBD", "#E8E9E9", "#C7C8CA")

# Function to map unique values to colors
get_color <- function(unique_values) {
  colors <- rep(base_colors, length.out = length(unique_values))
  setNames(colors, unique_values)
}

# Split data by `admin1name`
data_by_admin1 <- split(data_rearranged, data_rearranged$admin1Pcod)

for (admin1 in names(data_by_admin1)) {
  data_admin1 <- data_by_admin1[[admin1]]
  data_by_admin2 <- split(data_admin1, data_admin1$admin2Pcod)
  
  # Create a new workbook
  wb <- createWorkbook()
  
  for (admin2 in names(data_by_admin2)) {
    data_admin2 <- data_by_admin2[[admin2]]
    
    # Add a worksheet
    addWorksheet(wb, sheetName = admin2)
    
    # Write data to the worksheet
    writeData(wb, admin2, data_admin2)
    
    # Get unique `admin4Pcod` and assign colors
    unique_admin4 <- unique(data_admin2$admin4Pcod)
    color_mapping <- get_color(unique_admin4)
    
    for (i in seq_len(nrow(data_admin2))) {
      row_color <- color_mapping[data_admin2$admin4Pcod[i]]
      
      # Create style with background color only
      style <- createStyle(fgFill = row_color)
      
      # Apply background color style to the entire row
      addStyle(
        wb,
        sheet = admin2,
        style = style,
        rows = i + 1,  # +1 for the header row
        cols = 1:ncol(data_admin2),
        gridExpand = TRUE
      )
      
      # Create a combined style with background color and thick bottom border
      combined_style <- createStyle(
        fgFill = row_color,          # Preserve the background color
        border = c("top", "bottom", "right"), # Borders to apply
        borderColour = "#9A9A9C",    # Border color
        borderStyle = c("thick", "thick", "thin") # Border styles
      )
      
      # Find column indices for `number_of_ki_identification_needed` and `Ki_Identification`
      KI_community_role <- which(colnames(data_admin2) == "KI community role")
      KI_Phone_number <- which(colnames(data_admin2) == "KI Phone number")
      Comment <- which(colnames(data_admin2) == "Comment")
      Partner_organization_Name <- which(colnames(data_admin2) == "Partner organization Name")
      Focal_person_Address <- which(colnames(data_admin2) == "Focal person E-mail Address from Partner organization")
      
      # Apply combined style to the specific columns
      addStyle(
        wb,
        sheet = admin2,
        style = combined_style,
        rows = i + 1,  # Rows including header
        cols = c(KI_community_role, KI_Phone_number, Comment, Partner_organization_Name, Focal_person_Address),
        gridExpand = TRUE
      )
      
      
      # Set column widths
      setColWidths(wb,  sheet = admin2, cols = c(1, 3,   5, 7,  11,12,13,14,15,16), 
                   widths = c(8, 16, 16, 16, 8, 25,14,14,20,40))
      
      # Header style
      header_style <- createStyle(
        # wrapText = TRUE,
        fontSize = 9,
        fontColour = "black",
        halign = "center",
        fgFill = "#EE5859",
        border = "TopBottomLeftRight",
        borderColour = "black",
        borderStyle = "thick",
        textDecoration = "bold"
      )
      
      # Apply header style
      addStyle(
        wb,
        sheet = admin2,
        style = header_style,
        rows = 1,  # Header row
        cols = 1:ncol(data_admin2),
        gridExpand = TRUE
      )
      freezePane(wb, admin2, firstActiveRow = 2, firstActiveCol = "L")
      
    }
  }
  
  # Save the workbook
  file_name <- paste0("admin1_", admin1, ".xlsx")
  saveWorkbook(wb, file_name, overwrite = TRUE)
}

message("Excel files created for each admin1name with validations and styles applied!")
