create_formatted_excel <- function(data, file_name, sheet_name = "formated") {
  # Define base colors
  base_colors <- c("#E4DFD4", "#F6F4F0", "#DDDEDF", "#BCBCBD", "#E8E9E9", "#C7C8CA")
  
  # Function to map unique values to colors
  get_color <- function(unique_values) {
    colors <- rep(base_colors, length.out = length(unique_values))
    setNames(colors, unique_values)
  }
  
  # Create a new workbook
  wb <- createWorkbook()
  
  # Add a worksheet
  addWorksheet(wb, sheetName = sheet_name)
  
  # Write data to the worksheet
  writeData(wb, sheet_name, data)
  
  # Get unique `admin4Pcod` and assign colors
  unique_admin4 <- unique(data$admin4Pcod)
  color_mapping <- get_color(unique_admin4)
  
  # Apply styles for each row
  for (i in seq_len(nrow(data))) {
    row_color <- color_mapping[data$admin4Pcod[i]]
    
    # Create style with background color only
    style <- createStyle(fgFill = row_color)
    
    # Apply background color style to the entire row
    addStyle(
      wb,
      sheet = sheet_name,
      style = style,
      rows = i + 1,  # +1 for the header row
      cols = 1:ncol(data),
      gridExpand = TRUE
    )
    
    # Create a combined style with background color and thick bottom border
    combined_style <- createStyle(
      fgFill = row_color,          # Preserve the background color
      border = c("top", "bottom", "right"), # Borders to apply
      borderColour = "#9A9A9C",    # Border color
      borderStyle = c("thick", "thick", "thin") # Border styles
    )
    
    # Define specific columns for additional styling
    col_indices <- c(
      which(colnames(data) == "KI community role"),
      which(colnames(data) == "KI Phone number"),
      which(colnames(data) == "Comment"),
      which(colnames(data) == "Partner organization Name"),
      which(colnames(data) == "Focal person E-mail Address from Partner organization")
    )
    
    # Apply combined style to the specific columns
    addStyle(
      wb,
      sheet = sheet_name,
      style = combined_style,
      rows = i + 1,  # Rows including header
      cols = col_indices,
      gridExpand = TRUE
    )
  }
  
  # Set column widths
  setColWidths(
    wb,
    sheet = sheet_name,
    cols = seq_len(ncol(data)),
    widths = c(8, 11, 19, 11, 27, 11, 29, 11, 10, 15, 10, 14, 14, 16, 21, 40)
  )
  
  # Header style
  header_style <- createStyle(
    fontSize = 9,
    fontColour = "black",
    halign = "left",
    fgFill = "#EE5859",
    border = "TopBottomLeftRight",
    borderColour = "black",
    borderStyle = "thick",
    textDecoration = "bold"
  )
  
  # Apply header style
  addStyle(
    wb,
    sheet = sheet_name,
    style = header_style,
    rows = 1,  # Header row
    cols = seq_len(ncol(data)),
    gridExpand = TRUE)
  
  # Freeze the header row
  freezePane(wb, sheet_name, firstActiveRow = 2)
  
  # Add column filters
  addFilter(wb, sheet = sheet_name, rows = 1, cols = seq_len(ncol(data)))
  
  # Save the workbook
  saveWorkbook(wb, file_name, overwrite = TRUE)
  
  message("Excel file created and saved to: ", file_name)
}
