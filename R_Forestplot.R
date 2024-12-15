# Install required packages if not already installed
if (!require("readxl")) install.packages("readxl", dependencies=TRUE)
if (!require("forestplot")) install.packages("forestplot", dependencies=TRUE)

# Load libraries
library(readxl)
library(forestplot)

# Load your Excel file (update the file path to your local file location)
file_path <- "Forestplot_hazard_ratio_OS_IC.xlsx"  # Update the file path
data <- read_excel(file_path)

# Ensure column names are consistent
names(data) <- c("Variable","Reference", "Hazardratio","LowerCI", "UpperCI", "Pvalue")

# Clean the Pvalue column
data$Pvalue <- gsub("[^0-9.eE-]", "", data$Pvalue)  # Remove non-numeric characters
data$Pvalue <- as.numeric(data$Pvalue)             # Convert cleaned values to numeric

# Handle any NA values in Pvalue (optional, based on your data):
data <- data[!is.na(data$Pvalue), ]  # Remove rows with NA in Pvalue

# Format the Hazard Ratio and CI for display
data$HR_CI <- paste0(round(data$Hazardratio, 2), " [", 
                     round(data$LowerCI, 2), "-", 
                     round(data$UpperCI, 2), "]")

# Format p-values: show "<0.001" for very small values
data$Pvalue <- ifelse(data$Pvalue < 0.001, "<0.001", round(data$Pvalue, 3))

# Add a header row
forest_data <- rbind(
  c("Variable","Reference", "Hazard Ratio (95% CI)", "P-value"),  # Header row
  cbind(data$Variable, data$Reference, data$HR_CI, data$Pvalue)
)

# Create the forest plot
forestplot(
  labeltext = forest_data, 
  mean = c(NA, data$Hazardratio),  # NA for header row
  lower = c(NA, data$LowerCI), 
  upper = c(NA, data$UpperCI),
  xlab = "Hazard Ratio",
  is.summary = c(TRUE, rep(FALSE, nrow(data))),  # Bold the header row
  clip = c(-2.0, 10),  # Adjust the x-axis clipping range
  zero = 1,           # Line for no effect (OR = 1)
  boxsize = 0.1,      # Adjust the box size
  col = forestplot::fpColors(box = "black", line = "black", zero = "gray50"),
  title = "Forest Plot: Multivariable cox proportional hazard model for overall survival (intensive chemotherapy)",
  align = c("l","c", "c", "r"),  # Align Variable to left, OR to center, P-value to right
  graph.pos = 3  # Positions the graph column at the center
)
