# Install required packages if not already installed
if (!require("readxl")) install.packages("readxl", dependencies=TRUE)
if (!require("forestplot")) install.packages("forestplot", dependencies=TRUE)

# Load libraries
library(readxl)
library(forestplot)

# Load your Excel file (update the file path to your local file location)
file_path <- "Forestplot_overall_population.xlsx"  # Update the file path
data <- read_excel(file_path)

# Ensure column names are consistent
names(data) <- c("Variable", "OddsRatio", "LowerCI", "UpperCI", "Pvalue")

# Clean the Pvalue column
data$Pvalue <- gsub("[^0-9.eE-]", "", data$Pvalue)  # Remove non-numeric characters
data$Pvalue <- as.numeric(data$Pvalue)             # Convert cleaned values to numeric

# Handle any NA values in Pvalue (optional, based on your data):
data <- data[!is.na(data$Pvalue), ]  # Remove rows with NA in Pvalue

# Format the Odds Ratio and CI for display
data$OR_CI <- paste0(round(data$OddsRatio, 2), " [", 
                     round(data$LowerCI, 2), "-", 
                     round(data$UpperCI, 2), "]")

# Format p-values: show "<0.001" for very small values
data$Pvalue <- ifelse(data$Pvalue < 0.001, "<0.001", round(data$Pvalue, 3))

# Add a header row
forest_data <- rbind(
  c("Variable", "Odds Ratio (95% CI)", "P-value"),  # Header row
  cbind(data$Variable, data$OR_CI, data$Pvalue)
)

# Create the forest plot
forestplot(
  labeltext = forest_data, 
  mean = c(NA, data$OddsRatio),  # NA for header row
  lower = c(NA, data$LowerCI), 
  upper = c(NA, data$UpperCI),
  xlab = "Odds Ratio",
  is.summary = c(TRUE, rep(FALSE, nrow(data))),  # Bold the header row
  clip = c(-1.0, 10),  # Adjust the x-axis clipping range
  zero = 1,           # Line for no effect (OR = 1)
  boxsize = 0.1,      # Adjust the box size
  col = forestplot::fpColors(box = "black", line = "black", zero = "gray50"),
  title = "Forest Plot: Predictors of Sub-CR Response",
  align = c("l", "c", "r"),  # Align Variable to left, OR to center, P-value to right
  graph.pos = 2  # Positions the graph column at the center
)

