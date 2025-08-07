### === Co-author list organizer: creates author + affiliation formatting for papers === ###
### Output a list in a Word document or as an HTML string. 
### Requires an input Excel file with author info: names, affiliations, etc.

#### Load Required Packages ####
# Install packages if not already installed:
# install.packages("tidyverse")

library(tidyverse)  # For data manipulation
library(readxl)     # For reading Excel files
library(officer)    # For creating Word documents

### Clean Workspace and Set Locale (For non-Windows systems, adjust accordingly)
rm(list = ls())  # Clear all existing objects in the environment
Sys.setlocale(category = "LC_ALL", locale = "English_United States.1252") # Sets locale to ensure correct display of special characters (accents, etc.)

### If not working from the Rproj file, you will have to set the Working Directory (i.e., path where data is located)
#wd <- "C:/Path/To/File/"
#setwd(wd)

### Read Author Data from Excel file
# It should have a sheet named "Author Info", with metadata starting on row 2

authors_affil_long <- read_excel("Example_author_info.xlsx", sheet = "Author Info", na = "NA", skip = 1) %>% 
  arrange(order)

### Prepare Long Format Author-Affiliation Links 
# This converts affiliations from multiple columns into long format
affil_per_author <- authors_affil_long %>% 
  mutate(order = as.numeric(order)) %>% 
  pivot_longer(cols = -c(order, abrev, name_in_publication, surname, first_name, ORCID, email), 
               names_to = "aff_nr", values_to = "aff_name", values_drop_na = FALSE) %>% 
  mutate(aff_nr = as.numeric(substr(aff_nr, 8,8))) %>%  # Extract affiliation number (e.g., 1, 2)
  filter(!is.na(aff_name))  # Keep only valid affiliations

### Order Affiliations Globally 
# Generate a global ordering of unique affiliations
affil_order <- affil_per_author %>% 
  arrange(order, aff_nr) %>% 
  mutate(aff_order_temp = 1:nrow(affil_per_author)) %>% 
  group_by(aff_name) %>% 
  summarise(aff_order_temp = min(aff_order_temp)) %>% 
  arrange(aff_order_temp) %>% 
  mutate(aff_order = row_number()) %>% 
  select(-aff_order_temp)

### Create Author-Affiliation Mapping 
# Associate each author with the global affiliation number
affil_per_author_final <- affil_per_author %>% 
  left_join(affil_order) %>% 
  select(name_in_publication, aff_name, aff_order, aff_nr) %>% 
  mutate(aff_nr_name = paste(aff_order, aff_name, sep = "_")) %>% 
  mutate(aff_nr_txt = paste("aff_", aff_nr, sep = "")) %>% 
  select(-aff_name, -aff_nr, -aff_order) %>% 
  pivot_wider(names_from = aff_nr_txt, values_from = aff_nr_name) %>% 
  separate(aff_1, into = c("aff1_order", "aff1_name"), sep = "_") %>% 
  separate(aff_2, into = c("aff2_order", "aff2_name"), sep = "_")

# You can save this list if desired
#write.table(affil_per_author_final, file = "affil_per_author.txt", row.names = F, col.names = T, sep = "\t")

# Construct superscript references (e.g., Name<sup>1,2</sup>)
author_with_number <- affil_per_author_final %>% 
  select(!ends_with("_name")) %>% 
  unite(aff_nrs, ends_with("_order"), na.rm = TRUE, sep = ",") %>% 
  mutate(author_aff_nrs = paste0(name_in_publication, "<sup>", aff_nrs, "</sup>, "))

#### Output Option 1: Word Document ####
# Define function to format superscripts
superscript_text <- function(text) {
  ftext(text, prop = fp_text(font.family = "Times New Roman", font.size = 12, vertical.align = "superscript"))
}

# Create and initialize Word document
doc <- read_docx()

# Format author line
author_string <- list()
for (i in 1:nrow(author_with_number)) {
  author <- author_with_number$name_in_publication[i]
  aff_nrs <- author_with_number$aff_nrs[i]
  author_string <- append(author_string, list(ftext(author, prop = fp_text(font.family = "Times New Roman", font.size = 12))))
  for (char in strsplit(aff_nrs, "")[[1]]) {
    author_string <- append(author_string, list(superscript_text(char)))
  }
  if (i < nrow(author_with_number)) {
    author_string <- append(author_string, list(ftext(", ", prop = fp_text(font.family = "Times New Roman", font.size = 12))))
  }
}
author_fpar <- do.call(fpar, author_string)
doc <- body_add_fpar(doc, author_fpar, style = "Normal") # adds author line to doc

# Format affiliations line
doc <- body_add_par(doc, "", style = "Normal")  # blank line
affiliations_string <- list()
for (i in 1:nrow(affil_order)) {
  aff_number <- affil_order$aff_order[i]
  aff_name <- affil_order$aff_name[i]
  affiliations_string <- append(affiliations_string, list(superscript_text(as.character(aff_number))))
  affiliations_string <- append(affiliations_string, list(ftext(aff_name, prop = fp_text(font.family = "Times New Roman", font.size = 12))))
  affiliations_string <- append(affiliations_string, list(ftext(". ", prop = fp_text(font.family = "Times New Roman", font.size = 12))))
}
affiliations_fpar <- do.call(fpar, affiliations_string)
doc <- body_add_fpar(doc, affiliations_fpar, style = "Normal") # adds affil line to doc

# Save output Word file
print(doc, target = "output_author_list.docx")

#### Output Option 2: HTML Text for Manual Use ####
# Produces html formatted text to paste in https://www.w3schools.com/html/tryit.asp?filename=tryhtml_editor

affil_list <- affil_order %>% 
  mutate(aff_name_point = paste0(aff_name, ". ")) %>% 
  mutate(aff_nr_txt = paste0("<sup>", aff_order, "</sup>", aff_name_point))

n_end <- nchar(paste(unlist(author_with_number$author_aff_nrs), collapse ="")) - 2

# Copy HTML text to clipboard (for pasting into Word via online HTML viewer)
writeClipboard(paste(
  substr(paste(unlist(author_with_number$author_aff_nrs), collapse = ""), 1, n_end),
  paste(unlist(affil_list$aff_nr_txt), collapse = ""),
  sep = "<p></p>"
))
