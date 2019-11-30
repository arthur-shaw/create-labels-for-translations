
translate_varLablels = function(
	input_path, 	# file path to Excel file input
	output_path 	# desired file path to Stata .do file output
) {

# =============================================================================
# Load necessary libraries
# =============================================================================

# packages needed for this program 
packagesNeeded <- c(
	"readxl", 	# for ingesting Excel translation file
	"dplyr",	# for convenient data wrangling
	"stringr", 	# for removing undesirable content from translations
	"readr" 	# for writing text to Stata .do file
)														

# identify and install those packages that are not already installed
packagesToInstall <- packagesNeeded[!(packagesNeeded %in% installed.packages()[,"Package"])]
if(length(packagesToInstall)) install.packages(packagesToInstall, quiet = TRUE, repos = 'https://cloud.r-project.org/', dep = TRUE)

# load all needed packages
lapply(packagesNeeded, library, character.only = TRUE)

# =============================================================================
# Injest translation file
# =============================================================================

# confirm that file exists

# confirm that file is Excel

# injest content
translation <- read_excel(path = input_path, sheet = "Translations", col_names = TRUE)

# confirm that needed columns exist
	# Variable
	# Type
	# Translation

# =============================================================================
# Form variable label commands in Stata syntax
# =============================================================================

# filter down to question text
variables <- translation %>% filter(!is.na(Variable) & Type == "Title")

# construct Stata commands
variables <- variables %>% mutate(

	# remove HTML tags
	text_woLastTag = str_replace_all(Translation, pattern = "</[:alnum:]+>", replacement = ""),
	text_noHTML = str_replace_all(text_woLastTag, pattern = '<[[ ][:alpha:]]+[ ]*=[ ]*\"[#]*[:alnum:]+\">', replacement = ""),

	# remove newline and carriage return markers, if any
	text_noNewline = str_replace_all(Translation, pattern = "\\n|\\r", replacement = ""),

	# prepend Stata label syntax
	varLabel = paste0("capture label variable ", Variable, ' `\"', text_noHTML, '\"\';') # note: `;` used as delimiter
	)

# =============================================================================
# Write commands to file as Stata .do file
# =============================================================================

# write Stata commands to file
readr::write_lines("#delim;", path = output_path) # note: `;` used as delimiters because a few strings go to a new line for unknown reasons
readr::write_lines(variables$varLabel, path = output_path, append = TRUE)

}



translate_valLablels = function(
	input_path, 	# file path to Excel file input
	output_path, 	# desired file path to Stata .do file output
	label_prefix 	# prefix to label name
) {

# =============================================================================
# Load necessary libraries
# =============================================================================

# packages needed for this program 
packagesNeeded <- c(
	"readxl", 	# for ingesting Excel translation file
	"dplyr",	# for convenient data wrangling
	"stringr", 	# for removing undesirable content from translations
	"readr" 	# for writing text to Stata .do file
)														

# identify and install those packages that are not already installed
packagesToInstall <- packagesNeeded[!(packagesNeeded %in% installed.packages()[,"Package"])]
if(length(packagesToInstall)) install.packages(packagesToInstall, quiet = TRUE, repos = 'https://cloud.r-project.org/', dep = TRUE)

# load all needed packages
lapply(packagesNeeded, library, character.only = TRUE)

# =============================================================================
# Injest translation file
# =============================================================================

# confirm that file exists

# confirm that file is Excel

# injest content
translation <- read_excel(path = input_path, sheet = "Translations", col_names = TRUE)

# confirm that needed columns exist
	# Variable
	# Type
	# Translation

# =============================================================================
# Form variable label commands in Stata syntax
# =============================================================================

# filter down to answer options
defineLabels <- translation %>% filter(!is.na(Variable) & Type == "OptionTitle")

# -----------------------------------------------------------------------------
# Define value labels
# -----------------------------------------------------------------------------

# create command to define value labels
defineLabels <- mutate(defineLabels, 
	labelName = paste0(label_prefix, "_", Variable),
	command = paste0("label define ", labelName, " ", Index, " ", '`\"', Translation, '\"\'', ", add;")) # note: `;` used as delimiter

# write label defintions to a .do file
readr::write_lines("#delim;", path = output_path) # note: `;` used as delimiters because a few strings go to a new line for unknown reasons
readr::write_lines(defineLabels$command, path = output_path, append = TRUE)

# -----------------------------------------------------------------------------
# Attach value labels to variables
# -----------------------------------------------------------------------------

# create commands to attach value labels to variables
attachLabels <- defineLabels %>% distinct(Variable, .keep_all = TRUE) %>% 
	mutate(command = paste0("capture label values ", Variable, " ", labelName, ";"))

# =============================================================================
# Write commands to file as Stata .do file
# =============================================================================

# add commands to previously created file
readr::write_lines(attachLabels$command, path = output_path, append = TRUE)

}


# define file folder and file
fileDir <- "C:/Users/wb393438/Survey Solutions/create-labels-for-translations/"
excelFile <- "[LSMS PLUS QUESTIONNAIRE 2019 11o]LSMS PLUS  CAMBODIA QUESTIONNAIRE 2019 (1).xlsx"


# take data from the Excel translation file to create a .do file named varlabels.do
translate_varLablels(input_path = paste0(fileDir, excelFile), output_path = paste0(fileDir, "varlabels.do"))

# take data from the Excel translation file to create a .do file named val_labels.do
translate_valLablels(input_path = paste0(fileDir, excelFile), output_path = paste0(fileDir, "val_labels.do"), label_prefix = "kh")
