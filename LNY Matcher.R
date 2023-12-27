# This line sets your work directory
# Set it to wherever you have the plpanda sheet stored
setwd("/Users/alanwu/Downloads/Lions/Code/LNY Matcher") 

# This line clears memory every time program runs, feel free to delete
rm(list=ls())

# Install and load libraries, uncomment install lines on first run
# install.packages("dplyr")
# install.packages("openxlsx")
library(dplyr)
library(openxlsx)

# Load Performances 2023 - 2024 sheet as a CSV
plpanda = read.csv("Penn Lions Performances and Availability - Performances 2023-2024.csv")

# Makes new variable with only LNY gigs
plpanda_lny = plpanda[2:(which(plpanda[,1] == "Fall 2023")[1] - 1),]

# Removes cancelled gigs from lny plpanda
plpanda_lny = plpanda_lny[plpanda_lny[, 1] != "Cancelled",]

# Remove info-for-troupe columns and alphabetize
col = c(10:23, 25:40)
plpanda_lny = plpanda_lny[,col]
alphabatize = order(names(plpanda_lny), decreasing = FALSE) 
plpanda_lny = plpanda_lny[,alphabatize]

# Make availability score output table
availability_score = data.frame(plpanda_lny)
availability_score = availability_score[1:30,]
availability_score[] = 0
rownames(availability_score) = colnames(availability_score)

colin = 1
# Loop through each gig
for (row in 1:nrow(plpanda_lny)) {
  # Loop through each troupe member
  for (member in 1:ncol(plpanda_lny)) {
    # Set answer for current member
    
    if (plpanda_lny[row, member] == "") {
      member_answer = "n"
    } else if (startsWith(tolower(plpanda_lny[row, member]), "y") || 
               grepl("yes", tolower(plpanda_lny[row, member])) ||
               grepl("free", tolower(plpanda_lny[row, member]))) {
      member_answer = "y"
    } else if (startsWith(tolower(plpanda_lny[row, member]), "m") || 
               grepl("maybe", tolower(plpanda_lny[row, member]))) {
      member_answer = "m"
    } else if (grepl("after", tolower(plpanda_lny[row, member]))) {
      member_answer = "y"
    } else {
      member_answer = "n"
    }
    

    
    for (col in 1:ncol(plpanda_lny)) {
      if (plpanda_lny[row, member] == "") {
        member_answer = "n"
      } else if (startsWith(tolower(plpanda_lny[row, col]), "y") || 
          grepl("yes", tolower(plpanda_lny[row, col])) ||
          grepl("free", tolower(plpanda_lny[row, col]))) {
        other_answer = "y"
      } else if (startsWith(tolower(plpanda_lny[row, col]), "m") || 
                 grepl("maybe", tolower(plpanda_lny[row, col]))) {
        other_answer = "m"
      } else if (grepl("after", tolower(plpanda_lny[row, col]))) {
        member_answer = "y"
      } else {
        other_answer = "n"
      }
      
      
      if (other_answer == "y" & member_answer == "y") {
        score = 1
      } else if (other_answer == "y" & member_answer == "m") {
        score = 0.5
      } else if (other_answer == "m" & member_answer == "y") {
        score = 0.5
      } else if (other_answer == "m" & member_answer == "m") {
        score = 0.25
      } else if (other_answer == "n" & member_answer == "n") {
        score = 1
      } else {
        score = 0
      }
      
      availability_score[member, col] = availability_score[member, col] + score
      availability_score[col, member] = availability_score[col, member] + score
    }
  }
}

# Set diagonal to N/A
x = 1
while (x <= 30) {
  availability_score[x, x] = NA
  x = x + 1
}

# Set values to percentage
availability_score = availability_score / 58

# Create output excel file
write.xlsx(availability_score, file = "Availability Scores.xlsx", sheetName = "Sheet1")
