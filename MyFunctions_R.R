# MyFunctions_R.R

# Load necessary libraries
library(dplyr)
library(openxlsx)


# Define a function to process a DataFrame and a variable of interest
summarize_Multi_regression_model <- function(model) {
  # Check if the model is a glm object
  if (!inherits(model, "glm")) {
    stop("The provided model must be a glm object.")
  }
  
  # Check the family of the model
  family_name <- model$family$family
  
  if (family_name == "gaussian") {
    # Linear regression
    coef_summary <- summary(model)$coefficients
    conf_intervals <- confint(model)
    
    # Format p-values
    p_values <- coef_summary[, "Pr(>|t|)"]
    p_values_formatted <- ifelse(p_values < 0.001, "<0.001", formatC(p_values, format = "f", digits = 3))
    
    results_df <- data.frame(
      Variable = rownames(coef_summary)[rownames(coef_summary) != "(Intercept)"],
      Coefficient = formatC(coef_summary[rownames(coef_summary) != "(Intercept)", "Estimate"], format = "f", digits = 2),
      CI_2.5 = formatC(conf_intervals[rownames(coef_summary) != "(Intercept)", 1], format = "f", digits = 2),
      CI_97.5 = formatC(conf_intervals[rownames(coef_summary) != "(Intercept)", 2], format = "f", digits = 2),
      p_value = p_values_formatted[rownames(coef_summary) != "(Intercept)"]
    )
    
    # Rename columns to avoid issues with special characters
    colnames(results_df) <- c("Variable", "Coefficient", "2.5% CI", "97.5% CI", "p-value")
    
    message("The model is a linear regression (Gaussian).")
    
  } else if (family_name == "binomial") {
    # Logistic regression
    coef_summary <- summary(model)$coefficients
    odds_ratios <- exp(coef_summary[, "Estimate"])
    conf_intervals <- confint(model)
    odds_ratios_ci <- exp(conf_intervals)
    
    # Format p-values
    p_values <- coef_summary[, "Pr(>|z|)"]
    p_values_formatted <- ifelse(p_values < 0.001, "<0.001", formatC(p_values, format = "f", digits = 3))
    
    results_df <- data.frame(
      Variable = rownames(coef_summary)[rownames(coef_summary) != "(Intercept)"],
      Odds_Ratio = formatC(odds_ratios[rownames(coef_summary) != "(Intercept)"], format = "f", digits = 2),
      CI_2.5 = formatC(odds_ratios_ci[rownames(coef_summary) != "(Intercept)", 1], format = "f", digits = 2),
      CI_97.5 = formatC(odds_ratios_ci[rownames(coef_summary) != "(Intercept)", 2], format = "f", digits = 2),
      p_value = p_values_formatted[rownames(coef_summary) != "(Intercept)"]
    )
    
    # Rename columns to avoid issues with special characters
    colnames(results_df) <- c("Variable", "Odds Ratio", "2.5% CI", "97.5% CI", "p value")
    
    message("The model is a logistic regression (Binomial).")
    
  } else {
    stop("The function currently supports only Gaussian and Binomial families.")
  }
  
  # Return the DataFrame with results
  return(results_df)
}



# Function to add sheets to an existing Excel file
export_dfs_to_excel <- function(df_list, name) {
  new_file_path <- paste0(name, "_Multivariate_", format(Sys.Date(), "%m_%d_%Y"), ".xlsx")
  
  # Create a new workbook
  wb <- createWorkbook()
  
  # Add each dataframe as a separate sheet
  for (sheet_name in names(df_list)) {
    addWorksheet(wb, sheetName = sheet_name)
    writeData(wb, sheet = sheet_name, df_list[[sheet_name]], startCol = 1, startRow = 1, colNames = TRUE, rowNames = FALSE)
  }
  
  # Save the new workbook
  saveWorkbook(wb, new_file_path, overwrite = TRUE)
  
  
}


##########
#Stats:

my_fisher_test <- function(df, var1, var2) {
  # Check if the variables exist in the dataframe
  if (!(var1 %in% colnames(df)) || !(var2 %in% colnames(df))) {
    stop("One or both variable names do not exist in the dataframe.")
  }
  
  # Create a contingency table
  contingency_table <- table(df[[var1]], df[[var2]])
  
  # Print the contingency table
  cat("Contingency Table:\n")
  print(contingency_table)
  
  # Check dimensions for Fisher's test
  if (nrow(contingency_table) < 2 || ncol(contingency_table) < 2) {
    stop("Fisher's Exact Test requires a 2x2 or larger contingency table.")
  }
  
  # Perform Fisher's Exact Test
  fisher_result <- fisher.test(contingency_table)
  
  # Print the results
  cat("Fisher's Exact Test Result:\n")
  print(fisher_result)
  
  # Print the p-value
  cat("P-value:", fisher_result$p.value, "\n")
  
  # Return the results
  return(fisher_result)
}

# Function for propensity score matching
propensity_score_matching <- function(df, treatment_var, match_vars) {
  # Check if the treatment variable exists
  if (!(treatment_var %in% colnames(df))) {
    stop("The treatment variable does not exist in the dataframe.")
  }
  
  # Subset the DataFrame to only include non-missing values for the treatment variable
  df <- df[!is.na(df[[treatment_var]]), ]
  
  # Output the number of patients with each value of the treatment variable
  treatment_counts <- table(df[[treatment_var]], useNA = "ifany")
  print("Number of patients with each value of the treatment variable:")
  print(treatment_counts)
  
  # Prompt user for matching ratio
  match_ratio <- as.integer(readline(prompt = "Enter the matching ratio (e.g., 1 for 1:1, 2 for 1:2, etc.): "))
  
  # Check for missing values in matching variables
  for (var in match_vars) {
    if (!(var %in% colnames(df))) {
      stop(paste("The matching variable", var, "does not exist in the dataframe."))
    }
    n_missing_var <- sum(is.na(df[[var]]))
    if (n_missing_var > 0) {
      df <- df[!is.na(df[[var]]), ]
    }
  }
  
  # Perform propensity score matching with specified match ratio
  formula <- as.formula(paste(treatment_var, "~", paste(match_vars, collapse = " + ")))
  match_result <- matchit(formula, data = df, method = "nearest", distance = "glm", ratio = match_ratio)
  matched_data <- match.data(match_result)
  print(summary(match_result))
  print(table(matched_data[[treatment_var]]))
  
  
  
  # # Check if there are unmatched units and discard them if they exist
  # if (any(matched_data$weights == 0)) {
  #   unmatched_count <- sum(matched_data$weights == 0, na.rm = TRUE)
  #   cat("There were", unmatched_count, "unmatched patients.\n")
  #   
  #   # Keep only matched units
  #   matched_data <- matched_data[matched_data$weights > 0, ]
  #   
  #   # Output counts after matching
  #   matched_treatment_counts <- table(matched_data[[treatment_var]], useNA = "ifany")
  #   print("Number of patients with each value of the treatment variable after matching:")
  #   print(matched_treatment_counts)
  # } else {
  #   cat("All patients were matched; no unmatched units.\n")
  # }
  # Return the matched data
  return(matched_data)
}