library(tidyr)
library(caret)
library(ggplot2)
library(dplyr)
library(MatchIt)
library(readxl)
library(openxlsx)
library("survival")
library("survminer")
#library("xlsx")
library(dplyr)
library(writexl)
#install.packages("openxlsx")
library(openxlsx)


#### Regression func
base_dir <- "/Users/vardhaanambati/Research Scripts"
functions_path <- file.path(base_dir, "MyFunctions_R.R")
if (file.exists(functions_path)) {
  source(functions_path)
  print("Functions sourced successfully.")
} else {
  stop("Functions file not found.")
}

### Import File
#file_path <- "SFGH Trauma C1 C2 Fracture Mastersheet _08132024_clean.xlsx"
c12_df <- read.csv("/Users/vardhaanambati/Desktop/Research/Trauma/secure_ C1-C2 patient chart review/Datasets/C12_consec_multifriendly_dataset_Apr2025_Final.csv")
nrow(c12_df)


### Variable Type Casting
c12_df$`Expired.during.Admission` <- as.factor(c12_df$`Expired.during.Admission`)
c12_df$`Female` <- as.factor(c12_df$`Female`)
c12_df$`HTN.` <- as.factor(c12_df$`HTN.`)
c12_df$`DM.` <- as.factor(c12_df$`DM.`)
c12_df$`Neurodeficits` <- as.factor(c12_df$`Neurodeficits`)
c12_df$`Neurodeficits`
c12_df$`MRI` <- as.factor(c12_df$`MRI`)


c12_df$`C1.Landells.Type.I` <- as.factor(c12_df$`C1.Landells.Type.I`)
c12_df$`C1.Landells.Type.2` <- as.factor(c12_df$`C1.Landells.Type.2`)
c12_df$`C2.Dens.Anderson.and.D.Alonzo.Type.2` <- as.factor(c12_df$`C2.Dens.Anderson.and.D.Alonzo.Type.2`)
c12_df$`C2.Dens.Anderson.and.D.Alonzo.Type.3` <- as.factor(c12_df$`C2.Dens.Anderson.and.D.Alonzo.Type.3`)
c12_df$`C2.Dens.Anderson.and.D.Alonzo.Type.1.or.3` <- as.factor(c12_df$`C2.Dens.Anderson.and.D.Alonzo.Type.1.or.3`)
c12_df$`C2.lateral.mass.fracture` <- as.factor(c12_df$`C2.lateral.mass.fracture`)
c12_df$`Occipital.Condyle.Fracture` <- as.factor(c12_df$`Occipital.Condyle.Fracture`)

c12_df$`Injury.Severity.Score` <- as.numeric(c12_df$`Injury.Severity.Score`)

######### PROPENSITY MATCHING
# match_vars <- c('Age.of.Injury', 'Female', 'Injury.Severity.Score', 'Neurodeficits')
# c12_consec_died_df <- propensity_score_matching(c12_df, "Expired.during.Admission", match_vars)
# write.csv(c12_consec_died_df, file = "Datasets/c12_consec_died_PSM.csv", row.names = FALSE)
# #  
# c12_consec_surg_df <- propensity_score_matching(c12_df, "Surgery", match_vars)
# write.csv(c12_consec_surg_df, file = "Datasets/c12_consec_surg_PSM.csv", row.names = FALSE)
# #  
# c12_consec_union_df <- propensity_score_matching(c12_df, "NonUnion", match_vars)
# write.csv(c12_consec_union_df, file = "Datasets/c12_consec_Nonunion_PSM.csv", row.names = FALSE)

##########                     ##########
##########  CONSECUTIVE CASES  ##########
##########                     ##########
c12_consecutive_df <- c12_df[c12_df$C1.and.2.Injury == 1, ]
nrow(c12_consecutive_df)
name = "C12_CONSECUTIVE_Cases"

### Predictors of Death
death_formula_consec <- "Expired.during.Admission ~ Female + HTN. + Injury.Severity.Score + Neurodeficits + C2.Dens.Anderson.and.D.Alonzo.Type.1.or.3 + Surgery"
death_formula_consec <- as.formula(death_formula_consec)
model_Death_Consec <- glm(death_formula_consec, data = c12_consecutive_df, na.action = na.omit, family = 'binomial')
summary(model_Death_Consec)
death_predictors_Consec_df <- summarize_Multi_regression_model(model_Death_Consec)
death_predictors_Consec_df

### Predictors of Surgery
surg_formula_consec <- "Surgery ~ Female + DM. + Injury.Severity.Score + MRI + C2.Dens.Anderson.and.D.Alonzo.Type.1.or.3 + C2.Dens.Anderson.and.D.Alonzo.Type.2"
surg_formula_consec <- as.formula(surg_formula_consec)
model_Surg_consec <- glm(surg_formula_consec, data = c12_consecutive_df, na.action = na.omit, family = 'binomial')
summary(model_Surg_consec)
surg_predictors_Consec_df <- summarize_Multi_regression_model(model_Surg_consec)
surg_predictors_Consec_df


### Predictors of Union
nonUnion_formula_consec <- "NonUnion ~ Age.of.Injury + Mechanism.of.Injury + Injury.Severity.Score + C2.Dens.Anderson.and.D.Alonzo.Type.2 + C2.lateral.mass.fracture"
nonUnion_formula_consec <- as.formula(nonUnion_formula_consec)
non_union_consec_df <- c12_consecutive_df[c12_consecutive_df$Expired.during.Admission != 1, ]
non_union_consec_df <- non_union_consec_df[!is.na(non_union_consec_df$NonUnion), ]
nrow(non_union_consec_df)

model_NonUnion_Consec <- glm(nonUnion_formula_consec, data = non_union_consec_df, na.action = na.omit, family = 'binomial')
summary(model_NonUnion_Consec)
NonUnion_predictors_Consec_df <- summarize_Multi_regression_model(model_NonUnion_Consec)
NonUnion_predictors_Consec_df



##
Multi_dfs <- list("Expired Inpatient (Multi Log)" = death_predictors_Consec_df,
                   "Surgery Analysis (Multi Log)" = surg_predictors_Consec_df)#,
                  #"NonUnion Analysis (Multi Log)" = NonUnion_predictors_Consec_df)
export_dfs_to_excel(Multi_dfs, name)


####
# Install and load necessary packages
library(tidyverse)
library(ggplot2)
library(reshape2)
library(scales)

# Fracture types
# Original human-readable fracture types
fracture_types2 <- c(
  "C1 Landells Type 1 or 2",
  "C1 Landells Type 3",
  "C2 Dens Anderson and D'Alonzo Type 1 or 3",
  "C2 Dens Anderson and D'Alonzo Type 2",
  "C1 or C2 transverse foramen fracture",
  "C2 lateral mass fracture",
  "C2 Hangman's fracture",
  "Occipital Condyle Fracture"
)

# Replace spaces with dots for modeling
fracture_types_model <- gsub(" ", ".", fracture_types2)
fracture_types_model <- gsub("'", ".", fracture_types_model)

# Identify missing variables
missing_vars <- fracture_types_model[!(fracture_types_model %in% colnames(c12_consecutive_df))]

# Print the missing variables if any
if (length(missing_vars) > 0) {
  print("Missing variables:")
  print(missing_vars)
} else {
  print("All variables are present in the dataframe.")
}

# Custom replacements for pretty labels
replacement_dict <- c(
  "C1 Landells Type 1 or 2" = "C1 Landells Type I or II",
  "C1 Landells Type 3" = "C1 Landells Type III",
  "C2 Dens Anderson and D'Alonzo Type 1 or 3" = "C2 Odontoid Type I or III",
  "C2 Dens Anderson and D'Alonzo Type 2" = "C2 Odontoid Type II"
)

# Create pretty label dictionary using model names
default_labels <- setNames(fracture_types2, fracture_types_model)
full_label_dict <- default_labels
custom_dot_keys <- gsub(" ", ".", names(replacement_dict))
full_label_dict[custom_dot_keys] <- replacement_dict

# Function to retrieve pretty labels
pretty_labels <- function(label) {
  return(full_label_dict[[label]])
}

# Initialize matrix for results
results_matrix <- matrix(NA, nrow = length(fracture_types_model), ncol = length(fracture_types_model))
rownames(results_matrix) <- fracture_types_model
colnames(results_matrix) <- fracture_types_model

# Fill matrix with ORs where significant
for (i in 1:length(fracture_types_model)) {
  for (j in length(fracture_types_model):1) {
    outcome <- fracture_types_model[i]
    predictor <- fracture_types_model[j]

    if (i == j) {
      results_matrix[outcome, predictor] <- -1  # black diagonal
      next
    }
    
    # Prepare data for modeling
    df_temp <- c12_consecutive_df[, c(outcome, predictor)] %>%
      mutate(across(everything(), ~as.numeric(as.character(.)))) %>%
      drop_na()
    
    if (length(unique(df_temp[[outcome]])) < 2 || length(unique(df_temp[[predictor]])) < 2) next
    
    # Fit logistic regression model
    model <- glm(as.formula(paste0("`", outcome, "` ~ `", predictor, "`")),
                 data = df_temp, family = "binomial")
    coef_summary <- summary(model)$coefficients
    p_value <- coef_summary[2, 4]
    or_value <- exp(coef(model)[2])
    
    # Store OR if p-value is significant
    if (p_value < 0.05) {
      results_matrix[outcome, predictor] <- or_value
    }
  }
}

# Format for ggplot
results_df <- melt(results_matrix, na.rm = FALSE)
colnames(results_df) <- c("Outcome", "Predictor", "OR")

# Create labels for OR values
results_df$label <- ifelse(results_df$Outcome == results_df$Predictor, "",
                           ifelse(is.na(results_df$OR), "", round(results_df$OR, 2)))

# Apply fill color based on OR values
results_df$fill_color <- case_when(
  is.na(results_df$OR) ~ NA_real_,
  results_df$OR == -1 ~ -999,
  TRUE ~ results_df$OR
)

# Apply pretty labels
results_df$Outcome <- sapply(results_df$Outcome, pretty_labels)
results_df$Predictor <- sapply(results_df$Predictor, pretty_labels)

results_df$Predictor <- recode(
  results_df$Predictor,
  "C2 Dens Anderson and D'Alonzo Type 1 or 3" = "C2 Odontoid Type I or III",
  "C2 Dens Anderson and D'Alonzo Type 2" = "C2 Odontoid Type II"
)

results_df$Outcome <- recode(
  results_df$Outcome,
  "C2 Dens Anderson and D'Alonzo Type 1 or 3" = "C2 Odontoid Type I or III",
  "C2 Dens Anderson and D'Alonzo Type 2" = "C2 Odontoid Type II"
)

# Flip the order of the predictor variable by reversing the factor levels
#results_df$Predictor <- factor(results_df$Predictor, levels = rev(levels(results_df$Predictor)))

# Plot
p <- ggplot(results_df, aes(x = Predictor, y = Outcome, fill = fill_color)) +
  geom_tile(color = "black", size = 1.5) +
  geom_text(aes(label = label), fontface = "bold", size = 6, family = "Arial") +
  scale_fill_gradient2(
    low = "blue", high = "red", mid = "white", midpoint = 1,
    na.value = "white", name = "Odds Ratio",
    limits = c(-1,9),
    breaks = c(-1, 1, 5, 9),  # increase tick resolution
    guide = guide_colorbar(
      barwidth = 1.5,
      barheight = 30,
      title.position = "top",
      title.hjust = 0.5,
      ticks = TRUE,
      ticks.linewidth = 1,
      frame.colour = "black",       # border around the bar only
      frame.linewidth = 0.8
    )
  ) +
  geom_segment(
    aes(x = 1, y = nrow(results_matrix), xend = ncol(results_matrix), yend = 1),
    color = "black", size = 0.8, linetype = "dashed"
  ) +
  scale_y_discrete(limits = rev) +
  theme_minimal(base_family = "Arial") +
  theme(
    axis.text.x = element_text(angle = 45, hjust = 1, size = 15, face = "bold", family = "Arial"),
    axis.text.y = element_text(size = 15, face = "bold", family = "Arial"),
    axis.title = element_blank(),
    panel.grid = element_blank(),
    legend.title = element_text(face = "bold", size = 14),
    legend.text = element_text(size = 14),  # ← Increase label font size
    legend.background = element_blank(),  # no box around the whole legend
    legend.box.background = element_blank()
  )

p
#ggsave("/Users/vardhaanambati/Desktop/Research/Trauma/secure_ C1-C2 patient chart review/Write Up/Heatmap_v2.png", plot = p, dpi = 600, bg = "white")


# Print out the table with Outcome, Predictor, and OR
print("Outcome, Predictor, and OR Table:")
print(results_df)



colnames(c12_consecutive_df)


model <- glm('C1.Landells.Type.1.or.2 ~ Occipital.Condyle.Fracture', data = c12_consecutive_df, family = binomial)
model <- glm('C2.Dens.Anderson.and.D.Alonzo.Type.2 ~ C1.or.C2.transverse.foramen.fracture', data = c12_consecutive_df, family = binomial)

summary(model)
coefficients <- summary(model)$coefficients
OR <- exp(coefficients["Occipital.Condyle.Fracture!", "Estimate"])
p_value <- coefficients["C1.or.C2.transverse.foramen.fracture", "Pr(>|z|)"]
cat("Odds Ratio: ", OR, "\n")
cat("P-value: ", p_value, "\n")






library(survival)
library(survminer)

# Load and clean the data
c12_df <- read.csv("/Users/vardhaanambati/Desktop/Research/Trauma/secure_ C1-C2 patient chart review/Datasets/C12_consec_multifriendly_dataset_Apr2025_Final.csv")
c12_df <- c12_df[c12_df$`Discharge.Disposition` != "EXPIRED", ]
c12_df$Survival_Months <- as.numeric(c12_df$Survival_Months)
c12_df$Dead <- as.numeric(c12_df$Dead)

# Create survival object and fit
surv_obj <- Surv(time = c12_df$Survival_Months, event = c12_df$Dead)
fit <- survfit(surv_obj ~ 1)

# Plot with all requested customizations
library(survminer)

# Example survival plot
fit <- survfit(Surv(time, status) ~ sex, data = lung)

library(survminer)

# Example survival plot
fit <- survfit(Surv(time, status) ~ sex, data = lung)

km_plot <- ggsurvplot(fit, 
                      data = c12_df, 
                      conf.int = FALSE, 
                      risk.table = TRUE, 
                      risk.table.height = 0.25, 
                      risk.table.y.text.col = TRUE, 
                      risk.table.y.text = FALSE, 
                      xlab = "Time (Months)", 
                      ylab = "Survival Probability", 
                      title = "Overall Survival: Discharged Alive", 
                      ggtheme = theme_minimal(), 
                      surv.median.line = "hv",  # horizontal + vertical line at median survival 
                      risk.table.title = "Number at risk", 
                      tables.theme = theme_classic() + 
                        theme(axis.text = element_text(size = 14),
                              axis.title.x = element_text(size = 16), 
                              axis.title.y = element_text(size = 16)),           
                      risk.table.border = TRUE,        # adds box around table
                      font.tickslab = 14,              # increase size of axis tick labels
                      font.main = 16,                 # increase title size
                      font.x = 16,                    # increase x-axis font size
                      font.y = 16,                    # increase y-axis font size
                      font.table = 14, 
                      font.family = "Arial",          # change font to Arial
                      font.face = "bold",             # bold font
                      title.position = "center",      # properly center the title
                      risk.table.y.text.size = 4,     # Adjust text size of risk table y-axis
                      risk.table.x.text.size = 4)

# Now, add the axis lines separately
km_plot$plot <- km_plot$plot + 
  theme(axis.line = element_line(color = "black", size = .5))  # Adds x and y axis lines

# Print the final plot
print(km_plot)
surv_median(fit)
















average_score <- mean(c12_df$`Injury Severity Score`, na.rm = TRUE)
c12_df$Binary_Injury_Severity <- ifelse(c12_df$`Injury Severity Score` > average_score, 1, 0)


c12_censored_df <- c12_df %>%
  mutate(
    # Censoring at 12 months (1 year)
    `Known Survival Length` = ifelse(`Known Survival Length` >= 12, 12, `Known Survival Length`),
    
    # Updating Known_deceased based on the new Known_Survival_Length
    `Known deceased` = ifelse(`Known Survival Length` >= 12 & `Known deceased` == 1, 0, `Known deceased`)
  )


fit <- survfit(Surv(`Known Survival Length`, `Known deceased`) ~ Surgery, data = c12_censored_df)#c12_df)
print(fit)
ggsurvplot(fit,
           pval = TRUE, conf.int = FALSE,
           ggtheme = theme_classic() + theme(plot.title = element_text(hjust = 0.5)), # customize plot and risk table with a theme.
           
           title = "Discharged Alive: Surgery vs No Surgery", 
           risk.table = "abs_pct", # Add risk table
           risk.table.col = "strata", # Change risk table color by groups
           break.time.by = 3,     # break X axis in time intervals by 200.
           risk.table.y.text.col = T,
           xlab = "Time (months)",
           linetype = c(1, 1), # Change line type by groups
           legend.labs = c("No Surgery", "Surgery"),
           palette = c("Orange", "#2E9FDF"))


fit <- survfit(Surv(`Known Survival Length`, `Known deceased`) ~ Female, data = c12_censored_df)#c12_df)
print(fit)
ggsurvplot(fit,
           pval = TRUE, conf.int = FALSE,
           ggtheme = theme_classic() + theme(plot.title = element_text(hjust = 0.5)), # customize plot and risk table with a theme.
           
           title = "Discharged Alive - Survival: Male vs Female", 
           risk.table = "abs_pct", # Add risk table
           risk.table.col = "strata", # Change risk table color by groups
           break.time.by = 3,     # break X axis in time intervals by 200.
           risk.table.y.text.col = T,
           xlab = "Time (months)",
           linetype = c(1, 1), # Change line type by groups
           legend.labs = c("Male", "Female"),
           palette = c("Orange", "#2E9FDF"))



fit2 <- survfit(Surv(`Known Survival Length`, `Known deceased`) ~ Binary_Injury_Severity, data = c12_censored_df)#c12_df)
print(fit2)
ggsurvplot(fit2,
           pval = TRUE, conf.int = FALSE,
           ggtheme = theme_classic() + theme(plot.title = element_text(hjust = 0.5)), # customize plot and risk table with a theme.
           
           title = "ISS vs Survival", 
           risk.table = "abs_pct", # Add risk table
           risk.table.col = "strata", # Change risk table color by groups
           break.time.by = 3,     # break X axis in time intervals by 200.
           risk.table.y.text.col = T,
           xlab = "Time (months)",
           linetype = c(1, 1), # Change line type by groups
           legend.labs = c("Below Average ISS", "Above Average"),
           palette = c("Orange", "#2E9FDF"))

fit2 <- survfit(Surv(`Known Survival Length`, `Known deceased`) ~ Neurodeficits , data = c12_censored_df)#c12_df)
print(fit2)
ggsurvplot(fit2,
           pval = TRUE, conf.int = FALSE,
           ggtheme = theme_classic() + theme(plot.title = element_text(hjust = 0.5)), # customize plot and risk table with a theme.
           
           title = "Discharged Alive - Survival vs Neurodeficits", 
           risk.table = "abs_pct", # Add risk table
           risk.table.col = "strata", # Change risk table color by groups
           break.time.by = 3,     # break X axis in time intervals by 200.
           risk.table.y.text.col = T,
           xlab = "Time (months)",
           linetype = c(1, 1), # Change line type by groups
           legend.labs = c("No Neurodeficits", "Neurodeficits"),
           palette = c("Orange", "#2E9FDF"))



res.cox1 <- coxph(Surv(`Known Survival Length`, `Known deceased`) ~ Female + `Age of Injury`,
                  data =  c12_df)
summary(res.cox1)


##########            ##########
########## ALL CASES  ##########
##########            ##########
# name = "C12_All_Cases"
# ### Predictors of Death
# death_formula <- as.formula("Expired.during.Admission ~ Female + Injury.Severity.Score + Neurodeficits + C2.Dens.Anderson.and.D.Alonzo.Type.3")
# model_Death <- glm(death_formula, data = c12_df, na.action = na.omit, family = 'binomial')
# summary(model_Death)
# death_predictors_df <- summarize_Multi_regression_model(model_Death)
# death_predictors_df
# 
# ### Predictors of Surgery
# surg_formula <- as.formula("Surgery ~  Female + HTN. + DM. + Injury.Severity.Score + C1.Landells.Type.I + C1.Landells.Type.2 + C2.Dens.Anderson.and.D.Alonzo.Type.2 + C2.Dens.Anderson.and.D.Alonzo.Type.3")
# model_Surg <- glm(surg_formula, data = c12_df, na.action = na.omit, family = 'binomial')
# summary(model_Surg)
# surg_predictors_df <- summarize_Multi_regression_model(model_Surg)
# surg_predictors_df
# 
# 
# ### Predictors of Union
# nonUnion_formula <- as.formula("NonUnion ~ Injury.Severity.Score + Occipital.Condyle.Fracture + C2.lateral.mass.fracture")
# nrow(c12_df)
# non_union_df <- c12_df[c12_df$Expired.during.Admission != 1, ]
# non_union_df <- non_union_df[!is.na(non_union_df$NonUnion), ]
# nrow(non_union_df)
# 
# model_NonUnion <- glm(nonUnion_formula, data = non_union_df, na.action = na.omit, family = 'binomial')
# summary(model_NonUnion)
# NonUnion_predictors_df <- summarize_Multi_regression_model(model_NonUnion)
# NonUnion_predictors_df
# 
# ##
# Multi_dfs <- list("Expired InPatient (Multi Log)" = death_predictors_df,
#                   "Surgery Analysis (Multi Log)" = surg_predictors_df,
#                   "NonUnion Analysis (Multi Log)" = NonUnion_predictors_df)
# 
# export_dfs_to_excel(Multi_dfs, name)
