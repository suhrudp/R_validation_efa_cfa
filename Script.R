# Set the working directory to the specified path
setwd("~/Directory")

# Load necessary libraries for data manipulation, visualization, and analysis
library(readxl)
library(tidyverse)
library(ggsci)
library(ggthemes)
library(gtsummary)
library(gt)
library(flextable)
library(pROC)
library(psych)
library(Hmisc)
library(boot)
library(GPArotation)
library(lavaan)

# Read the third sheet from the 'Data.xlsx' Excel file into a dataframe 'dfsdq'
dfsdq <- read_xlsx("Data.xlsx", sheet = 3)
# Read the first sheet from the 'Data.xlsx' Excel file into a dataframe 'dfsvm'
dfsvm <- read_xlsx("Data.xlsx", sheet = 1)

# Combine the two dataframes 'dfsdq' and 'dfsvm' by columns
df <- cbind(dfsdq, dfsvm)

# Create a new variable 'SDQ_Scale_Cat' that is 1 if 'SDQ_Scale' is "At Risk", otherwise 0
# Convert 'SDQ_Scale_Cat' to a factor
df$`SDQ_Scale_Cat` <- ifelse(df$`SDQ_Scale` == "At Risk", 1, 0)
df$`SDQ_Scale_Cat` <- as_factor(df$`SDQ_Scale_Cat`)

# Print the names of the columns in 'df'
colnames(df)

# Create a summary table of selected columns and assign it to 'table1'
# Convert 'table1' to a flextable object and save it as a Word document
table1 <- df[,c(3,5,6,33,34,58)] %>%
  tbl_summary()
table1 %>% as_flex_table() %>% save_as_docx(path = "Table 1.docx")

# Perform ROC analysis using 'SDQ_Scale_Cat' as the response and 'Total_Score' as the predictor
roc <- roc(response = df$`SDQ_Scale_Cat`, predictor = df$`Total_Score`)

# Calculate the best threshold and associated statistics using the Youden index
params <- coords(roc, "best", best.method = "youden",
                 ret = c("threshold", "accuracy", "sensitivity", "specificity", "npv", "ppv"))
# Calculate confidence intervals for the best threshold and associated statistics
paramsci <- ci.coords(roc, "best", best.policy = "random", best.method = "youden",
                      ret = c("threshold", "accuracy", "sensitivity", "specificity", "npv", "ppv"))
# Redirect output to 'results.txt', print the summary table, ROC analysis, and calculated parameters
sink(file = "results.txt")
print(table1)
print(roc)
print(params)
print(paramsci)
# Redirect output back to the console
sink()

# Create a ROC plot and save it as 'Plot 1.png'
plot1 <- ggroc(roc, 
               lwd = 1.25) + 
  geom_abline(intercept = 1) + 
  labs(x = "Specificity", 
       y = "Sensitivity") +
  theme_minimal()
ggsave(plot1, 
       filename = "Plot 1.png", 
       dpi = 600, 
       height = 8, 
       width = 8,
       bg = "white")

# Prepare data for exploratory factor analysis (EFA) by selecting specific columns from 'dfsvm'
dfsvm <- dfsvm[, c(8:22)]
# Conduct EFA with 5 factors, oblimin rotation, and weighted least squares method
efa_result <- fa(r = cor(dfsvm),
                 nfactors = 5,
                 rotate = "oblimin",
                 fm = "wls")
# Extract eigenvalues and create a data frame for plotting
eigenvalues <- efa_result$values
factor_data <- data.frame(FactorNumber = 1:length(eigenvalues), Eigenvalue = eigenvalues)
# Create and save a scree plot as 'Plot 2.png'
plot2 <- ggplot(factor_data, aes(x = FactorNumber, y = Eigenvalue)) +
  geom_point() +
  geom_line() +
  theme_minimal() +
  scale_x_continuous(breaks = seq(0,15,1)) +
  geom_hline(yintercept = 1, linetype = "dashed") +
  labs(title = "Scree Plot", x = "Factor Number", y = "Eigenvalue")
ggsave(plot2, filename = "Plot 2.png", dpi = 600, width = 12, height = 8, bg = "white")
# Print the factor loadings and the complete EFA result
print(efa_result$loadings)
print(efa_result)

# Define a Confirmatory Factor Analysis (CFA) model
cfa_model <- '
  # Define the latent variables using the updated column names
  WLS2 =~ I_feel_like_hurting_or_harming_myself + I_feel_sad_a_lot + I_feel_it_would_be_very_good_if_I_was_not_there_at_all
  WLS5 =~ I_understand_how_other_people_feel_ + I_listen_to_others_and_do_what_they_say + I_help_others
  WLS1 =~ I_hit_someone_or_yell_at_someone_when_I_get_angry + Other_children_make_fun_of_me_ + 
          I_tease_others_make_fun_of_others_and_say_mean_things_to_others
  WLS3 =~ I_like_to_be_alone_and_I_like_to_play_by_myself
  WLS4 =~ I_have_good_friends
'
# Perform CFA with the specified model and data, then print the summary with fit measures and standardized results
cfa_result <- cfa(cfa_model, data = dfsvm)
summary(cfa_result, fit.measures = TRUE, standardized = TRUE)

# Calculate Cronbach's alpha for all items in dfsvm
cronbach_alpha_result <- alpha(dfsvm)

# Print the result
print(cronbach_alpha_result)
