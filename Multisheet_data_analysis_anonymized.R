# Multisheet two measurements timecourse analysis

# Packages
library(readxl)
library(dplyr)
library(purrr)
library(ggplot2)
library(lme4)
library(lmerTest)
library(stringr)

# Load data
input_file <- "data/input_data.xlsx"

# Get all sheet names
sheet_names <- excel_sheets(input_file)

# Read one sheet (timepoint)
read_timepoint_sheet <- function(sheet_name) {
  
  data_wide <- read_excel(input_file, sheet = sheet_name)
  
  # Select first four columns and assign neutral names
  data_wide <- data_wide[, 1:4]
  names(data_wide) <- c("Group", "Unit", "Measurement_A", "Measurement_B")
  
  data_wide %>%
    mutate(
      Timepoint = sheet_name,
      Group = str_squish(as.character(Group)),
      Unit = str_squish(as.character(Unit)),
      Measurement_A = as.numeric(gsub(",", ".", as.character(Measurement_A))),
      Measurement_B = as.numeric(gsub(",", ".", as.character(Measurement_B)))
    ) %>%
    filter(!is.na(Group), !is.na(Unit))
}

# Combine all sheets into one dataset
data <- map_dfr(sheet_names, read_timepoint_sheet)

# Clean and standardize timepoints
data <- data %>%
  mutate(
    Timepoint = str_extract(Timepoint, "\\d+"),
    Timepoint = paste0("T", Timepoint)
  )

# Anonymize group and unit labels
group_map <- setNames(
  paste0("Group_", seq_along(sort(unique(data$Group)))),
  sort(unique(data$Group))
)

unit_map <- setNames(
  paste0("Unit_", seq_along(sort(unique(data$Unit)))),
  sort(unique(data$Unit))
)

data <- data %>%
  mutate(
    Group = group_map[as.character(Group)],
    Unit = unit_map[as.character(Unit)]
  )

# Dynamic timepoint levels 
timepoint_levels <- paste0(
  "T",
  sort(unique(as.numeric(str_extract(as.character(data$Timepoint), "\\d+"))))
)

# Convert to factors
data <- data %>%
  mutate(
    Timepoint = factor(Timepoint, levels = timepoint_levels),
    Group = factor(Group),
    Unit = factor(Unit)
  )

# Data checks
print(data)
str(data)
unique(data$Group)
table(data$Group, useNA = "ifany")
table(data$Timepoint, useNA = "ifany")
sum(!is.na(data$Measurement_A))
sum(!is.na(data$Measurement_B))
sum(complete.cases(data[, c("Measurement_A", "Timepoint", "Group", "Unit")]))

# Summary statistics
summary_data <- data %>%
  group_by(Group, Timepoint) %>%
  summarise(
    n_a = sum(!is.na(Measurement_A)),
    mean_a = mean(Measurement_A, na.rm = TRUE),
    sd_a = sd(Measurement_A, na.rm = TRUE),
    
    n_b = sum(!is.na(Measurement_B)),
    mean_b = mean(Measurement_B, na.rm = TRUE),
    sd_b = sd(Measurement_B, na.rm = TRUE),
    .groups = "drop"
  )

print(summary_data)

# Plot Measurement A
plot_a <- ggplot(summary_data,
                 aes(x = Timepoint, y = mean_a, color = Group, group = Group)) +
  geom_line(linewidth = 1) +
  geom_point(size = 3) +
  geom_errorbar(aes(ymin = mean_a - sd_a,
                    ymax = mean_a + sd_a),
                width = 0.15, linewidth = 0.7) +
  theme_minimal(base_size = 13) +
  labs(
    title = "Measurement A over time",
    x = "Timepoint",
    y = "Mean measurement A ± SD",
    color = "Group"
  )

print(plot_a)

# Plot Measurement B
plot_b <- ggplot(summary_data,
                 aes(x = Timepoint, y = mean_b, color = Group, group = Group)) +
  geom_line(linewidth = 1) +
  geom_point(size = 3) +
  geom_errorbar(aes(ymin = mean_b - sd_b,
                    ymax = mean_b + sd_b),
                width = 0.15, linewidth = 0.7) +
  theme_minimal(base_size = 13) +
  labs(
    title = "Measurement B over time",
    x = "Timepoint",
    y = "Mean measurement B ± SD",
    color = "Group"
  )

print(plot_b)

# Plot with individual data points
plot_a_points <- ggplot(data,
                        aes(x = Timepoint, y = Measurement_A, color = Group, group = Group)) +
  geom_jitter(width = 0.08, height = 0, size = 2, alpha = 0.7) +
  stat_summary(fun = mean, geom = "line", linewidth = 1) +
  stat_summary(fun = mean, geom = "point", size = 3) +
  theme_minimal(base_size = 13) +
  labs(
    title = "Measurement A by timepoint",
    x = "Timepoint",
    y = "Measurement A",
    color = "Group"
  )

print(plot_a_points)

# Mixed-effects model for Measurement A
model_data_a <- data %>%
  filter(!is.na(Measurement_A), !is.na(Timepoint), !is.na(Group), !is.na(Unit))

model_a_mixed <- lmer(Measurement_A ~ Timepoint * Group + (1 | Unit), data = model_data_a)
summary(model_a_mixed)
anova(model_a_mixed)

# Mixed-effects model for Measurement B
model_data_b <- data %>%
  filter(!is.na(Measurement_B), !is.na(Timepoint), !is.na(Group), !is.na(Unit))

model_b_mixed <- lmer(Measurement_B ~ Timepoint * Group + (1 | Unit), data = model_data_b)
summary(model_b_mixed)
anova(model_b_mixed)

# Alternative ANOVA (if no repeated measures)
model_a_aov <- aov(Measurement_A ~ Timepoint * Group, data = model_data_a)
summary(model_a_aov)

model_b_aov <- aov(Measurement_B ~ Timepoint * Group, data = model_data_b)
summary(model_b_aov)

# Save outputs
write.csv(data, "processed_data.csv", row.names = FALSE)
write.csv(summary_data, "summary_data.csv", row.names = FALSE)

ggsave("plot_measurement_a.png", plot_a, width = 7, height = 5, dpi = 300)
ggsave("plot_measurement_b.png", plot_b, width = 7, height = 5, dpi = 300)

