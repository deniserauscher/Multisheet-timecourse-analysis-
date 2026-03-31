# Multisheet Timecourse Analysis in R

## Overview
This project provides a generalized and anonymized workflow for analyzing multi-sheet Excel data in R, including:

* Data preprocessing and cleaning
* Merging multiple sheets into a single dataset
* Calculation of summary statistics
* Visualization of measurements over time
* Statistical modeling using mixed-effects models and ANOVA

The code is fully anonymized and can be reused for different datasets with similar structure.

## Features

* Import of multi-sheet Excel data
* Automatic merging and timepoint extraction
* Anonymization of group and unit identifiers
* Summary statistics (mean, standard deviation, sample size)
* Timecourse visualization using ggplot2
* Linear mixed-effects models (lme4, lmerTest)
* Optional two-way ANOVA
* Clean and reproducible analysis pipeline

## Input Data Format
The input file must be an Excel workbook with multiple sheets.
Each sheet represents a timepoint and must contain at least four columns in the following order:

| Column        | Description                   |
| ------------- | ----------------------------- |
| Group         | Categorical grouping variable |
| Unit          | Measurement unit / position   |
| Measurement_A | Numeric measurement           |
| Measurement_B | Numeric measurement           |

Column names are ignored during import — only column order is used.

## Installation
Make sure the following R packages are installed:
```r
install.packages(c(
  "readxl",
  "dplyr",
  "purrr",
  "ggplot2",
  "lme4",
  "lmerTest",
  "stringr"
))
```

## Usage
1. Place your dataset in the `data/` directory:
```
data/input_data.xlsx
```
2. Adjust the file path in the script if necessary:
```r
input_file <- "data/input_data.xlsx"
```
3. Run the script:
```r
source("multisheet_two_measurements_timecourse_analysis.R")
```

## Output
The script generates:
* `processed_data.csv`
  Cleaned and combined dataset
* `summary_data.csv`
  Aggregated statistics per group and timepoint
* `plot_measurement_a_over_time.png`
  Timecourse plot of Measurement A (mean ± SD)
* `plot_measurement_b_over_time.png`
  Timecourse plot of Measurement B (mean ± SD)

## Statistical Analysis
The workflow includes:
* Linear mixed-effects models: for repeated measurements within units
* Two-way ANOVA (alternative): for datasets without hierarchical or repeated structure

The models test:
* Effect of timepoint
* Effect of group
* Interaction between timepoint and group

## Data Privacy
This repository uses fully anonymized variable names and structure.
No specific, experimental, or sensitive identifiers data is included.
