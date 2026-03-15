# OpticalTweezers-Analysis
## Discription
Performing statistical analysis on raw data from optical tweezers
## Code Components
- Statistical analysis requires two code files: **process_excel.m** and **merge_and_stats.m**.
- The former is used for preprocessing raw data, including zero-point alignment, sorting, and grouping by distance.
- The latter performs further statistical analysis on the preprocessed data, calculating the median, mean, and interquartile range of distances by distance groups, as well as the mean, standard deviation, and standard error of forces.
## Code Analysis Steps
The specific usage of the codes is as follows:
- Place the raw data into an Excel file (e.g., data.xlsx), with each dataset stored in a separate sheet.
-  Use the **process_excel.m** code for preprocessing to obtain the processed data (e.g., data_process.xlsx).
- Use the **merge_and_stats.m** code for statistical analysis to obtain the results (e.g., data_stats.xlsx).
