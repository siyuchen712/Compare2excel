# Compare2excel
Use python to quick compare 2 EXCEL file, which have several sheets. Generate a new table to highlight the difference of these 2 files.

## Notes:
**1. Because of the binary, we set the numbers which are in Compare Table and in the range of (-0.02 - 0.02) to 0.**

**2. There are difference between 'min cycle#' and 'max cycle#' might because Excel and Python will pick different one.**
   * For instance, a list which is [0, 0, 0]. Excel will say the first one is the max, and last one is the min. However, Python will say the last one is the min and max.
   
**3. The last sheet has the information of program version and input files (both manual and program).**
