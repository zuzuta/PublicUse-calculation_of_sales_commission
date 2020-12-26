### Summary

For public-use code of the automated calculation of sales commission in the company working now (at 2020/12/15).

Some proper norns are converted to anonymous.


### Description

1. Download shipping record file from the in-house database by .xlsx
2. Pass that file, master file and prepared sales-commission-database by .xlsx as command line arguments to ***"Original_to_Master.py"***
    - Copy the shipping record to master file
3. Input the payment receipt date on master file
4. Pass the updated master file and sales-commission-database as command line arguments to ***"Payment_Calculation.py"***
    - Copy the payment receipt date to master file, and calculate the commission payment schedule
5. Pass the master file, exceptional-commission-database and inputdate as command line arguments to ***"Exceptional_Commission_01.py"***
    - Calculate and input the exceptional sales commission on master file
6. Pass the master file, distributor's report and inputdate as command line arguments to ***"Exceptional_Commission_02.py"***
    - Calculate and input the exceptional sales commission on master file
7. Pass the master file as command line arguments to ***"Pivot_Table.py"***
    - Create pivot table, monthly commission payment amount by each sales rep.


### Improvement Plan

Automatic processing for inputting payment receopt date on master file insted of manual input.
