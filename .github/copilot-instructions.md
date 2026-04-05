# Project Overview
This is the place where data needs to be transfered from source excel file to the destination excel file.

## Goal of transformation
- data should always be read from the source_format_with_data.xlsm file having multiple tabs, each tab has the attendance for each month for all the team members.
- actions will be specified later but actions will relate to copying data of specific month (each tab represents each month) and then based on the target formatting would have to generate an excel report with name opal_batch1_agile_oracles_attendance_<month_name>_<year> where month_name and year would be the current values of month and year.
- Target file will have the exact formatting as present in destination_format.xlsx and only the data will be copied from the source_format_with_data excel file.

## Folder Structure
- source_format_with_data.xlsm — format of the excel file with source data
- destination_format.xlsx — tergetted format of the excel file where data from source is placed
- README.md - Contains details about content formatting for source and destination formatting
