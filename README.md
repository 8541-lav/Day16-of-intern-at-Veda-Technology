VEDA TECHNOLOGY

DAY 16 – README

Data Analytics Internship | Day 16

Task: Simple Profit Analysis

Dataset: Global Superstore

Tool: Microsoft Excel

Worksheet: Profit_Analysis

Project Overview

This project compares profitability across the Furniture, Office Supplies, and Technology categories in the Global Superstore dataset. Excel was used to calculate total profit and average profit by category and to create a column chart for visual comparison.

Objectives

Calculate total profit for each product category.

Calculate average profit for each product category.

Compare category-level profitability.

Create a column chart showing total profit by category.

Dataset Fields Used

Category – column O; Profit – column U. The analysis uses the GlobalSuperstore worksheet.

Excel Formulas Used

Furniture total profit: =SUMIF(GlobalSuperstore!O:O,"Furniture",GlobalSuperstore!U:U)

Office Supplies total profit: =SUMIF(GlobalSuperstore!O:O,"Office Supplies",GlobalSuperstore!U:U)

Technology total profit: =SUMIF(GlobalSuperstore!O:O,"Technology",GlobalSuperstore!U:U)

Category average profit (example: Furniture): =AVERAGEIF(GlobalSuperstore!O:O,"Furniture",GlobalSuperstore!U:U)

Grand total profit: =SUM(B2:B4)

Overall average profit: =AVERAGE(GlobalSuperstore!U2:U51291)

Results

Category

Total Profit

Average Profit

Furniture

285,204.72

28.88

Office Supplies

518,473.83

16.58

Technology

663,778.73

65.45

Grand Total

1,467,457.29

28.61

Note: The Grand Total average is the overall average profit per dataset record, not the sum of the three category averages.

Key Insight

Technology generated the highest total profit (663,778.73) and the highest average profit per record (65.45). Office Supplies generated 518,473.83 in total profit but had the lowest average profit per record (16.58). Furniture generated 285,204.72 in total profit, with an average profit of 28.88.

Files to Include in Submission

GlobalSuperstore.xlsx – source workbook, if permitted by the internship.

Day16_Profit_Analysis_README.docx – project documentation.

Day16_Profit_Analysis_Project_Report.docx – project report.

Screenshot of the completed Profit_Analysis worksheet and chart.
