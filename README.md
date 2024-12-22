# Data-Driven Academic Performance Analysis System

## Project Overview

The **Data-Driven Academic Performance Analysis System** is designed to automate the analysis of student exam performance in a school environment. By utilizing Excel formulas and Python scripts, the system calculates topic-wise, unit-wise, and subject-wise performance for each student. The results are then used to generate dynamic reports in DOCX format, which are further converted into PDFs for distribution. Additionally, Power BI dashboards are created using a 'SUMMARY' sheet, providing overall class performance insights.

## Tools & Technologies

- **Excel**: Used for data entry and performing all calculations using built-in formulas. Each student's data is structured across multiple sheets, and a 'SUMMARY' sheet aggregates class-wide performance.
- **Python**: For generating automated DOCX reports based on the analysis performed in Excel.
- **python-docx**: Python library to create DOCX reports.
- **docx2pdf**: For converting DOCX reports into PDFs.
- **Power BI**: For creating visual dashboards from the 'SUMMARY' sheet to display class-wide performance metrics.

## Project Workflow

1. **Data Entry in Excel**: 
   - Each student’s exam data is entered into individual sheets, with cells left blank for each question's score. 
   - The Excel workbook is pre-structured with formulas to automatically calculate topic-wise, unit-wise, and subject-wise performance, improvement rates, and overall performance metrics.
   
2. **Excel Calculations**:
   - Excel formulas are used to calculate performance based on the question marks entered.
   - A 'SUMMARY' sheet aggregates the class's overall performance, providing averages and comparative statistics for each subject and unit.

3. **Python Script for Report Generation**:
   - Python is used to extract data from the Excel workbook and generate a detailed DOCX report for each student.
   - The report includes performance insights, strength areas, weaknesses, and improvement rates based on the calculated data.

4. **DOCX to PDF Conversion**:
   - Once the DOCX reports are generated, another Python script uses **docx2pdf** to convert them into PDF format for easy distribution.

5. **Power BI Dashboards**:
   - The 'SUMMARY' sheet is used to create interactive Power BI dashboards that display overall class performance.
   - Visualizations include subject averages, performance trends, and comparative analysis to help educators and students track improvement areas.

## Setup & Usage

### Prerequisites

- **Excel**: Ensure the Excel workbook is structured as described, with formulas in place for each student and the 'SUMMARY' sheet for class-wide aggregation.
- **Python**: Install the necessary libraries by running:
  ```bash
  pip install python-docx docx2pdf pandas openpyxl

## File Descriptions

- **data/student_data.xlsx**: The main Excel workbook containing student data, with multiple sheets (one per student) and a 'SUMMARY' sheet for overall class performance.
- **scripts/generate_reports.py**: Python script that extracts data from the Excel workbook and generates DOCX reports with detailed performance insights for each student.
- **scripts/convert_to_pdf.py**: Python script to convert the generated DOCX reports into PDFs for distribution.
- **docs/reports/**: Folder containing generated reports for each student(all examinations).
- **docs/powerbi dashboard/Overall Performance Analysis.pbix**: Power BI file containing the interactive dashboards for class performance analysis.

