# CO-PO Attainment Analysis System

Automated CO-PO Attainment Analysis System using Flask and Python for Outcome-Based Education (OBE).

------------------------------------------------------------

## Project Overview

Outcome-Based Education (OBE) requires systematic evaluation of Course Outcomes (CO) and Program Outcomes (PO).
Manual CO-PO attainment calculations are time-consuming and prone to errors.

This project provides a web-based solution to automate:
- Direct CO Attainment
- Indirect CO Attainment
- Final CO Attainment
- PO and PSO Attainment

The system generates tables, charts, and downloadable Excel reports for academic analysis and accreditation.

------------------------------------------------------------

## Objectives

- Automate CO-PO attainment calculations
- Reduce manual errors in OBE evaluation
- Support configurable thresholds and weightages
- Generate Excel reports with charts
- Provide a simple web interface

------------------------------------------------------------

## Features

- Upload Excel input files
- Configurable attainment thresholds
- Adjustable CIE and SEE weightages
- Direct and indirect attainment calculation
- Automatic Excel report generation
- Error handling and input validation

------------------------------------------------------------

## Attainment Methodology

Direct CO Attainment:
Calculated using Continuous Internal Evaluation (CIE) and Semester End Examination (SEE).
Default weightage:
- CIE: 60 percent
- SEE: 40 percent

Indirect CO Attainment:
Calculated using student feedback survey ratings on a scale of 1 to 3.

Final CO Attainment:
Computed using:
- Direct CO Attainment: 80 percent
- Indirect CO Attainment: 20 percent

PO and PSO Attainment:
Calculated using CO-PO and CO-PSO mapping matrix and final CO values.

------------------------------------------------------------

## Technology Stack

Backend: Python, Flask  
Frontend: HTML, Bootstrap, JavaScript  
Data Processing: Pandas, NumPy  
Excel Handling: OpenPyXL  

------------------------------------------------------------

## Project Structure

co-po-attainment-analysis-system
|
|-- app.py
|-- requirements.txt
|-- README.md
|-- templates
|   |-- index.html
|   |-- results.html
|   |-- error.html
|-- static

------------------------------------------------------------

## Input Files

Student Data File:
- Sheet name: 1_Student_Marks
- Sheet name: 2_Tool_CO_Mapping

CO-PO Mapping File:
- CO to PO and PSO mapping values from 0 to 3

Survey File:
- Student feedback ratings for each CO

------------------------------------------------------------

## Output

- Direct CO attainment table
- Indirect CO attainment table
- Final CO attainment table
- PO and PSO attainment table
- Downloadable Excel report

------------------------------------------------------------

## How to Run the Project

Install dependencies:
pip install flask pandas numpy openpyxl

Run the application:
python app.py

Open browser and visit:
http://127.0.0.1:5000/

------------------------------------------------------------

## Applications

- Engineering colleges
- Outcome-Based Education analysis
- NBA and NAAC accreditation
- Academic performance evaluation

------------------------------------------------------------

## Future Enhancements

- User authentication
- PDF report generation
- Cloud deployment
- Historical data analysis

------------------------------------------------------------

## Project Type

Final Year Engineering Project  
Domain: Outcome-Based Education and Web Application
