# 📊 CO–PO Attainment Analysis System

Automated CO–PO Attainment Analysis System using Flask and Python for Outcome-Based Education (OBE).

---

## 📌 Project Overview

Outcome-Based Education (OBE) requires systematic evaluation of Course Outcomes (CO) and Program Outcomes (PO).  
Manual CO–PO attainment calculations are time-consuming and prone to errors.

This project provides a web-based solution to automate:
- Direct CO Attainment
- Indirect CO Attainment
- Final CO Attainment
- PO / PSO Attainment

The system generates tabular results, charts, and downloadable Excel reports for academic analysis and accreditation.

---

## 🎯 Objectives

- Automate CO–PO attainment calculations  
- Reduce manual errors in OBE evaluation  
- Support configurable thresholds and weightages  
- Generate Excel reports with charts  
- Provide a user-friendly web interface  

---

## ✨ Key Features

- Upload Excel files or provide direct file links  
- Configurable attainment thresholds (Level 3 / 2 / 1)  
- Adjustable CIE–SEE and Direct–Indirect weights  
- Automatic chart generation  
- Downloadable Excel report (.xlsx)  
- Sample input template generator  
- Error handling and validation  

---

## 🧠 Attainment Methodology

### Direct CO Attainment
Calculated using student performance in:
- Continuous Internal Evaluation (CIE)
- Semester End Examination (SEE)

Default weightage:
- 60% CIE
- 40% SEE

---

### Indirect CO Attainment
Calculated using student feedback surveys on a 1–3 rating scale.

---

### Final CO Attainment
Computed as a weighted average of:
- Direct CO Attainment (80%)
- Indirect CO Attainment (20%)

---

### PO / PSO Attainment
Derived using the CO–PO / CO–PSO mapping matrix and final CO attainment values.

---

## 🧰 Technology Stack

- Backend: Python, Flask  
- Frontend: HTML, Bootstrap, JavaScript  
- Data Processing: Pandas, NumPy  
- Excel Automation: OpenPyXL  
- Visualization: Chart.js  

---

## 📂 Project Structure

