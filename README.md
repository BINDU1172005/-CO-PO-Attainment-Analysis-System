📊 CO–PO Attainment Analysis System

Automated Outcome-Based Education Evaluation Tool

A web-based application that automates Course Outcome (CO) and Program Outcome (PO/PSO) attainment calculations as per Outcome-Based Education (OBE) standards.
The system processes academic data, applies configurable evaluation logic, and generates visual reports and Excel outputs for accreditation and academic analysis.

🎓 Project Type: Final Year Engineering Project
🏫 Domain: Outcome-Based Education (OBE), Data Analytics, Web Application
🛠️ Stack: Python, Flask, Pandas, OpenPyXL

🔍 About the Project

In Outcome-Based Education, evaluating CO and PO attainment manually is time-consuming, repetitive, and error-prone.
This project addresses the problem by providing a fully automated, configurable, and user-friendly system to calculate:

Direct CO Attainment (CIE + SEE)

Indirect CO Attainment (Student Surveys)

Final CO Attainment

Final PO / PSO Attainment

The system also generates professionally formatted Excel reports with charts, making it suitable for NBA / NAAC accreditation and internal academic reviews.

🎯 Objectives

Automate CO–PO attainment calculations

Minimize manual errors in OBE analysis

Support customizable thresholds and weightages

Provide visual and downloadable reports

Simplify data handling using Excel inputs

✨ Key Features

📥 Upload Excel files or use direct file links

⚙️ Configurable attainment thresholds (Level 3 / 2 / 1)

⚖️ Adjustable CIE–SEE and Direct–Indirect weights

📊 Automatic bar-chart generation

📄 Downloadable Excel report with multiple sheets

🧪 Sample template generator for easy input formatting

🚫 Robust error handling and validation

🧠 Attainment Methodology
1️⃣ Direct CO Attainment

Calculated using student performance in:

Continuous Internal Evaluation (CIE)

Semester End Examination (SEE)

Weighted average:

Default: 60% CIE + 40% SEE

2️⃣ Indirect CO Attainment

Derived from student feedback surveys using a 1–3 rating scale.

3️⃣ Final CO Attainment

Weighted combination of:

Direct CO Attainment (default: 80%)

Indirect CO Attainment (default: 20%)

4️⃣ PO / PSO Attainment

Computed using CO–PO / CO–PSO mapping matrix and final CO attainment values.

🧰 Technology Stack

Backend: Python, Flask

Frontend: HTML5, Bootstrap 5, JavaScript

Data Processing: Pandas, NumPy

Excel Automation: OpenPyXL

Visualization: Chart.js

📂 Project Structure


├── app.py
├── templates/
│   ├── index.html
│   ├── results.html
│   └── error.html
├── requirements.txt
└── README.md


📥 Input Files
Student Data File (Excel)

1_Student_Marks

2_Tool_CO_Mapping

CO–PO Mapping File

CO vs PO/PSO mapping (0–3 scale)

Survey File

Student ratings for each CO

📤 Output

Direct CO Attainment (Table + Chart)

Indirect CO Attainment (Table + Chart)

Final CO Attainment

Final PO / PSO Attainment

Downloadable Excel Report (.xlsx)

▶️ How to Run the Project

pip install flask pandas numpy openpyxl
python app.py

Open:

http://127.0.0.1:5000/

🎓 Applications

Engineering Colleges

NBA / NAAC Accreditation

Academic Outcome Analysis

Faculty Performance Review

🚀 Future Enhancements

Role-based authentication

PDF report generation

Cloud deployment

Historical attainment tracking

Dashboard analytics

🏁 Conclusion

The CO–PO Attainment Analysis System simplifies and standardizes OBE evaluation by automating complex calculations and report generation.
It provides a scalable, accurate, and efficient solution for academic institutions to analyze and improve learning outcomes.
