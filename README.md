📄 Automated OMR Evaluation & Scoring System

An automated OMR (Optical Mark Recognition) evaluation system built for the Innomatics Hackathon.
This project replaces manual evaluation of OMR sheets with a fast, accurate, and scalable pipeline that can process thousands of sheets in minutes.

🚀 Features

📸 OMR Capture – Works with scanned images or photos taken from a mobile phone

🛠️ Preprocessing – Corrects skew, rotation, and lighting issues using OpenCV

🔎 Bubble Detection – Detects filled bubbles using thresholding & contour analysis

📊 Answer Key Matching – Reads keys directly from Excel files (supports multiple sets like Set A & Set B)

🎯 Evaluation Engine – Calculates subject-wise scores (20 per subject × 5 subjects) and total (100)

📂 Results Export – Generates per-student results in CSV/Excel format

🌐 Future Scope – Streamlit dashboard for evaluators to upload sheets and view reports

🛠️ Tech Stack

Python – Core logic

OpenCV – Image preprocessing, bubble detection

Pandas / NumPy – Excel integration & data handling

Matplotlib (optional) – Visualization & debugging overlays

Streamlit (planned) – Simple web interface for evaluators

📂 Project Structure
├── evaluate_omr.py      # Main script to run OMR evaluation  
├── answer_keys.xlsx     # Excel file containing subject-wise answer keys  
├── omr_sample.jpg       # Example OMR sheet image  
├── results/             # Folder for saving student reports  
└── README.md            # Documentation  

🔑 How It Works

Upload OMR sheet images (JPG/PNG)

Provide answer keys in Excel format

Columns: Subject names (Python, SQL, EDA, etc.)

Rows: Q.No – Answer (e.g., 1 - a, 21 - c)

Run the evaluation:

python evaluate_omr.py


System outputs:

Subject-wise scores

Total score out of 100

CSV file with student results

📸 Example Workflow

Input: omr_sample.jpg + answer_keys.xlsx

Output:

Student ID: sample_001
Python: 15/20
EDA: 18/20
SQL: 17/20
Power BI: 16/20
Statistics: 19/20
Total: 85/100

📌 Future Improvements

✅ Build a Streamlit/Flask web dashboard

✅ Add real-time analytics & visualization

✅ Improve detection with ML models for faint/ambiguous marks

✅ Multi-sheet batch processing pipeline

🙌 Acknowledgements

This project is part of the Innomatics Hackathon challenge.
Special thanks to mentors and peers for guidance and feedback.

⚡ Fast, Accurate, and Transparent OMR Evaluation for Scalable Education Assessments.
