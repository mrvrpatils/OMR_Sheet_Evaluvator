import pandas as pd
import cv2
import numpy as np
import pytesseract
import mysql.connector
from mysql.connector import Error

# -------------------------
# CONFIGURATION
# -------------------------
# Path to Tesseract executable
#pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# MySQL Database Configuration
DB_CONFIG = {
    'host': 'localhost',
    'user': 'root',
    'password': 'mysql',
    'database': 'omr_results'
}

# -------------------------
# PART 1: OCR - Extract Text
# -------------------------
def extract_text_from_roi(image, roi_coords):
    """Extracts text from a given region of interest (ROI) in an image."""
    x, y, w, h = roi_coords
    roi_image = image[y:y+h, x:x+w]
    result = pytesseract.image_to_string(roi_image, config='--psm 6')
    if isinstance(result, str):
        text = result.strip()
    else:
        text = str(result)  # Convert dictionary to string
    return text

# -------------------------
# PART 2: Load Answer Keys by Set
# -------------------------
def load_answer_keys(excel_path, set_name):
    """Loads answer keys from a specific sheet (set) in the Excel file."""
    try:
        df = pd.read_excel(excel_path, sheet_name=set_name)
        print(f"Successfully loaded answer key from sheet: '{set_name}'")
    except ValueError:
        print(f"Warning: Worksheet named '{set_name}' not found. Falling back to the first sheet in the Excel file.")
        try:
            # Load the first sheet by its index (0)
            df = pd.read_excel(excel_path, sheet_name=0)
            print("Successfully loaded the first available sheet.")
        except Exception as e:
            raise ValueError(f"Could not find sheet '{set_name}' and also failed to load the first sheet. Error: {e}")

    answer_keys = {}
    subjects = df.columns
    for subject in subjects:
        for value in df[subject].dropna():
            s_value = str(value)
            if "-" in s_value:
                try:
                    q_num, ans = s_value.split("-", 1)
                    q_num = int(q_num.strip())
                    ans = ans.strip()
                    answer_keys[q_num] = {'answer': ans, 'subject': subject}
                except ValueError:
                    continue
    return answer_keys, subjects

# -------------------------
# PART 3: Extract Student Answers from OMR
# -------------------------
def extract_student_answers(image_path, subject_name_rois):
    student_answers = {}
    return student_answers

# -------------------------
# PART 4: Compare, Score, and Store
# -------------------------
def save_results_to_db(results):
    """Saves the final scores to the MySQL database."""
    conn = None
    cursor = None
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor()
        
        # Prepare the SQL query dynamically based on the subjects found
        # This now handles whitespace, case, and the specific "Satistics" typo
        subject_columns = [f"{sub.strip().lower().replace(' ', '_').replace('satistics', 'statistics')}_score" for sub in results['subject_scores']]
        columns = ['student_name', 'set_name', 'total_score'] + subject_columns
        placeholders = ['%s'] * len(columns)
        
        sql = f"INSERT INTO student_scores ({', '.join(columns)}) VALUES ({', '.join(placeholders)})"
        
        # Prepare the values to be inserted
        subject_values = list(results['subject_scores'].values())
        values = [results['student_name'], results['set_name'], results['total_score']] + subject_values
        
        cursor.execute(sql, tuple(values))
        conn.commit()
        print(f"Successfully saved results for {results['student_name']} to the database.")
        
    except Error as e:
        print(f"Error connecting to MySQL or saving data: {e}")
    finally:
        if conn and conn.is_connected():
            if cursor:
                cursor.close()
            conn.close()

def evaluate_and_process(image_path, excel_path):
    """Main function to orchestrate the entire OMR evaluation process."""
    try:
        # 1. Extract metadata using OCR
        # NOTE: These ROI coordinates are ESTIMATES and may need to be adjusted for your image.
        student_name_roi = (50, 50, 400, 100)
        set_name_roi = (600, 50, 200, 100)
        
        image_for_ocr = cv2.imread(image_path)
        student_name = str(extract_text_from_roi(image_for_ocr, student_name_roi))
        set_name = str(extract_text_from_roi(image_for_ocr, set_name_roi)).strip()

        if "set b" in set_name.lower() or "setb" in set_name.lower():
            set_name = "Set B"
        else:
            set_name = "Set A"
        
        print(f"Processing OMR for Student: '{student_name}', Set: '{set_name}'")

        # 2. Extract subject names from the image (assuming they are in a row at the top)
        subject_name_rois = [(110, 220, 80, 20), (260, 220, 100, 20), (410, 220, 60, 20), (560, 220, 60, 20), (710, 220, 70, 20)]  # Example ROIs
        extracted_subjects = []
        for roi in subject_name_rois:
            subject_name = extract_text_from_roi(image_for_ocr, roi).strip()
            extracted_subjects.append(subject_name)

        print(f"Extracted subject names: {extracted_subjects}")

        # 3. Load the correct answer key and subject list
        keys, subjects = load_answer_keys(excel_path, set_name)
        
        # 3. Extract student's marked answers
        student_answers = {}
        for i, subject_name in enumerate(extracted_subjects):
            # Calculate the ROI for the bubbles for this subject
            start_x = subject_name_rois[i][0]
            start_y = subject_name_rois[i][1] + 50  # Adjust this value based on the actual layout

            for question_num in range(1, 21):  # Assuming 20 questions per subject
                bubble_y = start_y + (question_num - 1) * 20  # Assuming 20 pixels vertical spacing

                # Define the ROIs for the four options (A, B, C, D)
                option_rois = [
                    (start_x, bubble_y, 15, 15),  # A
                    (start_x + 25, bubble_y, 15, 15),  # B
                    (start_x + 50, bubble_y, 15, 15),  # C
                    (start_x + 75, bubble_y, 15, 15)  # D
                ]

                # Check which option is marked
                marked_option = None
                for option_index, option_roi in enumerate(option_rois):
                    x, y, w, h = option_roi
                    if  image_for_ocr is not None and hasattr(image_for_ocr, 'shape') and image_for_ocr.shape is not None and x >= 0 and y >= 0 and x + w <= image_for_ocr.shape[1] and y + h <= image_for_ocr.shape[0]:
                        roi_image = image_for_ocr[y:y + h, x:x + w]
                        if roi_image is not None and roi_image.size > 0:
                            mean_intensity = cv2.mean(roi_image)
                            if isinstance(mean_intensity, tuple) and len(mean_intensity) > 0:
                                mean_intensity_value = mean_intensity[0]
                                # Adjust the threshold based on the image quality and lighting conditions
                                if mean_intensity_value < 100:
                                    marked_option = option_index
                                    break

                if marked_option is not None:
                    student_answers[f"{subject_name}_{question_num}"] = chr(65 + marked_option)  # A=65, B=66, etc.

        print(f"Detected student answers: {student_answers}")

        # 5. Calculate scores
        total_score = 0
        subject_scores = {subject: 0 for subject in subjects}

        for q_num, correct_info in keys.items():
            # Get the student's answer for this question
            student_ans = student_answers.get(f"{correct_info['subject']}_{q_num}")

            correct_ans = correct_info['answer']
            subject = correct_info['subject']
            
            if student_ans is not None and student_ans.lower() in correct_ans.lower().split(','):
                total_score += 1
                if subject in subject_scores:
                    subject_scores[subject] += 1

        # 6. Prepare results and save to DB
        results = {
            'student_name': student_name or "Unknown",
            'set_name': set_name,
            'total_score': total_score,
            'subject_scores': subject_scores
        }
        
        print("\n--- Evaluation Results ---")
        print(f"  Total Score: {total_score}")
        for subject, score in subject_scores.items():
            print(f"  {subject} Score: {score}")
        print("------------------------\n")
        
        save_results_to_db(results)
        return results
        
    except Exception as e:
        import traceback
        print(f"\n--- An unexpected error occurred ---")
        print(f"Error: {e}")
        print("Traceback:")
        traceback.print_exc()
        print("-------------------------------------\n")
        return None

if __name__ == "__main__":
    excel_file = "answer_keys.xlsx"
    omr_image = "omr_sample.jpg"

    evaluate_and_process(omr_image, excel_file)
