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
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

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
    text = pytesseract.image_to_string(roi_image, config='--psm 6').strip()
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
def extract_student_answers(image_path):
    image = cv2.imread(image_path)
    if image is None:
        raise FileNotFoundError(f"Could not read the image file at: {image_path}")
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    
    # Step 1: Pre-processing - Add blur to reduce noise
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    
    # Step 2: Thresholding - Use more common parameters
    thresh = cv2.adaptiveThreshold(blurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                                   cv2.THRESH_BINARY_INV, 11, 2)

    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    bubble_contours = []
    for cnt in contours:
        (x, y, w, h) = cv2.boundingRect(cnt)
        aspect_ratio = w / float(h)
        # Step 3: Contour Filtering - Loosen constraints
        if 20 <= w <= 50 and 20 <= h <= 50 and 0.7 <= aspect_ratio <= 1.3:
            bubble_contours.append(cnt)

    if not bubble_contours:
        return {}

    # Sort contours top-to-bottom
    bubble_contours.sort(key=lambda c: cv2.boundingRect(c)[1])
    
    # Step 4: Row Grouping - More robust row grouping
    question_rows = []
    if bubble_contours:
        current_row = [bubble_contours[0]]
        for i in range(1, len(bubble_contours)):
            first_bubble_in_row_y = cv2.boundingRect(current_row[0])[1]
            current_bubble_y = cv2.boundingRect(bubble_contours[i])[1]
            h = cv2.boundingRect(bubble_contours[i])[3]
            if abs(current_bubble_y - first_bubble_in_row_y) < h:
                current_row.append(bubble_contours[i])
            else:
                current_row.sort(key=lambda c: cv2.boundingRect(c)[0])
                question_rows.append(current_row)
                current_row = [bubble_contours[i]]
        current_row.sort(key=lambda c: cv2.boundingRect(c)[0])
        question_rows.append(current_row)

    student_answers = {}
    question_num = 1
    options_per_question = 4
    option_map = {0: 'A', 1: 'B', 2: 'C', 3: 'D'}
    
    all_questions_in_order = []
    if not question_rows:
        return {}

    # The original question reordering logic is kept, as it's likely correct for the assumed layout
    num_question_cols = 4 
    
    for col_idx in range(num_question_cols):
        for row in question_rows:
            # Check if the row has enough bubbles for the current column
            if len(row) >= (col_idx + 1) * options_per_question:
                question_bubbles = row[col_idx * options_per_question : (col_idx + 1) * options_per_question]
                all_questions_in_order.append(question_bubbles)

    for question_bubbles in all_questions_in_order:
        max_fill_ratio = 0
        marked_option_index = -1
        for j, cnt in enumerate(question_bubbles):
            mask = np.zeros(thresh.shape, dtype="uint8")
            cv2.drawContours(mask, [cnt], -1, 255, -1)
            fill = cv2.countNonZero(cv2.bitwise_and(thresh, thresh, mask=mask))
            
            area = cv2.contourArea(cnt)
            if area > 0:
                fill_ratio = fill / area
                if fill_ratio > max_fill_ratio:
                    max_fill_ratio = fill_ratio
                    marked_option_index = j
        
        # Step 5: Fill Ratio - Lower the threshold
        if max_fill_ratio > 0.4:
            student_answers[question_num] = option_map.get(marked_option_index)
        
        question_num += 1

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
        set_name = str(extract_text_from_roi(image_for_ocr, set_name_roi))

        if "set b" in set_name.lower():
            set_name = "Set B"
        else:
            set_name = "Set A"
        
        print(f"Processing OMR for Student: '{student_name}', Set: '{set_name}'")

        # 2. Load the correct answer key and subject list
        keys, subjects = load_answer_keys(excel_path, set_name)
        
        # 3. Extract student's marked answers
        student_answers = extract_student_answers(image_path)
        print(f"Detected {len(student_answers)} marked answers.")

        if not student_answers:
            print("\nError: Could not extract any student answers from the OMR sheet.")
            print("Please check the image quality and the bubble detection parameters.\n")
            return None

        # 4. Calculate scores
        total_score = 0
        subject_scores = {subject: 0 for subject in subjects}

        for q_num, correct_info in keys.items():
            student_ans = student_answers.get(q_num)
            correct_ans = correct_info['answer']
            subject = correct_info['subject']
            
            if student_ans is not None and student_ans.lower() in correct_ans.lower().split(','):
                total_score += 1
                if subject in subject_scores:
                    subject_scores[subject] += 1

        # 5. Prepare results and save to DB
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
