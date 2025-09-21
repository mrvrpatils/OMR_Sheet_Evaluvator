import streamlit as st
import cv2
import numpy as np
from PIL import Image
import io
import mysql.connector
from parse_keys import evaluate_and_process, save_results_to_db  # Import the main processing function

# Initialize session state for theme mode
if 'theme_mode' not in st.session_state:
    st.session_state['theme_mode'] = 'light'

# Custom CSS for styling
st.markdown(
    f"""
    <style>
    .reportview-container {{
        background: #f0f2f6;
    }}
   .sidebar .sidebar-content {{
        background: #e2eaf2;
    }}
    .css-12oz5g7 {{
        padding: 1rem 1rem 1.5rem;
    }}
    /* Add a modern color scheme */
    body {{
        color: #333;
        background-color: #f8f9fa;
    }}
    h1, h2, h3, h4, h5, h6 {{
        color: #007bff;
    }}
    .stButton>button {{
        color: #fff;
        background-color: #007bff;
        border: none;
        border-radius: 0.25rem;
        padding: 0.5rem 1rem;
    }}
    .stButton>button:hover {{
        background-color: #0056b3;
    }}
    /* Dark mode */
    [data-theme="dark"] {{
        body {{
            color: #eee;
            background-color: #111;
        }}
        h1, h2, h3, h4, h5, h6 {{
            color: #add8e6;
        }}
        .stButton>button {{
            color: #000;
            background-color: #add8e6;
        }}
        .stButton>button:hover {{
            background-color: #87ceeb;
        }}
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

def main():
    st.title("OMR Scanning and Result Processing")
    st.markdown("Upload an OMR sheet image or use your camera to capture one.")

    # Sidebar options
    with st.sidebar:
        st.header("Options")
        def set_theme():
            st.session_state['theme_mode'] = theme_mode.lower()

        theme_mode = st.radio("Select Theme:", ("Light", "Dark"), on_change=set_theme)

        image = None

        source = st.radio("Select Input Source:", ("Upload Image", "Camera"))

        if source == "Upload Image":
            uploaded_file = st.file_uploader("Upload OMR Sheet (JPG, JPEG)", type=["jpg", "jpeg"])
            if uploaded_file is not None:
                image = Image.open(uploaded_file)
        elif source == "Camera":
            captured_image = st.camera_input("Take a picture of the OMR sheet")
            if captured_image is not None:
                image_bytes = captured_image.read()
                image = Image.open(io.BytesIO(image_bytes))

    if image is not None:
        with st.expander("Uploaded Image and Processing Details", expanded=False):
            st.image(image, caption="Uploaded/Captured OMR Sheet", use_column_width=True)

            # Process the image
            try:
                # Convert PIL Image to OpenCV format
                img_array = np.array(image)
                img_cv2 = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)

                # Assuming evaluate_and_process takes image path and excel path
                # and returns a dictionary containing results
                # You might need to save the image to a temporary file
                # or modify evaluate_and_process to accept the image directly
                cv2.imwrite("temp_image.jpg", img_cv2)
                results = evaluate_and_process("temp_image.jpg", "answer_keys.xlsx")

                if st.button("Show Results"):
                    st.subheader("Results")
                    #st.write(f"Student Name: {results['student_name']}")
                    st.write(f"Set Name: {results['set_name']}")
                    for subject, score in results['subject_scores'].items():
                        st.write(f"{subject}: {score}")
                    st.write(f"Total Score: {results['total_score']}")

                if st.button("Save Results to Database"):
                    save_results_to_db(results)
                    st.success("Results saved to database!")

            except Exception as e:
                st.error(f"Error processing image: {e}")

# Apply theme based on session state
    js = f"""
<script>
function add_data_attribute(element, attr, value) {{
  element.setAttribute(attr, value);
}}

const target = parent.document.querySelector('body');
add_data_attribute(target, 'data-theme', '{st.session_state["theme_mode"]}');
</script>
"""
    st.markdown(js, unsafe_allow_html=True)

if __name__ == "__main__":
    main()