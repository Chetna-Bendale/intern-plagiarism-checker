import os
import zipfile
import shutil
import time
from flask import Flask, render_template, request
from googleapiclient.discovery import build
import fitz  # PyMuPDF
from docx import Document
from config import GOOGLE_API_KEY, SEARCH_ENGINE_ID


# ============== 1. CONFIGURATION AND SETUP ==============
app = Flask(__name__)

# ---!!! IMPORTANT: PASTE YOUR GOOGLE API CREDENTIALS HERE !!!---
# It's better to use environment variables, but for simplicity, you can paste them here.


# Folder paths
UPLOAD_FOLDER = 'uploads'
EXTRACTION_FOLDER = 'extracted_files'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(EXTRACTION_FOLDER, exist_ok=True)


# ============== 2. SUBMISSION REQUIREMENTS DATABASE ==============
# This dictionary is created from the spreadsheet image provided.
SUBMISSION_REQUIREMENTS = {
    "Python": ["Project Charter", "Requirement Elicitation Questionnaire", "SRS", "WBS", "Project Schedule", "RAID Log", "Lessons Learnt Log", "Project Report", "Software Design Specification Template", "Old Problem Statement Video", "Domain Overview Video Missing"],
    "Data Analytics": ["Project Charter", "Requirement Elicitation Questionnaire", "SRS", "WBS", "Project Schedule", "RAID Log", "Lessons Learnt Log", "Project Report", "Old Problem Statement Video", "Domain Overview Video Missing"],
    "Data Management": ["Project Charter", "Requirement Elicitation Questionnaire", "SRS", "WBS", "Project Schedule", "RAID Log", "Lessons Learnt Log", "Project Report", "Old Problem Statement Video", "Problem Statement Video Missing"],
    "Business Analysis": ["Project Charter", "Requirement Elicitation Questionnaire", "SRS", "WBS", "Project Schedule", "RAID Log", "Lessons Learnt Log", "Project Report", "Tasks Specific Deliverables", "Domain Overview Video Missing", "Webinar Missing"],
    "Content Creation - Graphics & Multimedia": ["Tasks Specific Deliverables"],
    "Marketing": ["Project Charter", "Requirement Elicitation Questionnaire", "WBS", "Project", "RAID Log", "Lessons Le", "Project Report", "Tasks Specific Deliverables", "Problem Statement Video Missing", "Webinar Missing", "Domain Overview Video Missing"]
    # You can add all other domains from the sheet here in the same format.
}


# ============== 3. HELPER FUNCTIONS ==============

def cleanup_folders():
    """Deletes old files from previous runs."""
    for folder in [UPLOAD_FOLDER, EXTRACTION_FOLDER]:
        for filename in os.listdir(folder):
            file_path = os.path.join(folder, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f'Failed to delete {file_path}. Reason: {e}')

def validate_submitted_files(domain, submitted_filenames):
    """Checks submitted files against the required list for the domain."""
    if domain not in SUBMISSION_REQUIREMENTS:
        return {'error': 'Selected domain not found in requirements list.'}
    
    required = SUBMISSION_REQUIREMENTS[domain]
    # Normalize filenames by removing extensions and making lowercase for comparison
    submitted_normalized = [os.path.splitext(f)[0].lower() for f in submitted_filenames]
    required_normalized = [r.lower() for r in required]
    
    missing_files = [req_orig for req_orig, req_norm in zip(required, required_normalized) if req_norm not in submitted_normalized]
    
    return {
        'required_files': required,
        'submitted_files': submitted_filenames,
        'missing_files': missing_files,
        'all_submitted': len(missing_files) == 0
    }

def extract_text_from_file(file_path):
    """Reads and returns text from .pdf or .docx files."""
    text = ""
    try:
        if file_path.endswith('.pdf'):
            with fitz.open(file_path) as doc:
                for page in doc:
                    text += page.get_text()
        elif file_path.endswith('.docx'):
            doc = Document(file_path)
            for para in doc.paragraphs:
                text += para.text + "\n"
    except Exception as e:
        print(f"Could not read file {os.path.basename(file_path)}. Reason: {e}")
    # Split text into sentences for checking
    return [sentence.strip() for sentence in text.split('.') if len(sentence.strip()) > 10]

def check_plagiarism_with_google(text_chunks):
    """Searches Google for each text chunk to find matches."""
    if GOOGLE_API_KEY == "PASTE_YOUR_API_KEY_HERE" or SEARCH_ENGINE_ID == "PASTE_YOUR_SEARCH_ENGINE_ID_HERE":
        return {"error": "API Key or Search Engine ID not configured in app.py."}

    service = build("customsearch", "v1", developerKey=GOOGLE_API_KEY)
    found_matches = []

    for chunk in text_chunks:
        # Google search has a free daily quota, so we add a delay to be safe
        time.sleep(1) 
        try:
            query = f'"{chunk}"' # Search for the exact phrase
            res = service.cse().list(q=query, cx=SEARCH_ENGINE_ID).execute()
            if 'items' in res and len(res['items']) > 0:
                # If there's a result, the text was likely copied
                match = {
                    'text': chunk,
                    'source_url': res['items'][0]['link'],
                    'source_title': res['items'][0]['title']
                }
                found_matches.append(match)
                print(f"Found match for: '{chunk}' at {match['source_url']}")
        except Exception as e:
            print(f"Could not perform search for chunk. Reason: {e}")
            # This can happen if you exceed the API quota
            if "quota" in str(e).lower():
                return {"error": "Google API daily quota exceeded. Please try again tomorrow."}

    return found_matches


# ============== 4. WEB ROUTES ==============

@app.route('/')
def index():
    """Displays the main upload page with the domain dropdown."""
    cleanup_folders()
    domains = list(SUBMISSION_REQUIREMENTS.keys())
    return render_template('index.html', domains=domains)

@app.route('/process', methods=['POST'])
def process_submission():
    """Handles file upload, validation, and plagiarism checking."""
    # Get form data
    zip_file = request.files.get('zip_file')
    domain = request.form.get('domain')

    if not zip_file or not domain:
        return "Missing zip file or domain selection.", 400
    
    # Save and extract the zip file
    zip_path = os.path.join(UPLOAD_FOLDER, zip_file.filename)
    zip_file.save(zip_path)
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(EXTRACTION_FOLDER)

    submitted_filenames = os.listdir(EXTRACTION_FOLDER)

    # --- Task 1: Validate submitted files ---
    validation_results = validate_submitted_files(domain, submitted_filenames)

    # --- Task 2: Check for external plagiarism ---
    plagiarism_report = []
    for filename in submitted_filenames:
        if filename.endswith(('.docx', '.pdf')):
            file_path = os.path.join(EXTRACTION_FOLDER, filename)
            sentences = extract_text_from_file(file_path)
            matches = check_plagiarism_with_google(sentences)
            if matches: # Only add to report if plagiarism was found
                plagiarism_report.append({
                    'filename': filename,
                    'matches': matches
                })
    
    return render_template('results.html', 
                           domain=domain, 
                           validation=validation_results, 
                           plagiarism_report=plagiarism_report)


# ============== 5. RUN THE APPLICATION ==============

if __name__ == '__main__':
    app.run(debug=True)
