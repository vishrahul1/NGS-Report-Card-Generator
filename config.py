import os

# Paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
CSV_PATH = os.path.join(DATA_DIR, "student_exam_data.csv")

# Excel sheet names
SCORE_SHEET_NAME = "Score_Sheet"
CLASS_WISE_SHEET_NAME = "Class_Wise"
EXAM_NAME_CELL = "A4"  # Cell in Class_Wise sheet containing exam name

# Columns to extract from Score_Sheet (first occurrence of duplicate headers)
REQUIRED_COLUMNS = [
    "Username",
    "First Name",
    "Batch",
    "Physics Score",
    "Chemistry Score",
    "Mathematics Score",
]

# CSV output columns
CSV_COLUMNS = REQUIRED_COLUMNS + ["Exam"]

# Academic year shown on report card
ACADEMIC_YEAR = "2025-26"

# Chart color scheme (matching template blue/orange/grey)
CHART_COLORS = {
    "series1": "#4472C4",  # Blue  – CDF
    "series2": "#ED7D31",  # Orange – JEE
    "series3": "#A5A5A5",  # Grey  – Class average (reserved)
}

# Subjects (order matters – matches template columns)
SUBJECTS = ["Mathematics", "Physics", "Chemistry"]

# Max marks per subject per exam type
CDF_MAX_PER_SUBJECT = 30
CDF_MAX_TOTAL = 90
JEE_MAX_PER_SUBJECT = 80
JEE_MAX_TOTAL = 240

# How to derive CDF score from the raw Excel score (which is JEE-scale out of 80).
# Set to True  → CDF = round(jee_score * CDF_MAX_PER_SUBJECT / JEE_MAX_PER_SUBJECT)
# Set to False → CDF column is left blank (fill in manually or via future mapping)
AUTO_SCALE_CDF = True

# Number of test rows in the template (Test-1 … Test-14)
MAX_TESTS = 11
