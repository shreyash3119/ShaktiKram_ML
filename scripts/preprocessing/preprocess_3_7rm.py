"""
ShaktiKram_ML — 3-7RM Strength Test
Component 1: Configuration and paths only.

This component does not read the workbook, create output folders,
or modify source data.
"""
from pathlib import Path


# 1. PATHS
# Script location:
# F:\Projects\ShaktiKram_ML\scripts\preprocessing\preprocess_3_7rm.py
# parents[0] = preprocessing; parents[1] = scripts; parents[2] = project root
PROJECT_ROOT = Path(__file__).resolve().parents[2]

INPUT_FILE = (
    PROJECT_ROOT / "datasets" / "progress_tracking" / "raw"
    / "3-7RM_Strength_Test" / "Data.xlsx"
)
OUTPUT_DIR = (
    PROJECT_ROOT / "datasets" / "progress_tracking" / "processed"
    / "3-7RM_Strength_Test"
)

PERFORMANCE_OUTPUT_FILE = "exercise_performance.csv"
PROFILE_OUTPUT_FILE = "participant_profile.csv"
REPORT_OUTPUT_FILE = "preprocessing_report.txt"
TRANSFORMATION_LOG_OUTPUT_FILE = "transformation_log.csv"


# 2. WORKSHEETS
PERFORMANCE_SHEET = "3-7RM"
PROFILE_SHEET = "General descriptives"


# 3. PARTICIPANT ROW RANGES (inclusive)
PERFORMANCE_FIRST_ROW = 3
PERFORMANCE_LAST_ROW = 30
PROFILE_FIRST_ROW = 3
PROFILE_LAST_ROW = 30


# 4. EXPECTED COUNTS
EXPECTED_PARTICIPANTS = 28
EXPECTED_EXERCISES = 15


# 5. PERFORMANCE WORKSHEET COLUMNS (1-based; Excel letters in comments)
PERFORMANCE_ID_COLUMN = 1       # A
EXERCISE_FIRST_COLUMN = 4       # D
EXERCISE_LAST_COLUMN = 33       # AG
EXERCISE_COLUMN_STEP = 2        # Weight/Reps pairs


# 6. PROFILE WORKSHEET COLUMNS (1-based; Excel letters in comments)
PROFILE_ID_COLUMN = 3                  # C
PROFILE_SEX_COLUMN = 4                 # D
PROFILE_AGE_COLUMN = 5                 # E
PROFILE_BODY_MASS_COLUMN = 6          # F
PROFILE_HEIGHT_COLUMN = 7             # G
PROFILE_BODY_FAT_PERCENT_COLUMN = 10  # J; column I deliberately excluded
PROFILE_TRAINING_EXPERIENCE_COLUMN = 11  # K
# Column L (training sessions/week) is deliberately excluded.


# 7. EXPECTED HEADER LABELS
EXPECTED_WEIGHT_HEADER = "Weight (kg)"
EXPECTED_REPS_HEADER = "Reps"


# 8. LABEL MAPS
EXERCISE_NAME_MAP = {
    "Tricpes extension": "Triceps extension",
}
SEX_CODE_MAP = {
    1: "Male",
    2: "Female",
}


# 9. MEASUREMENT VALIDATION POLICIES
ZERO_LOAD_POLICY = "retain_and_flag_for_review"
MIN_EXPECTED_REPS = 3
MAX_EXPECTED_REPS = 7
OUT_OF_RANGE_REPS_POLICY = "retain_and_flag"

MISSING_LOAD_POLICY = "exclude_and_log"
MISSING_REPS_POLICY = "exclude_and_log"
NEGATIVE_LOAD_POLICY = "exclude_and_log"
NON_WHOLE_REPS_POLICY = "exclude_and_log"
# Whole-number validation must accept 7 and 7.0:
# after numeric/finite checks, test value == int(value).


# 10. OUTPUT SCHEMAS
# qa_flag is empty for clean records. If flagged, use:
# ZERO_LOAD_REVIEW and/or REPS_OUTSIDE_3_7, joined with "|".
PERFORMANCE_OUTPUT_COLUMNS = [
    "subject_id", "exercise", "load_kg", "reps", "qa_flag"
]
PROFILE_OUTPUT_COLUMNS = [
    "subject_id", "sex_code", "sex", "age", "body_mass_kg",
    "height_m", "body_fat_pct", "training_experience_years"
]


def print_configuration_summary():
    """Print resolved paths and key configuration; create nothing."""
    print("=" * 65)
    print("SHAKTIKRAM ML — 3-7RM PREPROCESSING CONFIGURATION")
    print("=" * 65)
    print(f"Project root:       {PROJECT_ROOT}")
    print(f"Input workbook:     {INPUT_FILE}")
    print(f"Input exists:       {INPUT_FILE.exists()}")
    print(f"Output directory:   {OUTPUT_DIR}")
    print("Output directory created by this script: No")
    print(f"Worksheets:         {PERFORMANCE_SHEET!r}, {PROFILE_SHEET!r}")
    print(f"Expected people:    {EXPECTED_PARTICIPANTS}")
    print(f"Expected exercises: {EXPECTED_EXERCISES}")
    print(
        "Performance rows:  "
        f"{PERFORMANCE_FIRST_ROW}-{PERFORMANCE_LAST_ROW}"
    )
    print(f"Profile rows:       {PROFILE_FIRST_ROW}-{PROFILE_LAST_ROW}")
    print(f"Performance cols:   {PERFORMANCE_OUTPUT_COLUMNS}")
    print(f"Profile cols:       {PROFILE_OUTPUT_COLUMNS}")
    print(f"Zero-load policy:   {ZERO_LOAD_POLICY}")
    print(
        "Expected reps:      "
        f"{MIN_EXPECTED_REPS}-{MAX_EXPECTED_REPS} "
        f"({OUT_OF_RANGE_REPS_POLICY})"
    )
    print("Output filenames:")
    for filename in (
        PERFORMANCE_OUTPUT_FILE,
        PROFILE_OUTPUT_FILE,
        REPORT_OUTPUT_FILE,
        TRANSFORMATION_LOG_OUTPUT_FILE,
    ):
        print(f"  - {filename}")
    print("=" * 65)


if __name__ == "__main__":
    print_configuration_summary()
