"""Unit tests for GradeCalculator and student report generation."""

import pandas as pd
from src.grade_calculator import GradeCalculator, calculate_letter_grade, DEFAULT_CATEGORY_WEIGHTS


def test_calculate_letter_grade():
    assert calculate_letter_grade(98.0) == "A+"
    assert calculate_letter_grade(94.0) == "A"
    assert calculate_letter_grade(90.0) == "A-"
    assert calculate_letter_grade(88.0) == "B+"
    assert calculate_letter_grade(84.0) == "B"
    assert calculate_letter_grade(75.0) == "C"
    assert calculate_letter_grade(62.0) == "D"
    assert calculate_letter_grade(50.0) == "F"


def test_grade_calculator_weighted_categories_and_missing_excluded():
    # User requirements:
    # 1. Homework: 35%, Exams: 35%, Teams Discussion: 20%, Attendance/Participation: 10%
    # 2. Missing work excluded until graded
    weights = {
        "Homework": 0.35,
        "Exams": 0.35,
        "Teams Discussion": 0.20,
        "Attendance/Participation": 0.10,
    }

    calculator = GradeCalculator(weights)

    data = [
        # Student 1: Jane Doe
        # Homework: 10/10 (100%)
        {"Full Name": "Jane Doe", "First Name": "Jane", "Last Name": "Doe", "Email Address": "jane@example.com",
         "Assignments": "HW 1", "Grading Category": "Homework (35%)", "Points": 10.0, "Max Points": 10.0, "Status": "Returned"},
        # Discussion: 9/10 (90%)
        {"Full Name": "Jane Doe", "First Name": "Jane", "Last Name": "Doe", "Email Address": "jane@example.com",
         "Assignments": "Discussion 1", "Grading Category": "Teams Discussion (20%)", "Points": 9.0, "Max Points": 10.0, "Status": "Returned"},
        # Participation: 10/10 (100%)
        {"Full Name": "Jane Doe", "First Name": "Jane", "Last Name": "Doe", "Email Address": "jane@example.com",
         "Assignments": "Bellwork 1", "Grading Category": "Attendance/Participation (10%)", "Points": 10.0, "Max Points": 10.0, "Status": "Returned"},
        # Missing assignment (Points is None) -> MUST BE EXCLUDED from score!
        {"Full Name": "Jane Doe", "First Name": "Jane", "Last Name": "Doe", "Email Address": "jane@example.com",
         "Assignments": "HW 2", "Grading Category": "Homework (35%)", "Points": None, "Max Points": 10.0, "Status": "Viewed"},
    ]

    df = pd.DataFrame(data)
    profiles = calculator.process_records(df, course_name="English 101")

    assert len(profiles) == 1
    p = profiles[0]
    assert p.full_name == "Jane Doe"
    assert p.total_graded == 3
    assert p.total_missing == 1

    # Active weights: HW (0.35), Discussion (0.20), Participation (0.10) = 0.65 total
    # Total points: 1.00*0.35 + 0.90*0.20 + 1.00*0.10 = 0.35 + 0.18 + 0.10 = 0.63
    # Normalized: 0.63 / 0.65 = 96.923% -> 96.9% (A+)
    expected_pct = round((0.63 / 0.65) * 100.0, 1)
    assert p.overall_percent == expected_pct
    assert p.letter_grade == "A+"
    assert p.categories["Exams"].points_possible == 0.0
    assert p.categories["Exams"].score_percent is None  # Not yet assessed
