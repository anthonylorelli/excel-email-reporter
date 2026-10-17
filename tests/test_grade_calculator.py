"""Unit tests for GradeCalculator and student report generation."""

import pandas as pd
from src.grade_calculator import GradeCalculator, calculate_letter_grade, DEFAULT_CATEGORY_WEIGHTS


def test_calculate_letter_grade():
    assert calculate_letter_grade(98.0) == "A+"
    assert calculate_letter_grade(97.5) == "A+"
    assert calculate_letter_grade(97.4) == "A"
    assert calculate_letter_grade(94.0) == "A"
    assert calculate_letter_grade(92.5) == "A"
    assert calculate_letter_grade(90.0) == "A-"
    assert calculate_letter_grade(89.5) == "A-"
    assert calculate_letter_grade(88.0) == "B+"
    assert calculate_letter_grade(87.5) == "B+"
    assert calculate_letter_grade(87.4) == "B"
    assert calculate_letter_grade(84.0) == "B"
    assert calculate_letter_grade(82.5) == "B"
    assert calculate_letter_grade(80.0) == "B-"
    assert calculate_letter_grade(79.5) == "B-"
    assert calculate_letter_grade(78.0) == "C+"
    assert calculate_letter_grade(77.5) == "C+"
    assert calculate_letter_grade(75.0) == "C"
    assert calculate_letter_grade(72.5) == "C"
    assert calculate_letter_grade(70.0) == "C-"
    assert calculate_letter_grade(69.5) == "C-"
    assert calculate_letter_grade(68.0) == "D+"
    assert calculate_letter_grade(67.5) == "D+"
    assert calculate_letter_grade(65.0) == "D"
    assert calculate_letter_grade(62.5) == "D"
    assert calculate_letter_grade(62.0) == "D-"
    assert calculate_letter_grade(59.5) == "D-"
    assert calculate_letter_grade(59.4) == "F"
    assert calculate_letter_grade(50.0) == "F"


def test_grade_calculator_weighted_categories_and_missing_excluded():
    # User requirements:
    # 1. Homework: 35%, Essays: 35%, Teams Discussion: 20%, Attendance/Participation: 10%
    # 2. Missing work excluded until graded
    weights = {
        "Homework": 0.35,
        "Essays": 0.35,
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
    # Normalized: 0.63 / 0.65 = 96.923% -> 96.9% (A)
    expected_pct = round((0.63 / 0.65) * 100.0, 1)
    assert p.overall_percent == expected_pct
    assert p.letter_grade == "A"
    assert p.categories["Essays"].points_possible == 0.0
    assert p.categories["Essays"].score_percent is None  # Not yet assessed


def test_grade_calculator_with_graded_essays():
    weights = {
        "Homework": 0.35,
        "Essays": 0.35,
        "Teams Discussion": 0.20,
        "Attendance/Participation": 0.10,
    }
    calculator = GradeCalculator(weights)

    data = [
        {"Full Name": "John Smith", "First Name": "John", "Last Name": "Smith", "Email Address": "john@example.com",
         "Assignments": "Essay 1", "Grading Category": "Essays (35%)", "Points": 90.0, "Max Points": 100.0, "Status": "Returned"},
        {"Full Name": "John Smith", "First Name": "John", "Last Name": "Smith", "Email Address": "john@example.com",
         "Assignments": "HW 1", "Grading Category": "Homework (35%)", "Points": 10.0, "Max Points": 10.0, "Status": "Returned"},
        {"Full Name": "John Smith", "First Name": "John", "Last Name": "Smith", "Email Address": "john@example.com",
         "Assignments": "Disc 1", "Grading Category": "Teams Discussion (20%)", "Points": 10.0, "Max Points": 10.0, "Status": "Returned"},
        {"Full Name": "John Smith", "First Name": "John", "Last Name": "Smith", "Email Address": "john@example.com",
         "Assignments": "Att 1", "Grading Category": "Attendance/Participation (10%)", "Points": 10.0, "Max Points": 10.0, "Status": "Returned"},
    ]

    df = pd.DataFrame(data)
    profiles = calculator.process_records(df, course_name="Lit 101")
    p = profiles[0]

    # All 4 categories active:
    # Essays: 90% * 0.35 = 31.5%
    # HW: 100% * 0.35 = 35.0%
    # Disc: 100% * 0.20 = 20.0%
    # Att: 100% * 0.10 = 10.0%
    # Overall = 96.5% -> A
    assert p.categories["Essays"].score_percent == 90.0
    assert p.overall_percent == 96.5
    assert p.letter_grade == "A"
