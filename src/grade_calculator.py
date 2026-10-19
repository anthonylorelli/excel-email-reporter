"""Grade calculation engine for student assignment data with weighted categories."""

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd


@dataclass
class AssignmentRecord:
    name: str
    category: str
    original_category: str
    due_date: Optional[str]
    status: str
    points: Optional[float]
    max_points: float
    grade: Optional[str]
    feedback: Optional[str]
    is_graded: bool
    is_missing: bool
    due_date_obj: Optional[datetime] = None


@dataclass
class CategorySummary:
    name: str
    weight: float
    points_earned: float = 0.0
    points_possible: float = 0.0
    graded_count: int = 0
    missing_count: int = 0
    score_percent: Optional[float] = None
    is_active: bool = False


@dataclass
class StudentGradeProfile:
    course_name: str
    full_name: str
    first_name: str
    last_name: str
    email: str
    categories: Dict[str, CategorySummary]
    overall_percent: float
    letter_grade: str
    graded_assignments: List[AssignmentRecord] = field(default_factory=list)
    missing_assignments: List[AssignmentRecord] = field(default_factory=list)
    total_graded: int = 0
    total_missing: int = 0
    parent_emails: List[str] = field(default_factory=list)


DEFAULT_CATEGORY_WEIGHTS: Dict[str, float] = {
    "Homework": 0.35,
    "Essays": 0.35,
    "Teams Discussion": 0.20,
    "Attendance/Participation": 0.10,
}


def normalize_category_name(cat_str: Optional[str], weights: Dict[str, float]) -> str:
    """Matches raw category text (e.g. 'Homework (35%)', 'Class Participation (10%)') to a standard category."""
    if not cat_str:
        return "Other"
    
    cat_lower = str(cat_str).lower().strip()
    
    # Priority matching
    if "essay" in cat_lower or "exam" in cat_lower or "test" in cat_lower:
        return "Essays"
    if "homework" in cat_lower or "assignment" in cat_lower:
        return "Homework"
    if "discussion" in cat_lower:
        return "Teams Discussion"
    if "participation" in cat_lower or "attendance" in cat_lower or "bellwork" in cat_lower:
        return "Attendance/Participation"
    
    # Direct match in configured weights
    for canonical in weights:
        if canonical.lower() in cat_lower:
            return canonical
            
    return cat_str.strip()


def calculate_letter_grade(percent: float) -> str:
    """Converts a percentage grade to standard letter grade."""
    if percent >= 97.5:
        return "A+"
    if percent >= 92.5:
        return "A"
    if percent >= 89.5:
        return "A-"
    if percent >= 87.5:
        return "B+"
    if percent >= 82.5:
        return "B"
    if percent >= 79.5:
        return "B-"
    if percent >= 77.5:
        return "C+"
    if percent >= 72.5:
        return "C"
    if percent >= 69.5:
        return "C-"
    if percent >= 67.5:
        return "D+"
    if percent >= 62.5:
        return "D"
    if percent >= 59.5:
        return "D-"
    return "F"


class GradeCalculator:
    """Calculates student grades from assignment records using weighted categories."""

    def __init__(self, category_weights: Optional[Dict[str, float]] = None):
        self.weights = category_weights or DEFAULT_CATEGORY_WEIGHTS.copy()

    def process_records(
        self,
        df: pd.DataFrame,
        course_name: str = "Class Report",
        parent_contacts: Optional[Dict[Tuple[str, str], List[str]]] = None,
    ) -> List[StudentGradeProfile]:
        """Processes tabular assignment data into a list of StudentGradeProfile objects.
        
        Rules applied:
        1. Categorized by: Homework (35%), Essays (35%), Teams Discussion (20%), Attendance/Participation (10%).
        2. Missing/ungraded work (Points is None / NaN) is EXCLUDED until graded.
        3. Overall percent is normalized by the sum of active category weights.
        """
        if df.empty:
            return []

        # Find required columns (case-insensitive fallback)
        cols = {c.lower().strip(): c for c in df.columns}
        
        name_col = cols.get("full name", cols.get("student name", "Full Name"))
        first_col = cols.get("first name", "First Name")
        last_col = cols.get("last name", "Last Name")
        email_col = cols.get("email address", cols.get("email", "Email Address"))
        assign_col = cols.get("assignments", cols.get("assignment", "Assignments"))
        due_col = cols.get("due date", "Due Date")
        status_col = cols.get("status", "Status")
        points_col = cols.get("points", "Points")
        max_col = cols.get("max points", cols.get("points possible", "Max Points"))
        grade_col = cols.get("grade", "Grade")
        feedback_col = cols.get("feedback", "Feedback")
        cat_col = cols.get("grading category", cols.get("category", "Grading Category"))

        profiles: List[StudentGradeProfile] = []

        # Group by student
        for student_name, student_df in df.groupby(name_col):
            if not student_name or str(student_name).strip() == "" or str(student_name) == "nan":
                continue

            first_row = student_df.iloc[0]
            first_name = str(first_row.get(first_col, "")) if first_col in first_row and pd.notna(first_row.get(first_col)) else ""
            last_name = str(first_row.get(last_col, "")) if last_col in first_row and pd.notna(first_row.get(last_col)) else ""
            email = str(first_row.get(email_col, "")) if email_col in first_row and pd.notna(first_row.get(email_col)) else ""

            full_str = str(student_name).strip()
            if (not first_name.strip() or not last_name.strip()) and " " in full_str:
                parts = full_str.split()
                if not first_name.strip():
                    first_name = parts[0]
                if not last_name.strip():
                    last_name = parts[-1]

            matched_parents: List[str] = []
            if parent_contacts:
                lookup_key = (first_name.strip().lower(), last_name.strip().lower())
                if lookup_key in parent_contacts:
                    matched_parents = list(parent_contacts[lookup_key])
                elif " " in full_str:
                    parts = full_str.split()
                    alt_key = (parts[0].strip().lower(), parts[-1].strip().lower())
                    if alt_key in parent_contacts:
                        matched_parents = list(parent_contacts[alt_key])

            # Initialize category summaries with configured weights
            cat_summaries: Dict[str, CategorySummary] = {
                cat: CategorySummary(name=cat, weight=w)
                for cat, w in self.weights.items()
            }

            graded_list: List[AssignmentRecord] = []
            missing_list: List[AssignmentRecord] = []

            for _, row in student_df.iterrows():
                assign_title = str(row.get(assign_col, "Untitled Assignment"))
                raw_cat = str(row.get(cat_col, "")) if cat_col in row and pd.notna(row[cat_col]) else ""
                cat_canonical = normalize_category_name(raw_cat, self.weights)
                
                # If category is not in standard list, dynamically register with 0 weight (or fallback)
                if cat_canonical not in cat_summaries:
                    cat_summaries[cat_canonical] = CategorySummary(name=cat_canonical, weight=0.0)

                status = str(row.get(status_col, "Unknown")) if status_col in row and pd.notna(row[status_col]) else "Unknown"
                feedback = str(row.get(feedback_col, "")) if feedback_col in row and pd.notna(row[feedback_col]) else ""
                grade_str = str(row.get(grade_col, "")) if grade_col in row and pd.notna(row[grade_col]) else ""

                # Format due date string and parse datetime object
                due_raw = row.get(due_col) if due_col in row else None
                due_str = ""
                due_dt: Optional[datetime] = None
                if pd.notna(due_raw):
                    try:
                        ts = pd.to_datetime(due_raw, errors="coerce")
                        if pd.notna(ts):
                            due_dt = ts.to_pydatetime()
                            if due_dt.tzinfo is not None:
                                due_dt = due_dt.replace(tzinfo=None)
                            due_str = due_dt.strftime("%b %d, %Y")
                        else:
                            due_str = str(due_raw).split(" ")[0]
                    except Exception:
                        due_str = str(due_raw).split(" ")[0]

                # Points handling
                raw_pts = row.get(points_col) if points_col in row else None
                raw_max = row.get(max_col) if max_col in row else None

                max_pts = 0.0
                if pd.notna(raw_max):
                    try:
                        max_pts = float(raw_max)
                    except (ValueError, TypeError):
                        max_pts = 0.0

                is_graded = False
                earned_pts: Optional[float] = None

                if pd.notna(raw_pts) and str(raw_pts).strip() != "" and str(raw_pts).strip().lower() != "nan":
                    try:
                        earned_pts = float(raw_pts)
                        is_graded = True
                    except (ValueError, TypeError):
                        is_graded = False

                # If status is "Returned" and Points value is empty or null, count it as zero points
                if not is_graded and status.strip().lower() == "returned":
                    earned_pts = 0.0
                    is_graded = True

                is_missing = not is_graded and (status in ["Viewed", "Not turned in"] or "turned in" in status.lower())

                rec = AssignmentRecord(
                    name=assign_title,
                    category=cat_canonical,
                    original_category=raw_cat,
                    due_date=due_str,
                    status=status,
                    points=earned_pts,
                    max_points=max_pts,
                    grade=grade_str,
                    feedback=feedback,
                    is_graded=is_graded,
                    is_missing=is_missing,
                    due_date_obj=due_dt,
                )

                cat_obj = cat_summaries[cat_canonical]

                if is_graded and earned_pts is not None:
                    cat_obj.points_earned += earned_pts
                    cat_obj.points_possible += max_pts
                    cat_obj.graded_count += 1
                    cat_obj.is_active = True
                    graded_list.append(rec)
                else:
                    cat_obj.missing_count += 1
                    missing_list.append(rec)

            # Compute category percentage and total weighted score
            total_weighted_points = 0.0
            sum_active_weights = 0.0

            for cat_name, cat_obj in cat_summaries.items():
                if cat_obj.points_possible > 0:
                    pct = (cat_obj.points_earned / cat_obj.points_possible) * 100.0
                    cat_obj.score_percent = round(pct, 2)
                    total_weighted_points += (cat_obj.score_percent / 100.0) * cat_obj.weight
                    sum_active_weights += cat_obj.weight
                elif cat_obj.points_earned > 0:
                    # Extra credit with 0 max points
                    cat_obj.score_percent = 100.0
                    total_weighted_points += cat_obj.weight
                    sum_active_weights += cat_obj.weight

            # Normalize overall percentage by active weights
            if sum_active_weights > 0:
                overall_pct = round((total_weighted_points / sum_active_weights) * 100.0, 1)
            else:
                overall_pct = 0.0

            letter_grade = calculate_letter_grade(overall_pct)

            # Sort assignments deterministically in descending order by Assignment name
            graded_list.sort(key=lambda r: str(r.name).lower().strip(), reverse=True)
            missing_list.sort(key=lambda r: str(r.name).lower().strip(), reverse=True)

            profiles.append(
                StudentGradeProfile(
                    course_name=course_name,
                    full_name=full_str,
                    first_name=first_name.strip(),
                    last_name=last_name.strip(),
                    email=email.strip(),
                    categories=cat_summaries,
                    overall_percent=overall_pct,
                    letter_grade=letter_grade,
                    graded_assignments=graded_list,
                    missing_assignments=missing_list,
                    total_graded=len(graded_list),
                    total_missing=len(missing_list),
                    parent_emails=matched_parents,
                )
            )

        # Sort profiles alphabetically by last name, then first name
        profiles.sort(key=lambda p: (p.last_name.lower(), p.first_name.lower()))
        return profiles
