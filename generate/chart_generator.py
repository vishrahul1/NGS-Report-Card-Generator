"""
Chart Generator – creates a per-student bar chart PNG.

The chart shows average CDF and JEE scores per subject across all loaded exams.
This module is intentionally modular: update generate_student_chart() when the
exact series definition is clarified by the user.
"""

import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import SUBJECTS, CHART_COLORS, CDF_MAX_PER_SUBJECT, JEE_MAX_PER_SUBJECT, AUTO_SCALE_CDF


def compute_subject_averages(student_exams: list[dict]) -> tuple[dict, dict]:
    """
    Given a list of exam records for one student, compute per-subject
    AVERAGE PERCENTAGE for CDF and JEE exam types separately.

    Exam type is determined by the "Exam" field:
        starts with "CDF" → CDF bucket  (max marks = CDF_MAX_PER_SUBJECT = 30)
        starts with "JEE" → JEE bucket  (max marks = JEE_MAX_PER_SUBJECT = 80)

    Formula:
        avg  =  sum of non-AB scores  /  count of non-AB exams
        CDF % =  avg / 30 * 100
        JEE % =  avg / 80 * 100

    AB rows are skipped and do not count toward the total.

    Returns: (cdf_pcts, jee_pcts) – each is {subject: percentage_0_to_100}
    """
    subject_col_map = {
        "Mathematics": "Mathematics Score",
        "Physics":     "Physics Score",
        "Chemistry":   "Chemistry Score",
    }

    jee_scores = {s: [] for s in SUBJECTS}
    cdf_scores  = {s: [] for s in SUBJECTS}

    for exam in student_exams:
        exam_name = str(exam.get("Exam", "")).strip()
        is_cdf = exam_name.upper().startswith("CDF")
        is_jee = exam_name.upper().startswith("JEE")

        if not is_cdf and not is_jee:
            continue  # unknown exam type, skip

        for subject in SUBJECTS:
            col = subject_col_map[subject]
            val = exam.get(col, "")
            if val == "" or val is None:
                continue
            if str(val).strip().upper() == "AB":
                continue  # skip absent exams
            try:
                score = float(val)
                if is_cdf:
                    cdf_scores[subject].append(score)
                else:
                    jee_scores[subject].append(score)
            except (ValueError, TypeError):
                pass

    cdf_pcts = {
        s: round((sum(v) / len(v)) / CDF_MAX_PER_SUBJECT * 100, 1) if v else 0.0
        for s, v in cdf_scores.items()
    }
    jee_pcts = {
        s: round((sum(v) / len(v)) / JEE_MAX_PER_SUBJECT * 100, 1) if v else 0.0
        for s, v in jee_scores.items()
    }
    return cdf_pcts, jee_pcts


def generate_student_chart(
    student_username: str,
    student_exams: list[dict],
    output_path: str,
    class_averages: dict | None = None,
) -> None:
    """
    Generate a clustered 3D-style bar chart for the student and save as PNG.

    Parameters
    ----------
    student_username : str
        Used only for logging.
    student_exams : list[dict]
        All exam records for this student (from CSV).
    output_path : str
        Full path to save the PNG file.
    class_averages : dict | None
        Optional {subject: avg_score} for the class (JEE scale).
        If provided, a third bar series is drawn.
    """
    cdf_pcts, jee_pcts = compute_subject_averages(student_exams)

    subjects = SUBJECTS
    x = np.arange(len(subjects))
    n_series = 3 if class_averages else 2
    width = 0.25 if n_series == 3 else 0.30

    fig, ax = plt.subplots(figsize=(6, 3.5), dpi=130)
    fig.patch.set_facecolor("white")

    # Series 1: CDF average %
    cdf_vals = [cdf_pcts.get(s, 0) for s in subjects]
    bars1 = ax.bar(
        x - width / 2 if n_series == 2 else x - width,
        cdf_vals, width,
        label="CDF Avg %",
        color=CHART_COLORS["series1"],
        edgecolor="white", linewidth=0.5, zorder=3,
    )

    # Series 2: JEE average %
    jee_vals = [jee_pcts.get(s, 0) for s in subjects]
    bars2 = ax.bar(
        x + width / 2 if n_series == 2 else x,
        jee_vals, width,
        label="JEE Avg %",
        color=CHART_COLORS["series2"],
        edgecolor="white", linewidth=0.5, zorder=3,
    )

    # Series 3: Class average % (optional, reserved)
    if class_averages:
        class_vals = [class_averages.get(s, 0) for s in subjects]
        ax.bar(
            x + width, class_vals, width,
            label="Class Avg %",
            color=CHART_COLORS["series3"],
            edgecolor="white", linewidth=0.5, zorder=3,
        )

    # Value labels on bars
    for bar in list(bars1) + list(bars2):
        h = bar.get_height()
        if h > 0:
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                h + 0.8,
                f"{h:.1f}%",
                ha="center", va="bottom", fontsize=7, color="#333333",
            )

    ax.set_xticks(x)
    ax.set_xticklabels(subjects, fontsize=9)
    ax.set_ylabel("Average Percentage (%)", fontsize=9)
    ax.set_ylim(0, 115)
    ax.set_yticks(range(0, 101, 20))
    ax.set_yticklabels([f"{v}%" for v in range(0, 101, 20)], fontsize=8)
    ax.set_title("")
    ax.legend(fontsize=8, loc="upper right")
    ax.yaxis.grid(True, linestyle="--", alpha=0.5, zorder=0)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    plt.tight_layout()
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    plt.savefig(output_path, dpi=130, bbox_inches="tight", facecolor="white")
    plt.close(fig)
