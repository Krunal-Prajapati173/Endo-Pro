from pathlib import Path
import hashlib
import re

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from io import BytesIO


st.set_page_config(
    page_title="EduPro | Learner intelligence",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

ROOT = Path(__file__).resolve().parent
DATA_FILE = ROOT / "EduPro_Online_Platform.xlsx"
AGE_BINS = [0, 18, 26, 31, 36, 121]
AGE_LABELS = ["Under 18", "18–25", "26–30", "31–35", "36+"]
LEVEL_ORDER = ["Beginner", "Intermediate", "Advanced"]
COLORS = ["#236b62", "#e2a45d", "#6683a3", "#b86c5b", "#a4b69a", "#876b9c"]
TRANSACTION_COLUMNS = ["TransactionID", "UserID", "CourseID", "TeacherID", "TransactionDate", "Amount", "PaymentMethod"]
REQUIRED_UPLOAD_COLUMNS = ["UserID", "CourseID", "TeacherID", "TransactionDate", "Amount"]
UPLOAD_ALIASES = {
    "transactionid": "TransactionID", "transaction": "TransactionID", "enrollmentid": "TransactionID", "transactionnumber": "TransactionID",
    "userid": "UserID", "learnerid": "UserID", "studentid": "UserID",
    "courseid": "CourseID", "classid": "CourseID",
    "teacherid": "TeacherID", "instructorid": "TeacherID", "tutorid": "TeacherID",
    "transactiondate": "TransactionDate", "enrollmentdate": "TransactionDate", "purchasedate": "TransactionDate", "date": "TransactionDate", "datetime": "TransactionDate",
    "amount": "Amount", "transactionamount": "Amount", "paidamount": "Amount", "revenue": "Amount",
    "paymentmethod": "PaymentMethod", "paymenttype": "PaymentMethod", "payment": "PaymentMethod",
}


@st.cache_data(show_spinner="Reading workbook and checking its tables…")
def load_workbook(path: str, modified: float):
    """Read the source workbook; modified time makes edits invalidate the cache."""
    del modified
    sheets = pd.read_excel(path, sheet_name=None)
    required = {
        "Users": {"UserID", "Age", "Gender"},
        "Teachers": {"TeacherID", "TeacherName", "Expertise", "YearsOfExperience", "TeacherRating"},
        "Courses": {"CourseID", "CourseName", "CourseCategory", "CourseType", "CourseLevel", "CoursePrice", "CourseDuration", "CourseRating"},
        "Transactions": {"TransactionID", "UserID", "CourseID", "TeacherID", "TransactionDate", "Amount", "PaymentMethod"},
    }
    missing_sheets = set(required) - set(sheets)
    if missing_sheets:
        raise ValueError("Workbook is missing sheet(s): " + ", ".join(sorted(missing_sheets)))
    for sheet, columns in required.items():
        missing_columns = columns - set(sheets[sheet].columns)
        if missing_columns:
            raise ValueError(f"{sheet} is missing required columns: " + ", ".join(sorted(missing_columns)))

    users = sheets["Users"].copy()
    courses = sheets["Courses"].copy()
    teachers = sheets["Teachers"].copy()
    tx = sheets["Transactions"].copy()
    for frame, col in [(tx, "TransactionDate")]:
        frame[col] = pd.to_datetime(frame[col], errors="coerce")
    tx["Amount"] = pd.to_numeric(tx["Amount"], errors="coerce")
    users["Age"] = pd.to_numeric(users["Age"], errors="coerce")
    for col in ["CoursePrice", "CourseDuration", "CourseRating"]:
        courses[col] = pd.to_numeric(courses[col], errors="coerce")
    teachers["TeacherRating"] = pd.to_numeric(teachers["TeacherRating"], errors="coerce")
    teachers["YearsOfExperience"] = pd.to_numeric(teachers["YearsOfExperience"], errors="coerce")
    users["AgeGroup"] = pd.cut(users["Age"], bins=AGE_BINS, labels=AGE_LABELS, right=False)
    user_cols = [c for c in ["UserID", "Age", "Gender", "AgeGroup"] if c in users]
    course_cols = [c for c in ["CourseID", "CourseName", "CourseCategory", "CourseType", "CourseLevel", "CoursePrice", "CourseDuration", "CourseRating"] if c in courses]
    teacher_cols = [c for c in ["TeacherID", "TeacherName", "Expertise", "YearsOfExperience", "TeacherRating"] if c in teachers]
    joined = join_transactions(tx, users, courses, teachers)
    return users, courses, teachers, tx, joined


def join_transactions(tx, users, courses, teachers):
    user_cols = ["UserID", "Age", "Gender", "AgeGroup"]
    course_cols = ["CourseID", "CourseName", "CourseCategory", "CourseType", "CourseLevel", "CoursePrice", "CourseDuration", "CourseRating"]
    teacher_cols = ["TeacherID", "TeacherName", "Expertise", "YearsOfExperience", "TeacherRating"]
    joined = tx.merge(users[user_cols], on="UserID", how="left", validate="m:1")
    joined = joined.merge(courses[course_cols], on="CourseID", how="left", validate="m:1")
    return joined.merge(teachers[teacher_cols], on="TeacherID", how="left", validate="m:1")


def read_csv_table(file_bytes):
    return pd.read_csv(BytesIO(file_bytes), encoding="utf-8-sig", sep=None, engine="python")


def parse_transaction_csv(file_bytes, users, courses, teachers):
    """Validate an uploaded enrollment CSV before using it as the active data source."""
    try:
        uploaded = read_csv_table(file_bytes)
    except Exception as exc:
        return None, None, [f"Could not read this CSV: {exc}"]

    original_headers = [str(column).strip() for column in uploaded.columns]
    rename = {}
    for column in uploaded.columns:
        normalized = re.sub(r"[^a-z0-9]", "", str(column).strip().lower())
        canonical = UPLOAD_ALIASES.get(normalized)
        if canonical:
            rename[column] = canonical
    uploaded = uploaded.rename(columns=rename)
    duplicates = uploaded.columns[uploaded.columns.duplicated()].tolist()
    if duplicates:
        return None, None, [f"More than one CSV column maps to {', '.join(map(str, duplicates))}."]

    missing = [column for column in REQUIRED_UPLOAD_COLUMNS if column not in uploaded.columns]
    if missing:
        found = ", ".join(original_headers) if original_headers else "no headers"
        return None, None, [
            "Could not identify required fields: " + ", ".join(missing)
            + f". Detected headers: {found}. Use the CSV template, or name columns UserID, CourseID, TeacherID, TransactionDate, and Amount."
        ]
    if "TransactionID" not in uploaded.columns:
        uploaded["TransactionID"] = [f"UPLOAD-{i:08d}" for i in range(1, len(uploaded) + 1)]
    if "PaymentMethod" not in uploaded.columns:
        uploaded["PaymentMethod"] = "Not provided"
    uploaded = uploaded[TRANSACTION_COLUMNS].copy()
    problems = []
    if uploaded.empty:
        problems.append("The file has headers but no transaction rows.")
    blank_rows = int(uploaded[TRANSACTION_COLUMNS].isna().any(axis=1).sum())
    if blank_rows:
        problems.append(f"{blank_rows} row(s) have blank required values.")
    duplicate_ids = int(uploaded.TransactionID.duplicated().sum())
    if duplicate_ids:
        problems.append(f"{duplicate_ids} duplicate TransactionID value(s).")
    uploaded["TransactionDate"] = pd.to_datetime(uploaded["TransactionDate"], errors="coerce")
    bad_dates = int(uploaded.TransactionDate.isna().sum())
    if bad_dates:
        problems.append(f"{bad_dates} row(s) have an invalid TransactionDate.")
    uploaded["Amount"] = pd.to_numeric(uploaded["Amount"], errors="coerce")
    bad_amounts = int(uploaded.Amount.isna().sum())
    if bad_amounts:
        problems.append(f"{bad_amounts} row(s) have a non-numeric Amount.")
    for column, known, label in [
        ("UserID", users.UserID, "learner"),
        ("CourseID", courses.CourseID, "course"),
        ("TeacherID", teachers.TeacherID, "teacher"),
    ]:
        unknown = int((~uploaded[column].isin(known)).sum())
        if unknown:
            problems.append(f"{unknown} row(s) use a {label} ID not found in the workbook.")
    if problems:
        return None, None, problems
    return uploaded, join_transactions(uploaded, users, courses, teachers), []


def show_generic_csv(frame, filename):
    """Build a dashboard from whatever useful fields exist in the uploaded CSV."""
    dataset_title = Path(filename).stem.replace("_", " ").replace("-", " ").title()
    st.title(f"{dataset_title} dashboard")
    st.caption(f"{filename} · Dashboard created from the columns in your file")

    table = frame.copy()
    date_col = None
    date_candidates = [c for c in table.columns if any(k in str(c).lower() for k in ["date", "time"]) ]
    for candidate in date_candidates:
        parsed = pd.to_datetime(table[candidate], errors="coerce")
        if parsed.notna().mean() >= 0.8:
            table[candidate] = parsed
            date_col = candidate
            break
    if date_col is None:
        for candidate in table.select_dtypes(include="object").columns:
            parsed = pd.to_datetime(table[candidate], errors="coerce")
            if len(table) and parsed.notna().mean() >= 0.8 and parsed.nunique() > 1:
                table[candidate] = parsed
                date_col = candidate
                break

    numeric_columns = []
    for column in table.columns:
        if column == date_col:
            continue
        converted = pd.to_numeric(table[column], errors="coerce")
        if converted.notna().sum() > 0 and converted.notna().mean() >= 0.6:
            table[column] = converted
            numeric_columns.append(column)

    view = table
    if date_col and table[date_col].notna().any():
        first_day, last_day = table[date_col].min().date(), table[date_col].max().date()
        chosen_dates = st.date_input(
            "Filter dates",
            value=(first_day, last_day),
            min_value=first_day,
            max_value=last_day,
            key=f"csv_dates_{hashlib.sha1(filename.encode()).hexdigest()[:8]}",
        )
        if isinstance(chosen_dates, (tuple, list)) and len(chosen_dates) == 2:
            start, end = pd.Timestamp(chosen_dates[0]), pd.Timestamp(chosen_dates[1])
            view = table[table[date_col].between(start, end + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1))]

    st.subheader("At a glance")
    kpis = [("Rows in view", f"{len(view):,}"), ("Fields", f"{len(table.columns):,}")]
    if date_col and table[date_col].notna().any():
        kpis.append(("Date range", f"{table[date_col].min():%b %Y} - {table[date_col].max():%b %Y}"))
    for column in numeric_columns[: max(0, 4 - len(kpis))]:
        kpis.append((f"Total {column}", f"{view[column].sum(min_count=1):,.2f}"))
    metric_cols = st.columns(min(4, len(kpis)))
    for col, (label, value) in zip(metric_cols, kpis):
        col.metric(label, value)

    categorical_columns = [
        c for c in view.select_dtypes(include=["object", "category"]).columns
        if c != date_col and 1 < view[c].nunique(dropna=True) <= 40
    ]
    overview_tab, details_tab = st.tabs(["Overview", "Data & quality"])
    with overview_tab:
        if date_col and numeric_columns:
            selected = st.multiselect(
                "Measures to compare over time",
                numeric_columns,
                default=numeric_columns[: min(3, len(numeric_columns))],
            )
            if selected:
                trend = view[[date_col] + selected].dropna(subset=[date_col]).groupby(date_col, as_index=False)[selected].sum().sort_values(date_col)
                fig = px.line(trend, x=date_col, y=selected, markers=True, title=f"Measures over time by {date_col}", color_discrete_sequence=COLORS)
                st.plotly_chart(chart(fig, 420), width="stretch")
        elif numeric_columns:
            measure = st.selectbox("Choose a measure", numeric_columns)
            fig = px.histogram(view, x=measure, nbins=25, title=f"Distribution of {measure}", color_discrete_sequence=[COLORS[0]])
            st.plotly_chart(chart(fig, 380), width="stretch")
        else:
            st.info("This file has no numeric measures to chart. Use the data tab to inspect its categories and rows.")

        if categorical_columns:
            category = st.selectbox("Break down by", categorical_columns)
            counts = view[category].fillna("Blank").astype(str).value_counts().head(15).sort_values().rename_axis(category).reset_index(name="Rows")
            fig = px.bar(counts, x="Rows", y=category, orientation="h", title=f"Most common {category} values", color_discrete_sequence=[COLORS[1]])
            st.plotly_chart(chart(fig, 340), width="stretch")

        if numeric_columns:
            summary = view[numeric_columns].describe().T.rename_axis("Measure").reset_index()
            st.subheader("Numeric summary")
            st.dataframe(summary, hide_index=True, width="stretch")

    with details_tab:
        missing_cells = int(table.isna().sum().sum())
        duplicate_rows = int(table.duplicated().sum())
        q1, q2 = st.columns(2)
        q1.metric("Blank cells", f"{missing_cells:,}")
        q2.metric("Duplicate rows", f"{duplicate_rows:,}")
        st.subheader("Uploaded data")
        st.dataframe(view.head(2_000), hide_index=True, width="stretch")
        st.download_button("Download filtered CSV", view.to_csv(index=False).encode("utf-8-sig"), "dashboard_filtered_data.csv", "text/csv")


def quality_report(users, courses, teachers, tx, joined):
    return [
        ("User table has no blank cells", int(users.isna().sum().sum())),
        ("Course table has no blank cells", int(courses.isna().sum().sum())),
        ("Teacher table has no blank cells", int(teachers.isna().sum().sum())),
        ("Transaction table has no blank cells", int(tx.isna().sum().sum())),
        ("Learner IDs are unique", int(users.UserID.duplicated().sum())),
        ("Course IDs are unique", int(courses.CourseID.duplicated().sum())),
        ("Teacher IDs are unique", int(teachers.TeacherID.duplicated().sum())),
        ("Transaction IDs are unique", int(tx.TransactionID.duplicated().sum())),
        ("Transactions without a learner match", int(joined.Age.isna().sum())),
        ("Transactions without a course match", int(joined.CourseName.isna().sum())),
        ("Transactions without a teacher match", int(joined.TeacherName.isna().sum())),
        ("Transactions with invalid date", int(tx.TransactionDate.isna().sum())),
        ("Transactions with missing amount", int(tx.Amount.isna().sum())),
        ("Transactions with negative amount", int(tx.Amount.lt(0).sum())),
        ("Learner ages outside 0–120", int((~users.Age.between(0, 120)).sum())),
        ("Course ratings outside 0–5", int((~courses.CourseRating.between(0, 5)).sum())),
        ("Teacher ratings outside 0–5", int((~teachers.TeacherRating.between(0, 5)).sum())),
        ("Courses with negative listed price", int(courses.CoursePrice.lt(0).sum())),
    ]


def chart(fig, height=360):
    fig.update_layout(
        height=height,
        margin=dict(l=12, r=12, t=55, b=12),
        font=dict(family="Inter, Aptos, sans-serif", color="#273632", size=12),
        title=dict(font=dict(size=16, color="#263b35"), x=0.02),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    fig.update_xaxes(showgrid=False, linecolor="#e5e9e5", zeroline=False)
    fig.update_yaxes(gridcolor="#edf0ec", zeroline=False)
    return fig


st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@500;600;700;800&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
h1, h2, h3 { font-family: 'Manrope', sans-serif; letter-spacing: -0.035em; }
.stApp { background: #f6f7f3; }
[data-testid="stSidebar"] { background: #eef2eb; border-right: 1px solid #e1e7df; }
[data-testid="stMetric"] { background: #fff; border: 1px solid #e8ece6; border-radius: 14px; padding: 16px 18px; box-shadow: 0 3px 12px rgba(36,56,48,.035); }
[data-testid="stMetric"] { min-width: 0; height: auto; }
[data-testid="stMetricLabel"], [data-testid="stMetricLabel"] > div { color: #68766e; white-space: normal !important; overflow: visible !important; text-overflow: clip !important; line-height: 1.25; overflow-wrap: anywhere; }
[data-testid="stMetricValue"], [data-testid="stMetricValue"] > div { color: #244b43; font-family: 'Manrope', sans-serif; white-space: normal !important; overflow: visible !important; text-overflow: clip !important; overflow-wrap: anywhere; line-height: 1.15; font-size: clamp(1.2rem, 2.2vw, 2rem) !important; }
[data-testid="stMetricDelta"] { white-space: normal !important; overflow-wrap: anywhere; }
[data-testid="stPlotlyChart"] { min-width: 0; }
.stMarkdown, .stCaption, p { overflow-wrap: anywhere; }
.block-container { padding-top: 2rem; padding-bottom: 3rem; }
div[data-testid="stTabs"] button { font-weight: 600; }
@media (max-width: 900px) {
  .block-container { padding-left: 1rem; padding-right: 1rem; }
  [data-testid="stMetric"] { padding: 12px; }
}
</style>
""", unsafe_allow_html=True)

st.sidebar.markdown("## 🎓 EduPro")
st.sidebar.caption("Learner intelligence · Workbook edition")
if not DATA_FILE.exists():
    st.error(f"I could not find `{DATA_FILE.name}` beside this app. Place the workbook in `{ROOT}` and reload.")
    st.stop()

try:
    users, courses, teachers, tx_all, data = load_workbook(str(DATA_FILE), DATA_FILE.stat().st_mtime)
except Exception as exc:
    st.error(f"The workbook could not be loaded: {exc}")
    st.stop()

st.sidebar.markdown("### Upload CSV data")
uploaded_file = st.sidebar.file_uploader(
    "Choose a CSV file",
    type=["csv"],
    help="EduPro transaction files are checked against the workbook. Other tabular CSV files get a dashboard built from their own columns.",
)
template = pd.DataFrame(columns=TRANSACTION_COLUMNS).to_csv(index=False).encode("utf-8")
st.sidebar.download_button(
    "Download CSV template",
    template,
    "edupro_transactions_template.csv",
    "text/csv",
    width="stretch",
)

uploaded_transactions = None
uploaded_data = None
generic_csv = None
upload_problems = []
upload_mode = "Workbook"
if uploaded_file is not None:
    uploaded_bytes = uploaded_file.getvalue()
    uploaded_transactions, uploaded_data, upload_problems = parse_transaction_csv(
        uploaded_bytes, users, courses, teachers
    )
    if upload_problems:
        if upload_problems[0].startswith("Could not identify required fields:"):
            try:
                candidate = read_csv_table(uploaded_bytes)
                if not candidate.empty and len(candidate.columns) > 0:
                    generic_csv = candidate
                    st.sidebar.success("Custom CSV dashboard ready. Charts and summaries adapt to your file's columns.")
                else:
                    st.sidebar.error("CSV is empty. Upload a CSV with a header row and data rows.")
            except Exception as exc:
                st.sidebar.error(f"Could not read this CSV: {exc}")
        else:
            st.sidebar.error("CSV not used: " + " ".join(upload_problems))
    else:
        st.sidebar.success(f"CSV ready: {len(uploaded_transactions):,} valid transaction rows.")

if uploaded_transactions is not None:
    upload_mode = st.sidebar.radio("Dashboard data source", ["Workbook", "Uploaded transactions"], index=1)
    if upload_mode == "Uploaded transactions":
        tx_all, data = uploaded_transactions, uploaded_data
elif generic_csv is not None:
    upload_mode = st.sidebar.radio("Dashboard data source", ["Workbook", "Explore uploaded CSV"], index=1)
    if upload_mode == "Explore uploaded CSV":
        show_generic_csv(generic_csv, uploaded_file.name)
        st.stop()
else:
    st.sidebar.caption("EduPro transaction CSVs open in the dashboard. Other CSVs open in the general explorer.")

active_upload = upload_mode != "Workbook" if uploaded_file is not None and upload_mode else False
source_token = hashlib.sha1(uploaded_file.getvalue()).hexdigest()[:10] if active_upload else "workbook"

quality = quality_report(users, courses, teachers, tx_all, data)
min_date, max_date = tx_all.TransactionDate.min(), tx_all.TransactionDate.max()
if pd.isna(min_date) or pd.isna(max_date):
    st.error("No valid transaction dates were found. Check the TransactionDate column in the workbook.")
    st.stop()

st.sidebar.markdown("### Shape the view")
date_range = st.sidebar.date_input(
    "Transaction dates",
    value=(min_date.date(), max_date.date()),
    min_value=min_date.date(),
    max_value=max_date.date(),
    key=f"transaction_dates_{source_token}",
)
if isinstance(date_range, (tuple, list)) and len(date_range) == 2:
    start_date, end_date = pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1])
else:
    start_date = end_date = pd.Timestamp(date_range[0] if isinstance(date_range, (tuple, list)) else date_range)
gender_opts = sorted(data.Gender.dropna().unique().tolist())
category_opts = sorted(data.CourseCategory.dropna().unique().tolist())
level_opts = [x for x in LEVEL_ORDER if x in data.CourseLevel.dropna().unique()]
type_opts = sorted(data.CourseType.dropna().unique().tolist())
age_opts = [x for x in AGE_LABELS if x in data.AgeGroup.dropna().astype(str).unique()]
sel_age = st.sidebar.multiselect("Learner age", age_opts, default=age_opts)
sel_gender = st.sidebar.multiselect("Gender", gender_opts, default=gender_opts)
sel_cat = st.sidebar.multiselect("Course category", category_opts, default=category_opts)
sel_lvl = st.sidebar.multiselect("Course level", level_opts, default=level_opts)
sel_type = st.sidebar.multiselect("Course type", type_opts, default=type_opts)
source_label = uploaded_file.name if active_upload else "EduPro_Online_Platform.xlsx"
st.sidebar.caption(f"Active data: {source_label}")
if st.sidebar.button("Clear cached workbook data", width="stretch"):
    load_workbook.clear()
    st.rerun()

f = data.loc[
    data.TransactionDate.between(start_date, end_date + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1))
    & data.AgeGroup.astype(str).isin(sel_age)
    & data.Gender.isin(sel_gender)
    & data.CourseCategory.isin(sel_cat)
    & data.CourseLevel.isin(sel_lvl)
    & data.CourseType.isin(sel_type)
].copy()
learner_ids = f.UserID.dropna().unique()
fu = users[users.UserID.isin(learner_ids)].copy()
period_text = f"{start_date:%b %d, %Y} – {end_date:%b %d, %Y}"

st.markdown("# EduPro learner analytics")
st.markdown("Explore learner demographics, enrollment patterns, course demand, and recorded revenue.")
source_display = "EduPro workbook" if source_label == "EduPro_Online_Platform.xlsx" else source_label
st.caption(f"Source: {source_display} · Available data: {min_date:%b %d, %Y} to {max_date:%b %d, %Y}")
st.caption(f"Current selection: {period_text} · {len(f):,} enrollment records")
if not f.empty:
    st.caption("Definitions: enrollments count transaction records; learners are unique IDs. Revenue sums the recorded Amount field; currency is unspecified.")
else:
    st.warning("No transactions match these filters. Adjust a date or filter to see the dashboard.")

if f.empty:
    st.stop()

active = f.UserID.nunique()
top_category = f.CourseCategory.value_counts().idxmax()
paid_share = f.CourseType.eq("Paid").mean() * 100 if len(f) else 0
revenue = f.Amount.sum(min_count=1)
revenue = 0 if pd.isna(revenue) else revenue
revenue_per_paid = revenue / max(1, f.CourseType.eq("Paid").sum())
k1, k2, k3 = st.columns(3)
k1.metric("Total enrollments", f"{len(f):,}", help="Number of transaction rows matching the current filters.")
k2.metric("Unique learners", f"{active:,}", help="Distinct learner IDs represented by the filtered transactions.")
k3.metric("Avg. enrollments per learner", f"{len(f) / active:.2f}", help="Filtered enrollment rows divided by distinct learners.")
k4, k5 = st.columns(2)
k4.metric("Recorded revenue", f"{revenue:,.2f}", help="Sum of transaction Amount values; currency is not specified in the workbook.")
k5.metric("Leading course category", top_category, f"{int(f.CourseCategory.eq(top_category).sum()):,} enrollments")

overview, learners_tab, courses_tab, teachers_tab, quality_tab, data_tab = st.tabs([
    "Overview", "Learners", "Courses & revenue", "Teachers", "Data quality", "Explore & export"
])

with overview:
    st.subheader("The learning picture")
    left, right = st.columns([1.35, 1])
    with left:
        monthly = f.assign(Month=f.TransactionDate.dt.to_period("M").dt.to_timestamp()).groupby("Month").size().rename("Enrollments").reset_index()
        st.plotly_chart(chart(px.line(monthly, x="Month", y="Enrollments", markers=True, title="Enrollment activity over time", color_discrete_sequence=[COLORS[0]]), 340), width="stretch")
    with right:
        by_type = f.groupby("CourseType", dropna=False).agg(Enrollments=("TransactionID", "size"), Revenue=("Amount", "sum")).reset_index()
        st.plotly_chart(chart(px.bar(by_type, x="CourseType", y="Enrollments", color="CourseType", title="Free and paid enrollments", color_discrete_sequence=COLORS), 340), width="stretch")
    col1, col2 = st.columns(2)
    with col1:
        age = fu.AgeGroup.astype(str).value_counts().reindex(AGE_LABELS).dropna().rename_axis("Age group").reset_index(name="Learners")
        age = age[age.Learners.gt(0)]
        st.plotly_chart(chart(px.bar(age, x="Age group", y="Learners", title="Learners by age group", color_discrete_sequence=[COLORS[2]]), 320), width="stretch")
    with col2:
        cats = f.CourseCategory.value_counts().rename_axis("Category").reset_index(name="Enrollments").sort_values("Enrollments")
        st.plotly_chart(chart(px.bar(cats, x="Enrollments", y="Category", orientation="h", title="Enrollment by course category", color_discrete_sequence=[COLORS[1]]), 320), width="stretch")
    st.caption("Each chart uses the selected date range and sidebar filters. A peak in activity describes this workbook; it does not explain why the peak happened.")

with learners_tab:
    st.subheader("Learners and engagement")
    col1, col2 = st.columns(2)
    with col1:
        gender = fu.Gender.value_counts().rename_axis("Gender").reset_index(name="Learners")
        st.plotly_chart(chart(px.pie(gender, names="Gender", values="Learners", hole=.62, title="Learners by gender", color_discrete_sequence=COLORS), 340), width="stretch")
    with col2:
        per = f.groupby("UserID").size().rename("Enrollments")
        dist = per.value_counts().sort_index().rename_axis("Enrollments per learner").reset_index(name="Learners")
        st.plotly_chart(chart(px.bar(dist, x="Enrollments per learner", y="Learners", title="How many courses learners enrolled in", color_discrete_sequence=[COLORS[0]]), 340), width="stretch")
    by_age = f.groupby("AgeGroup", observed=False).agg(Enrollments=("TransactionID", "size"), Learners=("UserID", "nunique")).reset_index()
    by_age["Enrollments per learner"] = by_age.Enrollments.div(by_age.Learners.replace(0, np.nan)).round(2)
    by_age["AgeGroup"] = pd.Categorical(by_age.AgeGroup, AGE_LABELS, ordered=True)
    by_age = by_age.sort_values("AgeGroup").dropna(subset=["Enrollments"])
    st.plotly_chart(chart(px.bar(by_age, x="AgeGroup", y="Enrollments per learner", text="Enrollments per learner", title="Average enrollments per learner by age group", color_discrete_sequence=[COLORS[2]]), 310), width="stretch")
    st.caption("Engagement here means recorded enrollments only. The workbook does not contain attendance, completion, or learning outcome data.")

with courses_tab:
    st.subheader("What learners choose")
    col1, col2 = st.columns([1.2, 1])
    with col1:
        cdata = f.groupby("CourseCategory").agg(Enrollments=("TransactionID", "size"), Revenue=("Amount", "sum")).reset_index().sort_values("Enrollments")
        st.plotly_chart(chart(px.bar(cdata, x="Enrollments", y="CourseCategory", orientation="h", title="Category demand", color_discrete_sequence=[COLORS[0]]), 390), width="stretch")
    with col2:
        level = f.CourseLevel.value_counts().reindex(LEVEL_ORDER).dropna().rename_axis("Level").reset_index(name="Enrollments")
        st.plotly_chart(chart(px.pie(level, names="Level", values="Enrollments", hole=.62, title="Enrollment by course level", color_discrete_sequence=COLORS), 390), width="stretch")
    top = f.groupby("CourseName").agg(Enrollments=("TransactionID", "size"), Revenue=("Amount", "sum"), Rating=("CourseRating", "mean")).sort_values("Enrollments", ascending=False).head(12).reset_index()
    c1, c2 = st.columns([1.1, .9])
    with c1:
        top_sorted = top.sort_values("Enrollments")
        st.plotly_chart(chart(px.bar(top_sorted, x="Enrollments", y="CourseName", orientation="h", hover_data=["Revenue", "Rating"], title="Most enrolled courses", color_discrete_sequence=[COLORS[1]]), 420), width="stretch")
    with c2:
        paid = f.groupby("CourseType").agg(Enrollments=("TransactionID", "size"), Revenue=("Amount", "sum"), Average_transaction=("Amount", "mean")).reset_index()
        st.dataframe(paid, hide_index=True, width="stretch")
        st.markdown("**How to read revenue**")
        st.write(f"There are {paid_share:.1f}% paid-course enrollments in this view. Revenue is the sum recorded in `Transactions.Amount`; free enrollments contribute zero in this workbook.")
        st.caption(f"Mean amount per paid enrollment: {revenue_per_paid:,.2f} workbook currency units.")
    with st.expander("Course ratings and catalog detail"):
        catalog = courses[["CourseName", "CourseCategory", "CourseType", "CourseLevel", "CoursePrice", "CourseDuration", "CourseRating"]].copy()
        st.dataframe(catalog.sort_values("CourseRating", ascending=False), hide_index=True, width="stretch")

with teachers_tab:
    st.subheader("Instructor view")
    st.caption("A rating is a recorded course or teacher rating in the workbook. Enrollment volume and ratings do not establish teaching effectiveness or causation.")
    teacher_view = f.groupby(["TeacherID", "TeacherName", "Expertise"], dropna=False).agg(
        Enrollments=("TransactionID", "size"), Revenue=("Amount", "sum"),
        Avg_teacher_rating=("TeacherRating", "mean"), Avg_course_rating=("CourseRating", "mean"),
        Experience_years=("YearsOfExperience", "first"),
    ).reset_index().sort_values("Enrollments", ascending=False)
    c1, c2 = st.columns([1.15, 1])
    with c1:
        activity_fig = px.scatter(
            teacher_view,
            x="Enrollments",
            y="Avg_teacher_rating",
            color="Expertise",
            hover_name="TeacherName",
            hover_data={
                "TeacherID": True,
                "Expertise": True,
                "Enrollments": True,
                "Revenue": ":,.2f",
                "Avg_course_rating": ":.2f",
                "Experience_years": True,
            },
            log_x=True,
            title="Enrollment activity and teacher rating",
            color_discrete_sequence=COLORS,
        )
        activity_fig.update_layout(showlegend=False)
        activity_fig.update_traces(marker=dict(size=11, opacity=0.8, line=dict(width=1, color="white")))
        activity_fig.update_xaxes(title="Enrollments per teacher (log scale)")
        activity_fig.update_yaxes(title="Average teacher rating (0–5)", range=[0, 5.2])
        st.plotly_chart(chart(activity_fig, 390), width="stretch")
        st.caption("The log scale spreads out lower enrollment counts so they remain visible. Hover over a point for teacher, expertise, revenue, and rating details.")
    with c2:
        exp = f.groupby("Expertise").agg(Enrollments=("TransactionID", "size"), Teachers=("TeacherID", "nunique")).reset_index().sort_values("Enrollments")
        st.plotly_chart(chart(px.bar(exp, x="Enrollments", y="Expertise", orientation="h", title="Enrollments by teacher expertise", color_discrete_sequence=[COLORS[2]]), 390), width="stretch")
    teacher_table = teacher_view.rename(columns={
        "TeacherID": "Teacher ID",
        "TeacherName": "Teacher",
        "Expertise": "Subject area",
        "Enrollments": "Enrollments",
        "Revenue": "Revenue",
        "Avg_teacher_rating": "Teacher rating",
        "Avg_course_rating": "Course rating",
        "Experience_years": "Years experience",
    })
    st.dataframe(teacher_table, hide_index=True, width="stretch")

with quality_tab:
    st.subheader("Workbook health check")
    st.caption("These checks look for structural problems that could distort the dashboard. Zero means no issue was found by that check.")
    check_df = pd.DataFrame(quality, columns=["Check", "Issue count"])
    check_df["Status"] = np.where(check_df["Issue count"].eq(0), "Looks good", "Review")
    passed = int(check_df["Issue count"].eq(0).sum())
    q1, q2, q3 = st.columns(3)
    q1.metric("Checks passed", f"{passed} / {len(check_df)}")
    q2.metric("Workbook tables", "4")
    q3.metric("Transaction rows", f"{len(tx_all):,}")
    st.dataframe(check_df, hide_index=True, width="stretch")
    st.markdown("**Coverage summary**")
    st.write(f"{len(users):,} learner rows · {len(courses):,} course rows · {len(teachers):,} teacher rows · {len(tx_all):,} transaction rows.")
    st.write(f"The recorded transaction period is {min_date:%B %d, %Y} through {max_date:%B %d, %Y}. A complete workbook is still not proof that every real-world event was recorded.")

with data_tab:
    st.subheader("Explore the filtered records")
    cols = [c for c in ["TransactionDate", "TransactionID", "UserID", "Age", "AgeGroup", "Gender", "CourseName", "CourseCategory", "CourseType", "CourseLevel", "Amount", "PaymentMethod", "TeacherName", "Expertise"] if c in f]
    show = f[cols].sort_values("TransactionDate", ascending=False)
    st.caption(f"Showing {min(len(show), 1_000):,} of {len(show):,} matching records. Sensitive contact fields are excluded.")
    st.dataframe(show.head(1_000), hide_index=True, width="stretch")
    csv = show.to_csv(index=False).encode("utf-8-sig")
    st.download_button("Download filtered transactions (CSV)", csv, "edupro_filtered_transactions.csv", "text/csv", width="content")
    summary = f.groupby(["CourseCategory", "CourseType", "CourseLevel"], dropna=False).agg(Enrollments=("TransactionID", "size"), Learners=("UserID", "nunique"), Revenue=("Amount", "sum")).reset_index()
    st.download_button("Download category summary (CSV)", summary.to_csv(index=False).encode("utf-8-sig"), "edupro_category_summary.csv", "text/csv")

st.divider()
st.caption("EduPro Learner Intelligence · Descriptive analysis of the workbook snapshot · Built for learning and exploration")
