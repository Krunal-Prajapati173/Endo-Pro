from pathlib import Path

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether
)


ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "EduPro_Project_Learning_Guide.pdf"
book = pd.read_excel(ROOT / "EduPro_Online_Platform.xlsx", sheet_name=None)
users, courses, teachers, tx = (book[k] for k in ["Users", "Courses", "Teachers", "Transactions"])
joined = tx.merge(users[["UserID", "Age", "Gender"]], on="UserID").merge(
    courses[["CourseID", "CourseName", "CourseCategory", "CourseType", "CourseLevel"]], on="CourseID"
).merge(teachers[["TeacherID", "TeacherName"]], on="TeacherID")
period = f"{tx.TransactionDate.min():%B %d, %Y} to {tx.TransactionDate.max():%B %d, %Y}"
revenue = tx.Amount.sum()
paid_n = int(joined.CourseType.eq("Paid").sum())
top_category = joined.CourseCategory.value_counts().idxmax()
top_category_n = int(joined.CourseCategory.value_counts().max())
unique_learners = int(tx.UserID.nunique())
mean_per_learner = len(tx) / unique_learners

INK = colors.HexColor("#263b35")
GREEN = colors.HexColor("#236b62")
PALE = colors.HexColor("#eef3ed")
GOLD = colors.HexColor("#e2a45d")
MUTED = colors.HexColor("#5c6b64")
styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name="CoverTitle", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=27, leading=33, textColor=INK, alignment=TA_CENTER, spaceAfter=10))
styles.add(ParagraphStyle(name="CoverSub", parent=styles["Normal"], fontSize=12, leading=18, textColor=MUTED, alignment=TA_CENTER))
styles.add(ParagraphStyle(name="Section", parent=styles["Heading1"], fontName="Helvetica-Bold", fontSize=19, leading=24, textColor=GREEN, spaceBefore=8, spaceAfter=9))
styles.add(ParagraphStyle(name="Sub", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=12.5, leading=16, textColor=INK, spaceBefore=8, spaceAfter=4))
styles.add(ParagraphStyle(name="Body2", parent=styles["BodyText"], fontSize=10, leading=15, textColor=INK, spaceAfter=7))
styles.add(ParagraphStyle(name="Small2", parent=styles["BodyText"], fontSize=8.5, leading=12, textColor=MUTED, spaceAfter=5))
styles.add(ParagraphStyle(name="Callout", parent=styles["BodyText"], fontSize=10, leading=15, textColor=INK, backColor=PALE, borderPadding=9, spaceBefore=5, spaceAfter=10))
styles.add(ParagraphStyle(name="TableHead2", parent=styles["BodyText"], fontName="Helvetica-Bold", fontSize=9, leading=12, textColor=colors.white))


def P(text, style="Body2"):
    return Paragraph(text, styles[style])


def make_table(rows, widths):
    converted = []
    for ri, row in enumerate(rows):
        converted.append([Paragraph(str(cell), styles["TableHead2"] if ri == 0 else styles["Body2"]) for cell in row])
    table = Table(converted, colWidths=widths, repeatRows=1, hAlign="LEFT")
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), GREEN),
        ("BACKGROUND", (0, 1), (-1, -1), colors.white),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, PALE]),
        ("GRID", (0, 0), (-1, -1), .35, colors.HexColor("#dce5dc")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return table


def footer(canvas, doc):
    canvas.saveState()
    w, _ = A4
    canvas.setStrokeColor(colors.HexColor("#e0e7df"))
    canvas.line(18 * mm, 15 * mm, w - 18 * mm, 15 * mm)
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(MUTED)
    canvas.drawString(18 * mm, 10 * mm, "EduPro · Learn the dashboard, learn the data")
    canvas.drawRightString(w - 18 * mm, 10 * mm, f"{doc.page}")
    canvas.restoreState()


doc = SimpleDocTemplate(str(OUTPUT), pagesize=A4, rightMargin=18 * mm, leftMargin=18 * mm, topMargin=18 * mm, bottomMargin=22 * mm, title="EduPro Project Learning Guide", author="EduPro learner project")
story = []
story += [Spacer(1, 32 * mm), P("EDUPRO PROJECT GUIDE", "Small2"), Spacer(1, 8 * mm), P("Let’s understand<br/>the learning dashboard", "CoverTitle"), P("A friendly tour of the app, its data, and the analytics ideas inside it", "CoverSub"), Spacer(1, 16 * mm), P(f"A workbook with {len(users):,} learners, {len(courses)} courses, {len(teachers)} teachers, and {len(tx):,} enrollment records.", "Callout"), Spacer(1, 10 * mm), P("Made to help you explain what the numbers say, how they were calculated, and where their limits are.", "CoverSub"), PageBreak()]

story += [P("1. What did we build?", "Section"), P("Think of the dashboard as a set of windows looking at the same school data. The sidebar lets you choose dates, age groups, genders, categories, levels, and free or paid courses. Every chart responds to those choices. That makes it easier to ask a focused question, like: “For this date range, which category had the most enrollments?”"), P("The app is made with Python and Streamlit. Pandas reads and joins the tables, and Plotly draws the interactive charts. The Excel workbook is the source of truth. When its file changes, the app notices its new modified time and reloads the data."), make_table([["Dashboard area", "What you can learn"], ["Overview", "Enrollment trend, age mix, category demand, and free vs paid activity."], ["Learners", "Who enrolled and how many enrollments each learner has."], ["Courses & revenue", "Popular subjects, course levels, transaction amounts, and catalog ratings."], ["Teachers", "Recorded enrollments and ratings by instructor and expertise."], ["Data quality", "Checks for duplicate IDs, broken links, missing dates, and invalid amounts."], ["Explore & export", "Inspect filtered rows and download CSV summaries." ]], [42 * mm, 132 * mm]), Spacer(1, 8), P("Important: this workbook is a saved snapshot, not a live connection to an online school system. The dashboard can reflect a refreshed workbook, but it cannot show new events that were never added to that file.", "Callout"), P("2. Meet the four tables", "Section"), P("A spreadsheet often stores one kind of thing per table. IDs let us connect those tables without relying on names, which can be misspelled or repeated."), make_table([["Sheet", "One row represents", "Example key"], ["Users", "One learner and their profile", "UserID"], ["Courses", "One course and its details", "CourseID"], ["Teachers", "One teacher and their experience/rating", "TeacherID"], ["Transactions", "One recorded enrollment", "TransactionID" ]], [35 * mm, 86 * mm, 53 * mm]), P("Transactions is the activity table: it has a row for each recorded enrollment. A transaction connects to a learner, course, and teacher through their IDs. Joining means lining up those matching IDs so one enrollment row can also tell us the learner’s age, the course category, and the teacher name."), PageBreak()]

story += [P("3. Data analytics, in plain language", "Section"), P("An analysis usually has a question, a measurement, a grouping, and a careful explanation. Here are the main ideas used in this project."), make_table([["Idea", "Plain meaning", "In this dashboard"], ["Count", "How many records meet a rule?", "Enrollment count is the number of matching transaction rows."], ["Unique count", "How many different people or things?", "Learners means distinct UserIDs, not enrollment rows."], ["Average", "Total divided by number of items.", "Average enrollments per learner = enrollments ÷ distinct learners."], ["Share / percent", "A part divided by the whole, times 100.", "Paid share = paid enrollment rows ÷ all filtered enrollment rows × 100."], ["Group by", "Split records into useful buckets.", "Compare enrollments across age bands or categories."], ["Time trend", "Count records by date bucket.", "Transactions are grouped by month to show recorded activity over time."], ["Revenue sum", "Add the amount recorded for each row.", "Revenue is the sum of Transactions.Amount; units/currency are not specified."], ["Data quality check", "Look for structure that could break a result.", "IDs should be unique and each transaction should find its learner/course/teacher." ]], [37 * mm, 66 * mm, 71 * mm]), P("A percentage always needs a denominator. For example, “30% chose paid courses” means 30% of the filtered enrollment rows are marked Paid. It does not mean 30% of all people in the world or all learners who saw a course." , "Callout"), P("Four words that keep analysis honest", "Sub"), P("Describe: say what is visible in the data. Compare: show how groups differ. Explain: give a reason why it happened. Predict: estimate what may happen later. This workbook supports description and comparison. It does not, by itself, prove causes or reliably predict future behavior."), P("4. How to read the headline numbers", "Section"), make_table([["Metric", "Formula / meaning", "Watch out for"], ["Enrollments", "Count of filtered transaction rows.", "One learner can have more than one row."], ["Learners", "Distinct UserIDs in matching transactions.", "Learners with no matching transaction are not counted here."], ["Average enrollments / learner", "Filtered enrollments ÷ distinct filtered learners.", "An average can hide a mix of one-time and frequent learners."], ["Recorded revenue", "Sum of filtered Amount values.", "No currency is supplied, so use workbook currency units."], ["Leading category", "Category with the most matching rows.", "Popularity is not the same as quality or learning impact." ]], [38 * mm, 72 * mm, 64 * mm]), PageBreak()]

story += [P("5. What this workbook says today", "Section"), P(f"This workbook covers {period}. It contains {len(users):,} learner profiles, {len(courses):,} courses, {len(teachers):,} teacher profiles, and {len(tx):,} enrollment records. The recorded Amount values add to {revenue:,.2f} workbook currency units. There are {paid_n:,} enrollments attached to paid courses. The most frequent category is {top_category}, with {top_category_n:,} enrollment rows. These figures describe the complete workbook before dashboard filters."), P(f"Across the transaction table, there are {unique_learners:,} different learners, or about {mean_per_learner:.2f} transaction rows per learner on average. That average is not a promise that a typical learner takes exactly that many courses; individual learners may be above or below it."), P("The original workbook’s ID fields are unique, its tables have no blank cells, and every transaction points to a learner, course, and teacher. Those are good structural signs. A clean spreadsheet still cannot prove that every real-world enrollment was entered correctly."), P("6. A practical way to use the dashboard", "Section"), make_table([["Step", "Try this"], ["1 · Start broad", "Leave all filters selected and read the date coverage."], ["2 · Ask one question", "For example: which course category had the most activity?"], ["3 · Narrow the view", "Choose a date range or a learner group, then compare the chart."], ["4 · Check the denominator", "See whether the chart counts enrollment rows, learners, revenue, or ratings."], ["5 · Inspect the records", "Use Explore & export to see sample rows behind the summary."], ["6 · Explain with care", "Say what the chart shows, and separate that from your guess about why." ]], [33 * mm, 141 * mm]), P("Example explanation: “In the selected period, Programming had the most recorded enrollments. This tells us it was the most selected category in this workbook view. It does not tell us why learners selected it.” That is a strong data explanation because it is specific, understandable, and does not claim more than the evidence." , "Callout"), P("7. Things to remember", "Section"), P("- Enrollments are transaction rows; they are not the same as unique learners.<br/>- Revenue is recorded transaction amount, not necessarily profit or cash received after refunds.<br/>- Ratings are recorded ratings, not proof of course quality or teacher effectiveness.<br/>- Activity over time can show when enrollments were recorded, but not why they rose or fell.<br/>- A group with more enrollments may simply have more learners or more available courses.<br/>- Filters change the denominator, so always mention the view you are describing."), P("A final mental model", "Sub"), P("Data is a map, not the territory. The workbook gives us a structured view of recorded learning activity. Good analysis reads that map carefully, checks how it was drawn, and explains both what it shows and what it cannot show."), Spacer(1, 7), P("Keep learning: try changing one filter at a time and predict which number should move before you look. Then explain why it moved. That small habit builds strong analytical thinking.", "Callout")]

doc.build(story, onFirstPage=footer, onLaterPages=footer)
print(f"Created {OUTPUT}")
