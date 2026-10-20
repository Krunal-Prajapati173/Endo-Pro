# Endo-Pro · EduPro Learner Intelligence

An interactive Streamlit dashboard for exploring learner enrollments, course choices, recorded revenue, teacher activity, and workbook quality.

## Run the dashboard

Keep `EduPro_Online_Platform.xlsx` beside `app.py`, then run:

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

The workbook is a saved snapshot. The date filter uses the transaction dates in that file. If you replace or update the workbook, use **Clear cached workbook data** in the sidebar to force a reload.

The public workbook omits learner names and email addresses. The dashboard only needs learner IDs, ages, and genders from the Users sheet. Keep any original identifiable workbook private.

## Upload enrollment data

Use **Upload enrollment data** in the sidebar to load a CSV. EduPro transaction files need learner, course, and teacher IDs, a transaction date, and an amount. The importer recognizes common header spellings such as `StudentID`, `Instructor ID`, and `Enrollment Date`, and detects comma or semicolon separators. Transaction ID and payment method are optional; missing values are filled with generated IDs and `Not provided`. Uploaded IDs must match records in the workbook. A template download is next to the uploader.

Valid transaction uploads can replace the dashboard's transaction rows for the current view without modifying the workbook. Other CSV files get their own dashboard, with a date filter, measure trends or distributions, category breakdowns, summaries, a data quality view, and a filtered download.

## Create the learning guide PDF

```bash
python make_learning_guide.py
```

This creates `EduPro_Project_Learning_Guide.pdf` with a beginner friendly walkthrough of the project and its data analysis concepts.

## Metric notes

- An enrollment is a row in `Transactions`; a learner is a distinct `UserID`.
- Recorded revenue sums `Transactions.Amount`. The workbook does not specify a currency.
- Charts describe recorded activity; they do not establish causes, course outcomes, or teaching effectiveness.
- CSV exports contain filtered transaction rows and a grouped category summary. Learner contact details are not exported.
