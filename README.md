FMLA & Leave Claim Verifier

A Python + Streamlit web app that helps HR/benefits analysts verify hours worked for leave-of-absence requests (FMLA / short-term disability). It reads a pasted leave-request email and a pasted time-evaluation export, totals actual hours for each requested period, checks the boundary dates, and produces a formatted Excel report plus a ready-to-send reply.

    Demo data only. All names, IDs and numbers in this repository are fictional. Do not paste real employee data into a public deployment.

Features

    Request parsing: pulls employee name, ID, leave number, hire/termination dates and one or more requested date ranges out of raw email text with regular expressions.
    Flexible time-log parsing: accepts pipe-delimited (|) or tab-delimited tables copied from an SAP-style time evaluation report.
    Per-period totals: sums hours for every requested 12-month window, plus any hours that fall after the requested end date.
    Boundary check: flags when no shift is logged on the first or last day of the requested period.
    Excel report (openpyxl): formatted attendance table with live =SUM() / =COUNT() formulas and a side panel with period sums, employment history and missing-date alerts.
    Reply draft: generates the email response with the signature name and title you enter (kept in st.session_state, so it resets when the tab closes).

Run locally

pip install -r requirements.txt
streamlit run app.py

Then click Insert Anonymous Example Claim in the app, or paste the contents of sample_data/sample_claim.txt, and press Process Claim & Generate Report.
How it works

    Paste the leave-request email text together with the time-evaluation table.
    The app extracts the requested periods and parses the table into a pandas DataFrame.
    It computes hours per period, checks boundary dates, and builds the Excel workbook in memory.
    Download the report (<EmployeeID> - <LastName>.xlsx) and copy the drafted reply.

Tech stack

Python 3.10+, Streamlit, pandas, openpyxl, re.
Known limitations

    Expects dates as MM/DD/YYYY, MM/DD/YY, or Month D, YYYY.
    If no "X through Y" date range is found, it falls back to placeholder dates and shows a warning.
    Built around one specific export layout (11 columns); other layouts need changes to parse_sap_table.
    No automated tests yet.
