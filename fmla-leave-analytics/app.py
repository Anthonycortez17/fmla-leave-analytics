import io
import re
import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="HRIS Leave Claim Automation | Benefits Analytics",
    page_icon="📊",
    layout="wide"
)

# ==============================================================================
# DYNAMIC DARK/LIGHT ADAPTIVE CSS
# ==============================================================================
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    /* Global Typography */
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif !important;
    }

    /* Executive Header Banner */
    .corporate-header-card {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 50%, #334155 100%);
        padding: 26px 32px;
        border-radius: 16px;
        border-bottom: 5px solid #2563EB;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.25);
        margin-bottom: 25px;
    }

    .corporate-title-text {
        color: #FFFFFF !important;
        font-size: 2.1rem;
        font-weight: 800;
        margin: 0;
        letter-spacing: -0.5px;
        line-height: 1.2;
    }

    .corporate-subtitle-text {
        color: #94A3B8 !important;
        font-size: 1.02rem;
        font-weight: 500;
        margin-top: 6px;
    }

    /* Adaptive Card Containers */
    .card-container {
        background-color: rgba(37, 99, 235, 0.05);
        border: 1px solid rgba(37, 99, 235, 0.2);
        border-left: 6px solid #2563EB;
        padding: 20px 24px;
        border-radius: 12px;
        margin-bottom: 20px;
        backdrop-filter: blur(8px);
    }

    .card-header-title {
        font-size: 1.15rem;
        font-weight: 700;
        margin-bottom: 6px;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    /* Primary Action Button */
    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
        color: #FFFFFF !important;
        font-size: 1.1rem !important;
        font-weight: 700 !important;
        border-radius: 10px !important;
        border: none !important;
        height: 3.4em !important;
        box-shadow: 0 4px 14px rgba(37, 99, 235, 0.35) !important;
        transition: all 0.25s ease-in-out !important;
    }

    .stButton>button:hover {
        background: linear-gradient(135deg, #1D4ED8 0%, #1E40AF 100%) !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px rgba(37, 99, 235, 0.5) !important;
    }

    /* Input Fields Styling */
    .stTextInput input, .stTextArea textarea {
        border-radius: 10px !important;
        font-size: 0.98rem !important;
    }

    /* Expander Container Adaptive Styling */
    [data-testid="stExpander"] {
        border-radius: 12px !important;
        border: 1px solid rgba(37, 99, 235, 0.2) !important;
    }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# HEADER BANNER
# ==============================================================================
st.markdown("""
    <div class="corporate-header-card">
        <div class="corporate-title-text">📊 HRIS FMLA & Leave Claim Automation</div>
        <div class="corporate-subtitle-text">Automated Leave Request & ERP Time Verification Engine | HR Analytics</div>
    </div>
""", unsafe_allow_html=True)

# ==============================================================================
# SESSION USER SETTINGS (Dynamic Analyst Credentials)
# ==============================================================================
st.markdown("""
    <div class="card-container">
        <div class="card-header-title">👤 Analyst Signature Settings (Active Session Only)</div>
        <div style="font-size: 0.9rem; opacity: 0.85;">Configure your credentials below. These remain saved for your current browser tab.</div>
    </div>
""", unsafe_allow_html=True)

col_user_name, col_user_title = st.columns(2)

with col_user_name:
    analyst_name = st.text_input(
        "Your Name (for email signature):",
        value="Analyst Name",
        key="session_analyst_name",
        help="Type your name once. It stays saved as long as this tab is open and resets when closed."
    )

with col_user_title:
    analyst_title = st.text_input(
        "Your Title:",
        value="HRIS Analyst",
        key="session_analyst_title",
        help="Type your job title once. It stays saved for this browser session."
    )

st.markdown("<br>", unsafe_allow_html=True)

# ==============================================================================
# INSTRUCTIONS GUIDE
# ==============================================================================
with st.expander("📖 Step-by-Step Instructions (How to Use This Tool)", expanded=True):
    st.markdown("""
    1. **Copy Leave Request Email:** Open the leave request email from the third-party administrator in your inbox and copy the full text (contains Employee Name, ID, required date range, and employment history).
    2. **Export ERP Time Records:**
       * Open your timekeeping system (e.g., SAP / Workday / Kronos).
       * Run the time evaluation query for the employee across the requested date range.
       * Export or copy the raw tabbed/pipe-delimited table to your clipboard.
    3. **Paste & Process:**
       * Paste **both** the email text AND the copied ERP clipboard table together into the box below.
       * Click **🚀 Process Claim & Generate Report**.
    """)

# ==============================================================================
# ANONYMOUS MOCK DATA
# ==============================================================================
SAMPLE_ANONYMOUS_TEXT = """Employee Name: Jane Doe
Employee ID#: 01234567
Employee Phone #: (555) 019-2831
Leave Number: LV-999999

Hire Date: 04/04/2018 Termination Date: 08/24/2025
Hire Date: 08/25/2025 Termination Date: N/A

Additional information is required to determine eligibility and/or available leave entitlement under STD/FMLA/State law for the above employee. Please reply to this e-mail with the requested information below.

Hours worked:

Total number of actual hours worked during the 12-month period of 8/2/25 through 8/1/26: ___________
Total number of actual hours worked during the 12-month period of 8/3/25 through 8/2/26: ___________
Total number of actual hours worked during the 12-month period of 8/4/25 through 8/3/26: ___________

------------------------------------------------------------------------------------------------------------------------
| Pers.No.|Employee/app.name  |Period|Date      |TmType|TimeTyText  |   Number|Cost Center  |PSubarea |Subarea|Cost Ctr|
------------------------------------------------------------------------------------------------------------------------
|  01234567|DOE JANE          |202508|08/02/2025|0701  |Hours Worked|    8.00 |STORE-001    |Retail   |0101   |10001001|
|  01234567|DOE JANE          |202508|08/03/2025|0701  |Hours Worked|    8.00 |STORE-001    |Retail   |0101   |10001001|
|  01234567|DOE JANE          |202508|08/04/2025|0701  |Hours Worked|    8.00 |STORE-001    |Retail   |0101   |10001001|
|  01234567|DOE JANE          |202608|08/01/2026|0701  |Hours Worked|    8.00 |STORE-001    |Retail   |0101   |10001001|
|  01234567|DOE JANE          |202608|08/02/2026|0701  |Hours Worked|    8.00 |STORE-001    |Retail   |0101   |10001001|
|  01234567|DOE JANE          |202608|08/03/2026|0701  |Hours Worked|    8.00 |STORE-001    |Retail   |0101   |10001001|
------------------------------------------------------------------------------------------------------------------------"""

col_sample, col_clear = st.columns([3, 1])
with col_sample:
    if st.button("📋 Click Here to Insert Anonymous Example Claim"):
        st.session_state["claim_input"] = SAMPLE_ANONYMOUS_TEXT

with col_clear:
    if st.button("🗑️ Clear Text Box"):
        st.session_state["claim_input"] = ""

# Main Text Area
text_input = st.text_area(
    "Paste Combined Request Email + ERP Clipboard Table Here:",
    key="claim_input",
    height=260,
    placeholder="Paste Email Request and ERP Clipboard text here..."
)

include_all_sap_rows = st.checkbox(
    "Include ALL rows in output table (even if dates extend beyond requested end date)",
    value=True,
    help="Check this box if you want the total to match the total at the bottom of the ERP table regardless of the requested end date."
)

# ==============================================================================
# AUTOMATION ENGINE
# ==============================================================================
def parse_sap_table(text):
    lines = text.split('\n')
    data_rows = []
    
    pipe_rows = [l.strip() for l in lines if l.strip().startswith('|') and not l.strip().startswith('|*') and 'Pers.No.' not in l]
    if pipe_rows:
        for line_str in pipe_rows:
            parts = [p.strip() for p in line_str.split('|')[1:-1]]
            if len(parts) >= 11:
                data_rows.append(parts[:11])
        headers = ['Pers.No.', 'Name', 'Period', 'Date', 'TmType', 'TimeTyText', 'Number', 'Cost Ctr', 'PSubarea', 'Subarea', 'Cost Ctr Ref']
        df = pd.DataFrame(data_rows, columns=headers)
    else:
        sap_match = re.search(r"(Pers\.No\..*)", text, re.DOTALL)
        if not sap_match:
            return None
        raw_sap = sap_match.group(1).strip()
        try:
            df = pd.read_csv(io.StringIO(raw_sap), sep=r'\t+|\s{2,}', engine='python')
        except Exception:
            df = pd.read_csv(io.StringIO(raw_sap), sep=None, engine='python')
        headers = [str(col).strip() for col in df.columns[:11]]
        df.columns = headers + list(df.columns[11:])
        df = df.dropna(subset=[headers[0]]).copy()

    df['Number'] = pd.to_numeric(df['Number'].astype(str).str.replace(',', ''), errors='coerce')
    return df

def parse_dt(d_str):
    if not d_str:
        return datetime.date.today()
    d_str = d_str.strip().replace('_', '').strip()
    
    formats = (
        "%m/%d/%Y", "%m/%d/%y",
        "%B %d, %Y", "%B %d %Y", "%b %d, %Y", "%b %d %Y"
    )
    for fmt in formats:
        try:
            return datetime.datetime.strptime(d_str, fmt).date()
        except ValueError:
            pass
    return datetime.date.today()

def process_combined_text(raw_text, sender_name, sender_title, keep_all_sap=True):
    emp_name_match = re.search(r"Employee Name:\s*(.+)", raw_text)
    emp_id_match = re.search(r"Employee ID#:\s*(\d+)", raw_text)
    leave_num_match = re.search(r"Leave Number:\s*(.+)", raw_text)

    # MULTIPLE DATES REGEX: Finds all date range periods requested in the email
    dates_pattern = r"((?:[A-Za-z]+\s+\d{1,2},\s*\d{4})|(?:\d{1,2}/\d{1,2}/\d{2,4}))\s+through\s+((?:[A-Za-z]+\s+\d{1,2},\s*\d{4})|(?:\d{1,2}/\d{1,2}/\d{2,4}))"
    dates_matches = re.findall(dates_pattern, raw_text, re.IGNORECASE)

    requested_periods = []
    for s, e in dates_matches:
        s_dt = parse_dt(s)
        e_dt = parse_dt(e)
        if (s_dt, e_dt) not in requested_periods:
            requested_periods.append((s_dt, e_dt))
            
    if not requested_periods:
        # Fallback so the demo still runs; the user is warned because these dates are a placeholder.
        requested_periods.append((datetime.date(2025, 8, 2), datetime.date(2026, 8, 5)))
        st.warning("No 'X through Y' date range found in the request text. Using placeholder dates 08/02/2025 - 08/05/2026; verify before sending.")

    global_req_start = min([p[0] for p in requested_periods])
    global_req_end = max([p[1] for p in requested_periods])

    # EXTRACT HIRE / TERMINATION DATES
    employment_records = []
    emp_history_matches = re.findall(
        r"Hire\s*Date:\s*([^\s_]+(?:\s+\d{1,2},\s*\d{4})?)\s*Termination\s*Date:\s*([^\s_]+(?:\s+\d{1,2},\s*\d{4})?)",
        raw_text,
        re.IGNORECASE
    )
    for h_date, t_date in emp_history_matches:
        h_clean = h_date.strip().replace('_', '')
        t_clean = t_date.strip().replace('_', '')
        if h_clean:
            employment_records.append(f"Hire Date: {h_clean} | Term Date: {t_clean}")

    emp_name = emp_name_match.group(1).strip() if emp_name_match else "Employee"
    emp_id = emp_id_match.group(1).strip() if emp_id_match else "000000"
    leave_num = leave_num_match.group(1).strip() if leave_num_match else "N/A"

    clean_emp_id = str(emp_id).lstrip('0')
    last_name = emp_name.split()[-1]
    output_filename = f"{clean_emp_id} - {last_name}.xlsx"

    df = parse_sap_table(raw_text)
    if df is None or df.empty:
        st.error("Could not parse ERP table. Ensure table rows start with '|' pipe symbols or standard tab formatting.")
        return None, None, None, [], 0, 0

    df['Date_dt'] = pd.to_datetime(df['Date']).dt.date
    df = df.sort_values(by='Date_dt')

    # DUAL SUM CALCULATIONS FOR EACH PERIOD
    period_sums = []
    for s_dt, e_dt in requested_periods:
        p_df = df[(df['Date_dt'] >= s_dt) & (df['Date_dt'] <= e_dt)]
        period_sums.append((s_dt, e_dt, p_df['Number'].sum()))

    df_req_period = df[(df['Date_dt'] >= global_req_start) & (df['Date_dt'] <= global_req_end)]
    df_extra_period = df[df['Date_dt'] > global_req_end]

    global_requested_sum = df_req_period['Number'].sum()
    extra_sum = df_extra_period['Number'].sum()
    total_sap_sum = df['Number'].sum()

    # AUDIT CHECK: MISSING BOUNDARIES
    missing_dates = []
    if global_req_start not in df_req_period['Date_dt'].values:
        missing_dates.append(global_req_start.strftime("%m/%d/%Y"))
    if global_req_end not in df_req_period['Date_dt'].values:
        missing_dates.append(global_req_end.strftime("%m/%d/%Y"))

    # Build Excel Report
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "FMLA Hours Log"
    ws_out.views.sheetView[0].showGridLines = True

    NAVY_PRIMARY = "0F172A"
    ZEBRA_FILL = PatternFill(start_color="F8FAFC", fill_type="solid")
    WHITE_FILL = PatternFill(start_color="FFFFFF", fill_type="solid")
    GOLD_TOTAL_FILL = PatternFill(start_color="FEF08A", fill_type="solid")
    CARD_BG_FILL = PatternFill(start_color="F1F5F9", fill_type="solid")
    ALERT_BG_FILL = PatternFill(start_color="FEF2F2", fill_type="solid")

    FONT_BANNER = Font(name="Segoe UI", size=16, bold=True, color="FFFFFF")
    FONT_HEADER = Font(name="Segoe UI", size=12, bold=True, color="FFFFFF")
    FONT_BODY = Font(name="Segoe UI", size=11, color="1E293B")
    FONT_BOLD = Font(name="Segoe UI", size=11, bold=True, color="0F172A")
    FONT_TOTAL = Font(name="Segoe UI", size=12, bold=True, color="0F172A")
    FONT_KPI_VAL = Font(name="Segoe UI", size=26, bold=True, color="1E3A8A")
    FONT_KPI_LBL = Font(name="Segoe UI", size=9, bold=True, color="64748B")
    FONT_ALERT_TITLE = Font(name="Segoe UI", size=11, bold=True, color="991B1B")
    FONT_ALERT_BODY = Font(name="Segoe UI", size=11, bold=True, italic=True, color="B91C1C")

    BORDER_SUBTLE = Side(border_style="thin", color="E2E8F0")
    BORDER_GRID = Border(left=BORDER_SUBTLE, right=BORDER_SUBTLE, top=BORDER_SUBTLE, bottom=BORDER_SUBTLE)
    BORDER_TOTAL = Border(top=Side(border_style="medium", color="0F172A"), bottom=Side(border_style="double", color="0F172A"), left=BORDER_SUBTLE, right=BORDER_SUBTLE)

    ws_out.merge_cells("A1:K2")
    ws_out["A1"] = "  FMLA / STD HOURS WORKED VERIFICATION REPORT"
    ws_out["A1"].font = FONT_BANNER
    ws_out["A1"].fill = PatternFill(start_color=NAVY_PRIMARY, fill_type="solid")
    ws_out["A1"].alignment = Alignment(horizontal="left", vertical="center")

    ws_out["A3"] = "  Official Attendance & Shift Log | HR & Benefits Analytics"
    ws_out["A3"].font = Font(name="Segoe UI", size=10, italic=True, color="64748B")

    start_table_row = 5
    table_headers = ['Pers.No.', 'Name', 'Period', 'Date', 'TmType', 'TimeTyText', 'Number', 'Cost Ctr', 'PSubarea', 'Subarea', 'Cost Ctr Ref']
    for col_i, h_txt in enumerate(table_headers, start=1):
        cell = ws_out.cell(row=start_table_row, column=col_i, value=h_txt)
        cell.fill = PatternFill(start_color=NAVY_PRIMARY, fill_type="solid")
        cell.font = FONT_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER_GRID

    ws_out.row_dimensions[start_table_row].height = 32

    export_df = df if keep_all_sap else df_req_period
    filtered_vals = export_df.iloc[:, :11].values

    for row_idx, r_vals in enumerate(filtered_vals, start=start_table_row + 1):
        ws_out.row_dimensions[row_idx].height = 22
        row_fill = ZEBRA_FILL if row_idx % 2 == 0 else WHITE_FILL
        for col_i in range(1, 12):
            val = r_vals[col_i - 1]
            if col_i in [1, 3, 5, 10, 11]:
                try: val = int(float(val))
                except: pass
            elif col_i == 7:
                try: val = float(val)
                except: pass

            cell = ws_out.cell(row=row_idx, column=col_i, value=val)
            cell.font = FONT_BODY
            cell.fill = row_fill
            cell.border = BORDER_GRID

            if col_i in [1, 3, 5, 10, 11]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_i == 4:
                cell.number_format = "MM/DD/YYYY"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_i == 7:
                cell.number_format = "#,##0.00"
                cell.alignment = Alignment(horizontal="right", vertical="center")

    total_row = start_table_row + len(filtered_vals) + 1
    ws_out.row_dimensions[total_row].height = 26
    ws_out.cell(row=total_row, column=6, value="Grand Total (Full Output Table):").font = FONT_TOTAL
    ws_out.cell(row=total_row, column=6).alignment = Alignment(horizontal="right", vertical="center")
    ws_out.cell(row=total_row, column=6).border = BORDER_TOTAL

    sum_cell = ws_out.cell(row=total_row, column=7, value=f"=SUM(G6:G{total_row-1})")
    sum_cell.font = FONT_TOTAL
    sum_cell.number_format = "#,##0.00"
    sum_cell.fill = GOLD_TOTAL_FILL
    sum_cell.alignment = Alignment(horizontal="right", vertical="center")
    sum_cell.border = BORDER_TOTAL

    for c_idx in range(1, 12):
        if c_idx not in [6, 7]:
            ws_out.cell(row=total_row, column=c_idx).border = Border(top=Side(border_style="medium", color="0F172A"), bottom=Side(border_style="thin", color="0F172A"))

    # Missing Notes Under Table
    missing_note_row = (total_row - 1) + 5
    for i, m_date in enumerate(missing_dates):
        ws_out.cell(row=missing_note_row + i, column=1, value=f"Employee did not work {m_date}").font = FONT_ALERT_BODY

    # Metrics Card
    ws_out.merge_cells("M1:P2")
    ws_out["M1"] = "CLAIM AUDIT & METRICS CARD"
    ws_out["M1"].font = FONT_BANNER
    ws_out["M1"].fill = PatternFill(start_color=NAVY_PRIMARY, fill_type="solid")
    ws_out["M1"].alignment = Alignment(horizontal="center", vertical="center")

    info_card = [
        ("Employee Name:", emp_name),
        ("Employee ID#:", str(clean_emp_id)),
        ("Leave Number:", str(leave_num)),
    ]

    for idx, (lbl, val) in enumerate(info_card, start=5):
        ws_out.cell(row=idx, column=13, value=lbl).font = Font(name="Segoe UI", size=10, bold=True, color="64748B")
        ws_out.merge_cells(start_row=idx, start_column=14, end_row=idx, end_column=16)
        ws_out.cell(row=idx, column=14, value=val).font = FONT_BOLD
        for c_idx in range(13, 17):
            ws_out.cell(row=idx, column=c_idx).fill = CARD_BG_FILL
            ws_out.cell(row=idx, column=c_idx).border = BORDER_GRID

    # Add Requested Periods & Corresponding Sums Side-by-Side on the Metrics Card
    curr_r = 8
    ws_out.cell(row=curr_r, column=13, value="Requested Periods & Sums:").font = Font(name="Segoe UI", size=10, bold=True, color="64748B")
    curr_r += 1

    for s_dt, e_dt, p_sum in period_sums:
        p_str = f"{s_dt.strftime('%m/%d/%Y')} – {e_dt.strftime('%m/%d/%Y')}:"
        ws_out.cell(row=curr_r, column=13, value=p_str).font = Font(name="Segoe UI", size=10, bold=True, color="0F172A")
        ws_out.merge_cells(start_row=curr_r, start_column=14, end_row=curr_r, end_column=16)
        ws_out.cell(row=curr_r, column=14, value=f"{p_sum:,.2f} hrs").font = FONT_BOLD
        for c_idx in range(13, 17):
            ws_out.cell(row=curr_r, column=c_idx).fill = CARD_BG_FILL
            ws_out.cell(row=curr_r, column=c_idx).border = BORDER_GRID
        curr_r += 1

    # Add Employment History Records
    if employment_records:
        ws_out.cell(row=curr_r, column=13, value="Employment History:").font = Font(name="Segoe UI", size=10, bold=True, color="64748B")
        for rec in employment_records:
            ws_out.merge_cells(start_row=curr_r, start_column=14, end_row=curr_r, end_column=16)
            ws_out.cell(row=curr_r, column=14, value=rec).font = FONT_BOLD
            for c_idx in range(13, 17):
                ws_out.cell(row=curr_r, column=c_idx).fill = CARD_BG_FILL
                ws_out.cell(row=curr_r, column=c_idx).border = BORDER_GRID
            curr_r += 1

    ws_out.merge_cells(f"M{curr_r+1}:N{curr_r+1}")
    ws_out[f"M{curr_r+1}"] = "GRAND TOTAL (FULL EXPORT)"
    ws_out[f"M{curr_r+1}"].font = FONT_KPI_LBL
    ws_out[f"M{curr_r+1}"].alignment = Alignment(horizontal="center")

    ws_out.merge_cells(f"M{curr_r+2}:N{curr_r+4}")
    ws_out[f"M{curr_r+2}"] = f"=G{total_row}"
    ws_out[f"M{curr_r+2}"].font = FONT_KPI_VAL
    ws_out[f"M{curr_r+2}"].number_format = "#,##0.00"
    ws_out[f"M{curr_r+2}"].alignment = Alignment(horizontal="center", vertical="center")

    ws_out.merge_cells(f"O{curr_r+1}:P{curr_r+1}")
    ws_out[f"O{curr_r+1}"] = "TOTAL SHIFTS LOGGED"
    ws_out[f"O{curr_r+1}"].font = FONT_KPI_LBL
    ws_out[f"O{curr_r+1}"].alignment = Alignment(horizontal="center")

    ws_out.merge_cells(f"O{curr_r+2}:P{curr_r+4}")
    ws_out[f"O{curr_r+2}"] = f"=COUNT(G6:G{total_row-1})"
    ws_out[f"O{curr_r+2}"].font = FONT_KPI_VAL
    ws_out[f"O{curr_r+2}"].alignment = Alignment(horizontal="center", vertical="center")

    for r in range(curr_r+1, curr_r+5):
        for c_idx in range(13, 17):
            cell = ws_out.cell(row=r, column=c_idx)
            cell.fill = PatternFill(start_color="F8FAFC", fill_type="solid")
            cell.border = BORDER_GRID

    ws_out.merge_cells(f"M{curr_r+6}:P{curr_r+6}")
    ws_out[f"M{curr_r+6}"] = "MISSING BOUNDARY DATES"
    ws_out[f"M{curr_r+6}"].font = FONT_ALERT_TITLE
    ws_out[f"M{curr_r+6}"].fill = ALERT_BG_FILL
    ws_out[f"M{curr_r+6}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    all_exceptions = [f"Employee did not work {md}" for md in missing_dates]
    if all_exceptions:
        for i, exc in enumerate(all_exceptions):
            r_idx = curr_r + 7 + i
            ws_out.merge_cells(start_row=r_idx, start_column=13, end_row=r_idx, end_column=16)
            ws_out.cell(row=r_idx, column=13, value=exc).font = FONT_ALERT_BODY
    else:
        ws_out.merge_cells(f"M{curr_r+7}:P{curr_r+7}")
        ws_out[f"M{curr_r+7}"] = "No missing boundary dates detected."
        ws_out[f"M{curr_r+7}"].font = FONT_BOLD

    ws_out.column_dimensions['A'].width = 38
    ws_out.column_dimensions['B'].width = 28
    ws_out.column_dimensions['C'].width = 14
    ws_out.column_dimensions['D'].width = 16
    ws_out.column_dimensions['E'].width = 14
    ws_out.column_dimensions['F'].width = 17
    ws_out.column_dimensions['G'].width = 18
    ws_out.column_dimensions['H'].width = 18
    ws_out.column_dimensions['I'].width = 14
    ws_out.column_dimensions['J'].width = 14
    ws_out.column_dimensions['K'].width = 17
    ws_out.column_dimensions['L'].width = 4
    ws_out.column_dimensions['M'].width = 30
    ws_out.column_dimensions['N'].width = 22
    ws_out.column_dimensions['O'].width = 28
    ws_out.column_dimensions['P'].width = 28

    excel_buffer = io.BytesIO()
    wb_out.save(excel_buffer)
    excel_buffer.seek(0)

    # DYNAMIC EMAIL RESPONSE
    notes_str = f"\n(Note: Employee did not work on {' and '.join(missing_dates)})\n" if missing_dates else "\n"
    history_str = f"\nEmployment History Record:\n" + "\n".join([f"• {rec}" for rec in employment_records]) + "\n" if employment_records else ""

    hours_breakdown_str = ""
    for i, (s_dt, e_dt, p_sum) in enumerate(period_sums):
        hours_breakdown_str += f"• {p_sum:,.2f} hours worked during period {s_dt.strftime('%m/%d/%Y')} through {e_dt.strftime('%m/%d/%Y')}\n"
    
    if extra_sum > 0:
        extra_start_date = (global_req_end + datetime.timedelta(days=1)).strftime('%m/%d/%Y')
        max_sap_date = df['Date_dt'].max().strftime('%m/%d/%Y')
        hours_breakdown_str += f"• {extra_sum:,.2f} additional hours worked after requested period ({extra_start_date} through {max_sap_date})\n"
    
    if len(period_sums) > 1 or extra_sum > 0:
        hours_breakdown_str += f"• Total hours in attached ERP report: {total_sap_sum:,.2f} hours\n"

    email_response = f"""Per your request. Please see attached report to determine eligible hours. 

{hours_breakdown_str}{notes_str}{history_str}
Direct any questions to your supervisor.

Thank you,

{sender_name}
{sender_title}"""

    return email_response, excel_buffer, output_filename, period_sums, extra_sum, total_sap_sum

# Process Action
if st.button("🚀 Process Claim & Generate Report"):
    if not text_input.strip():
        st.warning("Please paste text before clicking process.")
    else:
        email_txt, excel_data, filename, period_sums, extra_sum, total_sum = process_combined_text(
            text_input,
            analyst_name.strip() if analyst_name.strip() else "Analyst Name",
            analyst_title.strip() if analyst_title.strip() else "HRIS Analyst",
            keep_all_sap=include_all_sap_rows
        )
        if email_txt:
            st.success("Claim audit complete!")

            # Display Dynamic Sum Breakdown Card in Web UI
            st.markdown("### 📊 Calculated Hours Breakdown")
            
            num_cols = len(period_sums) + (1 if extra_sum > 0 else 0) + 1
            cols = st.columns(num_cols)
            
            for i, (s_dt, e_dt, p_val) in enumerate(period_sums):
                with cols[i]:
                    st.metric(f"Period ({s_dt.strftime('%m/%d/%y')}–{e_dt.strftime('%m/%d/%y')})", f"{p_val:,.2f} hrs")
            
            col_idx = len(period_sums)
            if extra_sum > 0:
                with cols[col_idx]:
                    st.metric("Post-Period Extra Hours", f"{extra_sum:,.2f} hrs")
                col_idx += 1

            with cols[col_idx]:
                st.metric("Grand Total (Full Output Table)", f"{total_sum:,.2f} hrs")

            st.subheader("✉️ Ready-to-Send Email Reply")
            st.code(email_txt, language="text")
            
            st.subheader("📥 Executive Excel Report")
            st.download_button(
                label=f"Download {filename}",
                data=excel_data,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )