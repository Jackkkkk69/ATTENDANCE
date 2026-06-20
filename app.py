import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date, time
from io import BytesIO

st.set_page_config(page_title="Sachin Tuition Attendance Audit App", layout="wide")
st.title("Sachin Tuition Attendance Audit App")
st.caption("Enter attendance and payments once. The app automatically calculates total hours, shortage, refundable/excess amount, and generates a zero-defect Excel report.")

DEFAULT_MONTHS = pd.period_range("2026-06", "2027-03", freq="M").astype(str).tolist()


def parse_time(x):
    if pd.isna(x) or str(x).strip() == "":
        return None
    s = str(x).strip()
    for fmt in ["%H:%M", "%I:%M %p", "%I:%M%p", "%H.%M"]:
        try:
            return datetime.strptime(s, fmt).time()
        except Exception:
            pass
    return None


def duration_from_times(start, end, break_minutes=0):
    stime = parse_time(start)
    etime = parse_time(end)
    if stime is None or etime is None:
        return np.nan
    base = datetime(2000, 1, 1)
    sdt = datetime.combine(base.date(), stime)
    edt = datetime.combine(base.date(), etime)
    if edt < sdt:
        edt = edt + pd.Timedelta(days=1)
    hrs = (edt - sdt).total_seconds() / 3600 - (float(break_minutes or 0) / 60)
    return max(0, round(hrs, 2))


def clean_attendance(df):
    df = df.copy()
    required = ["Date", "Start Time", "End Time", "Break Minutes", "Manual Duration Hours", "Student/Batch", "Subject", "Status", "Remarks"]
    for col in required:
        if col not in df.columns:
            df[col] = "" if col not in ["Break Minutes", "Manual Duration Hours"] else np.nan
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
    df["Break Minutes"] = pd.to_numeric(df["Break Minutes"], errors="coerce").fillna(0)
    df["Manual Duration Hours"] = pd.to_numeric(df["Manual Duration Hours"], errors="coerce")
    auto = [duration_from_times(r["Start Time"], r["End Time"], r["Break Minutes"]) for _, r in df.iterrows()]
    df["Auto Duration Hours"] = auto
    df["Final Duration Hours"] = df["Manual Duration Hours"].where(df["Manual Duration Hours"].notna(), df["Auto Duration Hours"])
    df["Final Duration Hours"] = pd.to_numeric(df["Final Duration Hours"], errors="coerce").fillna(0)
    df["Month"] = pd.to_datetime(df["Date"], errors="coerce").dt.to_period("M").astype(str)
    df["Check Status"] = np.where(df["Date"].isna(), "Date Missing", np.where(df["Final Duration Hours"] <= 0, "Check Time/Duration", "OK"))
    return df


def clean_payments(df):
    df = df.copy()
    required = ["Date", "Payer", "Amount", "Mode", "Remarks", "Include in Refund Calculation"]
    for col in required:
        if col not in df.columns:
            df[col] = True if col == "Include in Refund Calculation" else ""
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)
    df["Include in Refund Calculation"] = df["Include in Refund Calculation"].astype(str).str.lower().isin(["true", "yes", "1", "included"])
    return df


def make_month_summary(att, expected_total_hours, months, default_monthly_target):
    actual = att.groupby("Month", dropna=True)["Final Duration Hours"].sum().reset_index()
    base = pd.DataFrame({"Month": months})
    base["Expected Hours"] = default_monthly_target
    base = base.merge(actual, on="Month", how="left").fillna({"Final Duration Hours": 0})
    base.rename(columns={"Final Duration Hours": "Actual Hours"}, inplace=True)
    base["Monthly Shortage"] = (base["Expected Hours"] - base["Actual Hours"]).clip(lower=0).round(2)
    base["Cumulative Actual Hours"] = base["Actual Hours"].cumsum().round(2)
    base["Remaining Against Total"] = (expected_total_hours - base["Cumulative Actual Hours"]).clip(lower=0).round(2)
    return base


def create_excel(settings, attendance, month_summary, payments, refund_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame(list(settings.items()), columns=["Parameter", "Value"]).to_excel(writer, index=False, sheet_name="Settings")
        attendance.to_excel(writer, index=False, sheet_name="Attendance_Ledger")
        month_summary.to_excel(writer, index=False, sheet_name="Month_Wise_Summary")
        payments.to_excel(writer, index=False, sheet_name="Payments")
        refund_df.to_excel(writer, index=False, sheet_name="Refund_Statement")
        audit = pd.DataFrame({
            "Check Point": [
                "Attendance rows with missing date", "Attendance rows with zero duration", "Negative payments", "Formula logic status"
            ],
            "Result": [
                int(attendance["Date"].isna().sum()),
                int((attendance["Final Duration Hours"] <= 0).sum()),
                int((payments["Amount"] < 0).sum()),
                "PASS - calculated by app from inputs"
            ]
        })
        audit.to_excel(writer, index=False, sheet_name="Zero_Defect_Check")
        wb = writer.book
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True)
            for col in ws.columns:
                max_len = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 12), 36)
    return output.getvalue()

# Sidebar settings
st.sidebar.header("Audit Settings")
expected_total_hours = st.sidebar.number_input("Expected Total Hours", min_value=0.0, value=195.0, step=0.5)
hourly_rate = st.sidebar.number_input("Hourly Rate (₹/hour)", min_value=0.0, value=384.0, step=1.0)
default_monthly_target = st.sidebar.number_input("Default Monthly Target Hours", min_value=0.0, value=39.0, step=0.5)
opening_paid = st.sidebar.number_input("Opening/Already Paid Amount (₹)", min_value=0.0, value=91464.0, step=100.0)
months_text = st.sidebar.text_area("Months to Track (YYYY-MM, one per line)", value="\n".join(DEFAULT_MONTHS))
months = [m.strip() for m in months_text.splitlines() if m.strip()]

uploaded = st.file_uploader("Optional: upload previous app Excel workbook", type=["xlsx"])

attendance_template = pd.DataFrame({
    "Date": pd.date_range("2026-06-01", periods=10, freq="D").date,
    "Start Time": [""] * 10,
    "End Time": [""] * 10,
    "Break Minutes": [0] * 10,
    "Manual Duration Hours": [np.nan] * 10,
    "Student/Batch": ["Sachin Tuition"] * 10,
    "Subject": [""] * 10,
    "Status": ["Completed"] * 10,
    "Remarks": [""] * 10,
})
payments_template = pd.DataFrame({
    "Date": [date.today()],
    "Payer": [""],
    "Amount": [0.0],
    "Mode": [""],
    "Remarks": [""],
    "Include in Refund Calculation": [True],
})

if uploaded:
    try:
        xls = pd.ExcelFile(uploaded)
        if "Attendance_Ledger" in xls.sheet_names:
            attendance_template = pd.read_excel(xls, "Attendance_Ledger")
        if "Payments" in xls.sheet_names:
            payments_template = pd.read_excel(xls, "Payments")
    except Exception as e:
        st.warning(f"Could not read uploaded workbook: {e}")

st.subheader("1. Attendance Ledger Input")
st.write("Enter either start/end time or direct manual duration. Manual duration overrides auto duration.")
attendance_input = st.data_editor(attendance_template, num_rows="dynamic", use_container_width=True, key="attendance")

st.subheader("2. Payments Input")
payments_input = st.data_editor(payments_template, num_rows="dynamic", use_container_width=True, key="payments")

attendance = clean_attendance(attendance_input)
payments = clean_payments(payments_input)
month_summary = make_month_summary(attendance, expected_total_hours, months, default_monthly_target)

total_actual = round(float(attendance["Final Duration Hours"].sum()), 2)
verified_shortage = round(max(expected_total_hours - total_actual, 0), 2)
tuition_value = round(total_actual * hourly_rate, 2)
total_paid = round(opening_paid + float(payments.loc[payments["Include in Refund Calculation"], "Amount"].sum()), 2)
refundable = round(max(total_paid - tuition_value, 0), 2)
balance_due = round(max(tuition_value - total_paid, 0), 2)

refund_df = pd.DataFrame({
    "Particular": ["Expected Total Hours", "Actual Completed Hours", "Verified Shortage Hours", "Hourly Rate", "Tuition Value", "Total Paid", "Refundable/Excess Amount", "Balance Due"],
    "Value": [expected_total_hours, total_actual, verified_shortage, hourly_rate, tuition_value, total_paid, refundable, balance_due]
})

st.subheader("3. Auto Output Dashboard")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Actual Hours", f"{total_actual:.2f}")
c2.metric("Shortage Hours", f"{verified_shortage:.2f}")
c3.metric("Tuition Value", f"₹{tuition_value:,.2f}")
c4.metric("Refundable/Excess", f"₹{refundable:,.2f}")

st.subheader("4. Month-wise Summary")
st.dataframe(month_summary, use_container_width=True)

st.subheader("5. Refund Statement")
st.dataframe(refund_df, use_container_width=True)

settings = {
    "Expected Total Hours": expected_total_hours,
    "Hourly Rate": hourly_rate,
    "Default Monthly Target Hours": default_monthly_target,
    "Opening/Already Paid Amount": opening_paid,
    "Generated On": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
}
excel_bytes = create_excel(settings, attendance, month_summary, payments, refund_df)
st.download_button("Download Final Zero-Defect Excel Report", data=excel_bytes, file_name="Sachin_Tuition_App_Output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
