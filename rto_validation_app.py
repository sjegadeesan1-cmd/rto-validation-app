import streamlit as st
import pandas as pd
import numpy as np
import re
import html
from io import BytesIO

st.title("RTO Exception Validation Tool")

rto_file = st.file_uploader("Upload RTO Automation File", type="xlsx")
plan_file = st.file_uploader("Upload RTO Plan File", type="xlsx")

if rto_file and plan_file:
    rto_automation = pd.read_excel(rto_file, sheet_name="Sheet1", engine="openpyxl")
    rto_automation['Input'] = rto_automation['Input'].astype(str).apply(html.unescape).str.replace(r"\s+", " ", regex=True).str.strip()

    pattern = re.compile(
        r"^(.*?)\((\d+?)\).*?for\s+(.*?)\s+from\s+([\d]{2}-[A-Za-z]{3}-[\d]{4})\s+to\s+([\d]{2}-[A-Za-z]{3}-[\d]{4})\s+on\s+([\d]{2}-[A-Za-z]{3}-[\d]{4})",
        re.IGNORECASE
    )

    def extract_all(text):
        match = pattern.search(text)
        if match:
            return pd.Series({
                "Employee Name": match.group(1).strip(),
                "Employee ID": int(match.group(2)),
                "Category": match.group(3).strip(),
                "From": match.group(4),
                "To": match.group(5),
                "On": match.group(6)
            })
        else:
            return pd.Series({"Employee Name": None, "Employee ID": None, "Category": None, "From": None, "To": None, "On": None})

    extracted_df = rto_automation['Input'].apply(extract_all)

    for col in ['From', 'To', 'On']:
        extracted_df[col] = pd.to_datetime(extracted_df[col], format='%d-%b-%Y', errors='coerce')

    working_days = [np.busday_count(start.date(), end.date() + pd.Timedelta(days=1))
                    if pd.notnull(start) and pd.notnull(end) else None
                    for start, end in zip(extracted_df['From'], extracted_df['To'])]
    extracted_df['#of working days exception raised for'] = working_days

    extracted_df['#Number of request raised'] = extracted_df.groupby('Employee ID')['Employee ID'].transform('count')

    rto_plan = pd.read_excel(plan_file, sheet_name="Sheet1", engine="openpyxl", header=0)
    rto_plan.columns = rto_plan.columns.str.strip()

    if 'Employee ID' in rto_plan.columns:
        rto_plan['Employee ID'] = pd.to_numeric(rto_plan['Employee ID'], errors='coerce')
    else:
        st.error("Column 'Employee ID' not found in RTO Plan file.")

    if 'Depute Branch' in rto_plan.columns:
        depute_branch_map = rto_plan.drop_duplicates('Employee ID').set_index('Employee ID')['Depute Branch'].to_dict()
        extracted_df['Base Branch'] = extracted_df['Employee ID'].map(depute_branch_map)
    else:
        st.error("Column 'Depute Branch' not found in RTO Plan file.")

    wfh_flags = rto_plan.iloc[:, 3:-1].applymap(lambda x: str(x).strip().upper() == 'WFH')
    wfh_counts = wfh_flags.groupby(rto_plan['Employee ID']).sum().sum(axis=1)
    extracted_df['#of days exception given as per RTO roaster'] = extracted_df['Employee ID'].map(wfh_counts).fillna(0).astype(int)

    # ✅ Apply HR approval only to specific Employee IDs
    approved_ids = [2550156, 2549827, 2549950, 2549825, 2549774, 2446423, 2549786]
    extracted_df['Exception approved by HR'] = extracted_df['Employee ID'].apply(
        lambda x: "Exception approved by HR due to medical issues" if x in approved_ids else ""
    )

    remarks = []
    for idx, row in extracted_df.iterrows():
        raised = row['#of working days exception raised for']
        given = row['#of days exception given as per RTO roaster']
        if pd.isna(raised) or pd.isna(given):
            remarks.append("Data incomplete for validation")
        elif raised > given:
            remarks.append(f"{int(raised - given)} day(s) additional exception raised")
        elif raised == given:
            remarks.append("Good to approve")
        else:
            remarks.append(f"{int(given - raised)} day(s) less than roaster")

    extracted_df['RTO validation remarks'] = remarks

    final_df = pd.concat([rto_automation['Input'], extracted_df], axis=1)

    output = BytesIO()
    final_df.to_excel(output, index=False)
    output.seek(0)

    st.download_button("Download Validated Output", output, "RTO_Validation_Output.xlsx")

