# Import necessary libraries
import streamlit as st  # For building the web app interface
import pandas as pd     # For handling Excel and tabular data
import numpy as np      # For numerical operations like working day calculations
import re               # For pattern matching using regular expressions
import html             # For decoding HTML entities
from io import BytesIO  # For creating downloadable Excel output

# Title of the web app
st.title("RTO Exception Validation Tool")

# Upload two Excel files: one for RTO Automation and one for RTO Plan
rto_file = st.file_uploader("Upload RTO Automation File", type="xlsx")
plan_file = st.file_uploader("Upload RTO Plan File", type="xlsx")

# Proceed only if both files are uploaded
if rto_file and plan_file:
    # Read the RTO Automation file
    rto_automation = pd.read_excel(rto_file, sheet_name="Sheet1", engine="openpyxl")

    # Clean the 'Input' column: decode HTML, remove extra spaces, and strip whitespace
    rto_automation['Input'] = rto_automation['Input'].astype(str).apply(html.unescape).str.replace(r"\s+", " ", regex=True).str.strip()

    # Define a pattern to extract details like Employee Name, ID, Category, Dates
    pattern = re.compile(
        r"^(.*?)\((\d+?)\).*?for\s+(.*?)\s+from\s+([\d]{2}-[A-Za-z]{3}-[\d]{4})\s+to\s+([\d]{2}-[A-Za-z]{3}-[\d]{4})\s+on\s+([\d]{2}-[A-Za-z]{3}-[\d]{4})",
        re.IGNORECASE
    )

    # Function to extract data from each 'Input' line using the pattern
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
            # If pattern doesn't match, return empty values
            return pd.Series({"Employee Name": None, "Employee ID": None, "Category": None, "From": None, "To": None, "On": None})

    # Apply the extraction function to each row
    extracted_df = rto_automation['Input'].apply(extract_all)

    # Convert date strings to actual datetime objects
    for col in ['From', 'To', 'On']:
        extracted_df[col] = pd.to_datetime(extracted_df[col], format='%d-%b-%Y', errors='coerce')

    # Calculate number of working days between 'From' and 'To' dates
    working_days = [np.busday_count(start.date(), end.date() + pd.Timedelta(days=1))
                    if pd.notnull(start) and pd.notnull(end) else None
                    for start, end in zip(extracted_df['From'], extracted_df['To'])]
    extracted_df['#of working days exception raised for'] = working_days
    # Count how many requests each employee has raised
    extracted_df['#Number of request raised'] = extracted_df.groupby('Employee ID')['Employee ID'].transform('count')

    # Read the RTO Plan file
    rto_plan = pd.read_excel(plan_file, sheet_name="Sheet1", engine="openpyxl", header=0)
    rto_plan.columns = rto_plan.columns.str.strip()  # Clean column names

    # Ensure 'Employee ID' column is numeric
    if 'Employee ID' in rto_plan.columns:
        rto_plan['Employee ID'] = pd.to_numeric(rto_plan['Employee ID'], errors='coerce')
    else:
        st.error("Column 'Employee ID' not found in RTO Plan file.")

    # Map each employee to their base branch using 'Depute Branch'
    if 'Depute Branch' in rto_plan.columns:
        depute_branch_map = rto_plan.drop_duplicates('Employee ID').set_index('Employee ID')['Depute Branch'].to_dict()
        extracted_df['Base Branch'] = extracted_df['Employee ID'].map(depute_branch_map)
    else:
        st.error("Column 'Depute Branch' not found in RTO Plan file.")

    # Identify which days are marked as 'WFH' (Work From Home)
    # This checks each cell in the daily columns and marks True if it's 'WFH'
    wfh_flags = rto_plan.iloc[:, 3:-1].applymap(lambda x: str(x).strip().upper() == 'WFH')

    # Count how many WFH days each employee has in the plan
    wfh_counts = wfh_flags.groupby(rto_plan['Employee ID']).sum().sum(axis=1)

    # Map WFH counts to the extracted data
    extracted_df['#of days exception given as per RTO roaster'] = extracted_df['Employee ID'].map(wfh_counts).fillna(0).astype(int)

    # Apply HR approval only to specific Employee IDs
    approved_ids = [2550156, 2549827, 2549950, 2549825, 2549774, 2446423, 2549786, 2549996, 2550111]
    extracted_df['Exception approved by HR'] = extracted_df['Employee ID'].apply(
        lambda x: "Exception approved by HR due to medical issues" if x in approved_ids else ""
    )

    # Generate validation remarks based on comparison between raised and allowed exceptions
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

    # Add remarks to the final DataFrame
    extracted_df['RTO validation remarks'] = remarks

    # Combine original input with extracted and validated data
    final_df = pd.concat([rto_automation['Input'], extracted_df], axis=1)

    # Prepare the final output file for download
    output = BytesIO()
    final_df.to_excel(output, index=False)
    output.seek(0)

    # Provide a download button for the validated Excel file
    st.download_button("Download Validated Output", output, "RTO_Validation_Output.xlsx")
