import streamlit as st
import pandas as pd
import unicodedata
from io import BytesIO

st.set_page_config(page_title="Vietnam Address Validation Tool", layout="wide")

# Remove Vietnamese tone marks
def remove_tones(text):
    if not isinstance(text, str):
        return text
    text = unicodedata.normalize('NFD', text)
    return ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')

# Normalize string columns (lowercase, strip spaces, remove tones)
def normalize_col(col):
    return col.astype(str).str.lower().str.strip().apply(remove_tones)

# Load and preprocess files
def load_file(file, form=True):
    df = pd.read_excel(file)
    df.columns = df.columns.str.strip()
    if form:
        df['Account Number'] = df['Account Number'].astype(str).str.strip()
        df['New Address Line 1'] = normalize_col(df['New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only'])
        df['New Address Line 2'] = normalize_col(df['New Address Line 2 (Street Name)-In English Only'])
    else:
        df['AC_NUM'] = df['AC_NUM'].astype(str).str.strip()
        df['Address_Line1'] = normalize_col(df['Address_Line1'])
        df['Address_Line2'] = normalize_col(df['Address_Line2'])
    return df

# Match logic
def match_addresses(form_df, ups_df):
    matched = []
    unmatched = []

    for _, row in form_df.iterrows():
        account = row['Account Number']
        new_line1 = row['New Address Line 1']
        new_line2 = row['New Address Line 2']
        billing_same = row.get('Is Your New Billing Address the Same as Your Pickup and Delivery Address?', '').strip().lower()
        matched_rows = ups_df[ups_df['AC_NUM'] == account]

        if matched_rows.empty:
            row['Unmatched Reason'] = 'Account not found in UPS file'
            unmatched.append(row)
            continue

        found = False
        for _, old_row in matched_rows.iterrows():
            old_line1 = old_row['Address_Line1']
            old_line2 = old_row['Address_Line2']

            # Must match both address no. and street, case-insensitive and no tones
            if new_line1 in old_line1 and new_line2 in old_line2:
                match_info = {
                    'AC_NUM': account,
                    'AC_Address_Type': '01' if billing_same == 'yes' else '',
                    'AC_Name': old_row['AC_Name'],
                    'Address_Line1': remove_tones(row['New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only']),
                    'Address_Line2': remove_tones(row['New Address Line 2 (Street Name)-In English Only']),
                    'City': row['City / Province'],
                    'Postal_Code': '',
                    'Country_Code': 'VN',
                    'Attention_Name': row.get('Full Name of Contact-In English Only', ''),
                    'Address_Line22': row.get('New Address Line 3 (Ward/Commune)-In English Only', ''),
                    'Address_Country_Code': 'VN'
                }
                matched.append(match_info)
                found = True
                break

        if not found:
            row['Unmatched Reason'] = 'Address not similar enough to UPS record'
            unmatched.append(row)

    return pd.DataFrame(matched), pd.DataFrame(unmatched)

# Expand matched data to upload format
def expand_matched_for_upload(matched_df):
    expanded = []

    for _, row in matched_df.iterrows():
        acct = row['AC_NUM']
        if row['AC_Address_Type'] == '01':
            for code in ['1', '2', '6']:
                new_row = row.copy()
                new_row['AC_Address_Type'] = code
                expanded.append(new_row)
        else:
            row['AC_Address_Type'] = '02'  # pickup
            expanded.append(row)

    return pd.DataFrame(expanded)

# File export helper
def to_excel_file(dfs, sheetnames):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for df, name in zip(dfs, sheetnames):
            df.to_excel(writer, index=False, sheet_name=name)
    output.seek(0)
    return output

# App UI
st.title("Vietnam Customer Address Validation Tool")

form_file = st.file_uploader("Upload Microsoft Forms Response File", type=["xlsx"])
ups_file = st.file_uploader("Upload UPS System Address File", type=["xlsx"])

if form_file and ups_file:
    with st.spinner("Processing..."):
        form_df = load_file(form_file, form=True)
        ups_df = load_file(ups_file, form=False)

        matched_df, unmatched_df = match_addresses(form_df, ups_df)
        upload_df = expand_matched_for_upload(matched_df)

        st.success("Validation complete.")

        st.subheader("Matched Addresses")
        st.write(matched_df)

        st.subheader("Unmatched Addresses with Reason")
        st.write(unmatched_df)

        st.subheader("Formatted Upload Template")
        st.write(upload_df)

        excel_data = to_excel_file(
            [matched_df, unmatched_df, upload_df],
            ["Matched", "Unmatched", "Upload_Template"]
        )
        st.download_button("Download All Results", excel_data, file_name="validation_results.xlsx")
