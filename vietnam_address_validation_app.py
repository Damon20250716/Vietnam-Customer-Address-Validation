import streamlit as st
import pandas as pd
import unicodedata
from io import BytesIO

# --- Helper Functions ---
def remove_tones(text):
    if not isinstance(text, str):
        return text
    text = unicodedata.normalize('NFD', text)
    return ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')

def normalize_address(val):
    if not isinstance(val, str):
        return ''
    val = val.strip().lower()
    val = remove_tones(val)
    return val

def fuzzy_column_match(df_columns, target):
    """Attempt to find the closest match to a target column name."""
    for col in df_columns:
        if normalize_address(col) == normalize_address(target):
            return col
    return None

def load_file(uploaded_file, form=True):
    df = pd.read_excel(uploaded_file)
    if form:
        required_fields = [
            'Account Number',
            'Is Your New Billing Address the Same as Your Pickup and Delivery Address?',
            'New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only',
            'New Address Line 2 (Street Name)-In English Only'
        ]
        col_map = {}
        for field in required_fields:
            match = fuzzy_column_match(df.columns, field)
            if not match:
                raise ValueError(f"Missing column in Forms file: {field}")
            col_map[field] = match
        df = df.rename(columns=col_map)
    else:
        if 'AC_NUM' not in df.columns:
            raise ValueError("Missing 'AC_NUM' column in UPS file.")
        df['AC_NUM'] = df['AC_NUM'].astype(str).str.strip()
    return df

def match_address(row, ups_df):
    account = str(row['Account Number']).strip()
    same_address = row['Is Your New Billing Address the Same as Your Pickup and Delivery Address?'].strip().lower() == 'yes'
    new_line1 = normalize_address(row['New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only'])
    new_line2 = normalize_address(row['New Address Line 2 (Street Name)-In English Only'])

    matched_rows = []
    if same_address:
        filtered = ups_df[ups_df['AC_NUM'].astype(str).str.strip() == account]
        for _, sys_row in filtered.iterrows():
            sys_line1 = normalize_address(sys_row['Address_Line1'])
            sys_line2 = normalize_address(sys_row['Address_Line2'])
            if new_line1 in sys_line1 and new_line2 in sys_line2:
                matched_rows.append((sys_row, '01'))
                break
    else:
        # Logic for billing/pickup/delivery (02/03/13) if needed
        pass

    return matched_rows

def generate_upload_template(matched):
    rows = []
    for match in matched:
        sys_row, code = match
        if code == '01':
            for c in ['1', '2', '6']:
                rows.append({
                    'AC_NUM': sys_row['AC_NUM'],
                    'AC_Address_Type': c,
                    'AC_Name': sys_row['AC_Name'],
                    'Address_Line1': sys_row['Address_Line1'],
                    'Address_Line2': sys_row['Address_Line2'],
                    'City': sys_row['City'],
                    'Postal_Code': sys_row['Postal_Code'],
                    'Country_Code': sys_row['Country_Code'],
                    'Attention_Name': sys_row['Attention_Name'],
                    'Address_Line22': sys_row['Address_Line2'],
                    'Address_Country_Code': sys_row['Country_Code'],
                })
    return pd.DataFrame(rows)

def convert_df(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# --- Streamlit UI ---
st.title("Vietnam Address Validation Tool v2")

form_file = st.file_uploader("Upload Microsoft Forms Response File (.xlsx)", type=["xlsx"])
ups_file = st.file_uploader("Upload UPS System File (.xlsx)", type=["xlsx"])

if form_file and ups_file:
    try:
        forms_df = load_file(form_file, form=True)
        ups_df = load_file(ups_file, form=False)

        matched_results = []
        unmatched_rows = []

        for _, row in forms_df.iterrows():
            try:
                matched = match_address(row, ups_df)
                if matched:
                    matched_results.extend(matched)
                else:
                    row['Unmatched Reason'] = 'No matching AC_NUM or Address Line mismatch'
                    unmatched_rows.append(row)
            except Exception as e:
                row['Unmatched Reason'] = str(e)
                unmatched_rows.append(row)

        matched_df = pd.DataFrame([dict(sys_row) | {'Address_Type_Code': code} for sys_row, code in matched_results])
        unmatched_df = pd.DataFrame(unmatched_rows)

        st.success(f"Matched: {len(matched_df)}, Unmatched: {len(unmatched_df)}")

        # Show download buttons
        st.download_button("Download Matched File", convert_df(matched_df), file_name="Matched_Results.xlsx")
        st.download_button("Download Unmatched File", convert_df(unmatched_df), file_name="Unmatched_Responses.xlsx")

        # Uploading template
        template_df = generate_upload_template(matched_results)
        st.download_button("Download Upload Template", convert_df(template_df), file_name="Upload_Template.xlsx")

    except Exception as e:
        st.error(f"❌ An error occurred: {str(e)}")
