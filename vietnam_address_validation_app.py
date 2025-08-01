
import streamlit as st
import pandas as pd
from io import BytesIO
import unicodedata

def remove_tones(text):
    if not isinstance(text, str):
        return text
    text = unicodedata.normalize('NFD', text)
    return ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')

def normalize_col(col):
    return col.astype(str).str.lower().str.strip().apply(remove_tones)

def address_match(new_line1, new_line2, old_line1, old_line2):
    if not all(isinstance(x, str) for x in [new_line1, new_line2, old_line1, old_line2]):
        return False
    return (remove_tones(new_line1).strip().lower() == remove_tones(old_line1).strip().lower() and
            remove_tones(new_line2).strip().lower() == remove_tones(old_line2).strip().lower())

def process_files(forms_df, ups_df):
    matched_rows, unmatched_rows, upload_template_rows = [], [], []
    ups_df['Account Number_norm'] = normalize_col(ups_df['Account Number'])
    forms_df['Account Number_norm'] = normalize_col(forms_df['Account Number'])

    for col in ["AC_Name", "Attention_Name", "Address Line 1", "Address Line 2", "City", "Address Line 3"]:
        if col in ups_df.columns:
            ups_df[col] = ups_df[col].astype(str).apply(remove_tones).str.strip()

    for col in forms_df.columns:
        if isinstance(forms_df[col].iloc[0], str):
            forms_df[col] = forms_df[col].astype(str).apply(remove_tones).str.strip()

    ups_grouped = ups_df.groupby('Account Number_norm')
    processed_form_indices = set()

    for idx, form_row in forms_df.iterrows():
        acc_norm = form_row['Account Number_norm']
        if acc_norm not in ups_grouped.groups:
            unmatched_rows.append(form_row.to_dict())
            continue
        matched_rows.append(form_row.to_dict())
        processed_form_indices.add(idx)
        upload_template_rows.append({
            "AC_NUM": form_row["Account Number"],
            "AC_Address_Type": "01",
            "invoice option": "",
            "AC_Name": "SAMPLE",
            "Address_Line1": "SAMPLE",
            "Address_Line2": "",
            "City": "SAMPLE",
            "Postal_Code": "700000",
            "Country_Code": "VN",
            "Attention_Name": "SAMPLE",
            "Address_Line22": "SAMPLE",
            "Address_Country_Code": "VN"
        })

    unmatched_rows.extend(forms_df.loc[~forms_df.index.isin(processed_form_indices)].to_dict('records'))
    matched_df = pd.DataFrame(matched_rows)
    unmatched_df = pd.DataFrame(unmatched_rows)
    upload_template_df = pd.DataFrame(upload_template_rows)
    cols_order = [
        "AC_NUM", "AC_Address_Type", "invoice option", "AC_Name", "Address_Line1",
        "Address_Line2", "City", "Postal_Code", "Country_Code", "Attention_Name",
        "Address_Line22", "Address_Country_Code"
    ]
    upload_template_df = upload_template_df[cols_order]
    return matched_df, unmatched_df, upload_template_df

def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def main():
    st.set_page_config(page_title="Vietnam Address Validation Tool", layout="wide")
    st.title("🇻🇳 Vietnam Address Validation Tool")
    st.write("Upload Microsoft Forms response file and UPS system address file to validate and generate upload template.")

    forms_file = st.file_uploader("Upload Microsoft Forms Response File (.xlsx)", type=["xlsx"])
    ups_file = st.file_uploader("Upload UPS System Address File (.xlsx)", type=["xlsx"])

    if forms_file and ups_file:
        with st.spinner("Processing files..."):
            forms_df = pd.read_excel(forms_file)
            ups_df = pd.read_excel(ups_file)
            matched_df, unmatched_df, upload_template_df = process_files(forms_df, ups_df)
            st.success(f"✅ Completed: {len(matched_df)} matched, {len(unmatched_df)} unmatched.")

            if not matched_df.empty:
                st.download_button("Download Matched Records", to_excel_bytes(matched_df), "matched_records.xlsx")
            if not unmatched_df.empty:
                st.download_button("Download Unmatched Records", to_excel_bytes(unmatched_df), "unmatched_records.xlsx")
            if not upload_template_df.empty:
                st.download_button("Download Upload Template", to_excel_bytes(upload_template_df), "upload_template.xlsx")

if __name__ == "__main__":
    main()
