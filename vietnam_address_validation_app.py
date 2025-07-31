import streamlit as st
import pandas as pd
import os

def normalize_address(s):
    return str(s).strip().lower()

def match_address(old, new):
    return normalize_address(old) == normalize_address(new)

def is_valid_change(old_line3, new_line3):
    return normalize_address(old_line3) != normalize_address(new_line3)

def expand_billing_to_three_rows(row, upload_columns):
    billing_rows = []
    for code in ['1', '2', '6']:
        new_row = row.copy()
        new_row['Address Code'] = code
        billing_rows.append(new_row[upload_columns])
    return billing_rows

def convert_to_upload_template(df):
    upload_data = []
    upload_columns = [
        'Customer Code', 'Company Name', 'Address Code',
        'New Address Line 1', 'New Address Line 2', 'New Address Line 3',
        'City / Province', 'Contact Name', 'Phone Number', 'Email'
    ]

    for _, row in df.iterrows():
        address_type = row['Address Type']
        if address_type == '01':  # All
            upload_data.append(row[upload_columns])
        elif address_type == '03':  # Billing → split into 3 rows
            upload_data.extend(expand_billing_to_three_rows(row, upload_columns))
        else:
            upload_data.append(row[upload_columns])

    return pd.DataFrame(upload_data)

def main():
    st.title("Vietnam Customer Address Validation Tool")

    forms_file = st.file_uploader("Upload Microsoft Forms Response File", type=["xlsx"])
    ups_file = st.file_uploader("Upload UPS System File", type=["xlsx"])

    if forms_file and ups_file:
        forms_df = pd.read_excel(forms_file)
        ups_df = pd.read_excel(ups_file)

        matched_rows = []
        unmatched_rows = []

        for _, form_row in forms_df.iterrows():
            account = str(form_row.get("Account Number")).strip()
            is_all = str(form_row.get("Is Your New Billing Address the Same as Your Pickup and Delivery Address?")).strip().lower() == "yes"

            if is_all:
                match_rows = ups_df[(ups_df["Account Number"].astype(str).str.strip() == account)]
                if not match_rows.empty:
                    for _, sys_row in match_rows.iterrows():
                        addr_type = sys_row["Address Type"]
                        if addr_type in ["01", "02", "03", "13"]:
                            match = {
                                'Customer Code': account,
                                'Company Name': form_row.get("Company Name", ""),
                                'Address Type': "01",
                                'Address Code': '01',
                                'New Address Line 1': form_row.get("New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", ""),
                                'New Address Line 2': form_row.get("New Address Line 2 (Street Name)-In English Only", ""),
                                'New Address Line 3': form_row.get("New Address Line 3 (Ward/Commune)-In English Only", ""),
                                'City / Province': form_row.get("City / Province", ""),
                                'Contact Name': form_row.get("Full Name of Contact-In English Only", ""),
                                'Phone Number': form_row.get("Contact Phone Number", ""),
                                'Email': form_row.get("Please Provide Your Email Address-In English Only", "")
                            }
                            matched_rows.append(match)
                        break
                else:
                    unmatched_rows.append(form_row)
            else:
                found_match = False
                for prefix in ["First", "Second", "Third"]:
                    pickup1 = form_row.get(f"{prefix} New Pick Up Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
                    pickup2 = form_row.get(f"{prefix} New Pick Up Address Line 2 (Street Name)-In English Only", "")
                    pickup3 = form_row.get(f"{prefix} New Pick Up Address Line 3 (Ward/Commune)-In English Only", "")
                    if pickup1 and pickup2:
                        match_rows = ups_df[(ups_df["Account Number"].astype(str).str.strip() == account) & (ups_df["Address Type"] == "02")]
                        for _, sys_row in match_rows.iterrows():
                            if match_address(sys_row["Address Line 1"], pickup1) and match_address(sys_row["Address Line 2"], pickup2):
                                match = {
                                    'Customer Code': account,
                                    'Company Name': form_row.get("Company Name", ""),
                                    'Address Type': "02",
                                    'Address Code': '02',
                                    'New Address Line 1': pickup1,
                                    'New Address Line 2': pickup2,
                                    'New Address Line 3': pickup3,
                                    'City / Province': form_row.get("City / Province", ""),
                                    'Contact Name': form_row.get("Full Name of Contact-In English Only", ""),
                                    'Phone Number': form_row.get("Contact Phone Number", ""),
                                    'Email': form_row.get("Please Provide Your Email Address-In English Only", "")
                                }
                                matched_rows.append(match)
                                found_match = True
                                break
                if not found_match:
                    unmatched_rows.append(form_row)

        matched_df = pd.DataFrame(matched_rows)
        unmatched_df = pd.DataFrame(unmatched_rows)
        upload_df = convert_to_upload_template(matched_df)

        st.success("Validation complete.")
        st.download_button("Download Matched File", matched_df.to_csv(index=False), "matched.csv")
        st.download_button("Download Unmatched File", unmatched_df.to_csv(index=False), "unmatched.csv")
        st.download_button("Download Upload Template", upload_df.to_csv(index=False), "upload_template.csv")

if __name__ == "__main__":
    main()