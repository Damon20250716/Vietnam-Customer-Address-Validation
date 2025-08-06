import streamlit as st
import pandas as pd
from io import BytesIO
import unicodedata
import difflib
import re

# Remove Vietnamese tones from text
def remove_tones(text):
    if not isinstance(text, str):
        return ""
    text = unicodedata.normalize('NFD', text)
    text = ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')
    return text

def clean_addr_str(s):
    if not isinstance(s, str):
        return ""
    s = s.lower()
    s = unicodedata.normalize('NFD', s)
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')  # remove tones
    s = re.sub(r'[.,]', '', s)  # remove dots and commas
    s = ' '.join(s.split())  # normalize whitespace
    return s

def fuzzy_match(s1, s2, threshold=0.7):
    seq = difflib.SequenceMatcher(None, s1, s2)
    ratio = seq.ratio()
    return ratio >= threshold

def is_address_match(forms_line1, forms_line2, ups_line1, ups_line2):
    f_line1 = clean_addr_str(forms_line1)
    f_line2 = clean_addr_str(forms_line2)
    u_line1 = clean_addr_str(ups_line1)
    u_line2 = clean_addr_str(ups_line2)

    cond1 = (f_line1 in u_line1) or (u_line1 in f_line1) or fuzzy_match(f_line1, u_line1)
    cond2 = (f_line2 in u_line2) or (u_line2 in f_line2) or fuzzy_match(f_line2, u_line2)

    return cond1 and cond2

def process_files(forms_df, ups_df):
    matched_rows = []
    unmatched_rows = []
    upload_template_rows = []

    # Normalize Account Number for matching (lowercase & strip only)
    ups_df['Account Number_norm'] = ups_df['Account Number'].astype(str).str.lower().str.strip()
    forms_df['Account Number_norm'] = forms_df['Account Number'].astype(str).str.lower().str.strip()

    ups_grouped = ups_df.groupby('Account Number_norm')

    processed_form_indices = set()

    # For quick UPS AC_Name, Postal_Code etc retrieval per account number
    ups_ac_info = ups_df.groupby('Account Number_norm').first()

    for idx, form_row in forms_df.iterrows():
        acc_norm = form_row['Account Number_norm']
        is_same_billing = str(form_row.get("Is Your New Billing Address the Same as Your Pickup and Delivery Address?", "")).strip().lower()

        if acc_norm not in ups_grouped.groups:
            unmatched_dict = form_row.to_dict()
            unmatched_dict['Unmatched Reason'] = "Account Number not found in UPS data"
            unmatched_rows.append(unmatched_dict)
            continue

        ups_acc_df = ups_grouped.get_group(acc_norm)

        ups_pickup_count = (ups_acc_df['Address Type'] == '02').sum()

        # === Case 1: Billing same as Pickup & Delivery ===
        if is_same_billing == "yes":
            new_addr1 = form_row.get("New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
            new_addr2 = form_row.get("New Address Line 2 (Street Name)-In English Only", "")
            new_addr3 = form_row.get("New Address Line 3 (Ward/Commune)-In English Only", "")
            city = form_row.get("City / Province", "")
            contact = form_row.get("Full Name of Contact-In English Only", "")

            matched_in_ups = False
            ups_row_for_template = None
            for _, ups_row in ups_acc_df.iterrows():
                if ups_row["Address Type"] == '01':  # address type all
                    if is_address_match(new_addr1, new_addr2, ups_row["Address Line 1"], ups_row["Address Line 2"]):
                        matched_in_ups = True
                        ups_row_for_template = ups_row
                        break

            if matched_in_ups:
                matched_dict = form_row.to_dict()
                # Remove tones for output
                matched_dict["New Address Line 1 (Tone-free)"] = remove_tones(new_addr1)
                matched_dict["New Address Line 2 (Tone-free)"] = remove_tones(new_addr2)
                matched_dict["New Address Line 3 (Tone-free)"] = remove_tones(new_addr3)
                matched_rows.append(matched_dict)
                processed_form_indices.add(idx)

                # Upload template rows - 3 rows with code 1, 2, 6 for billing address
                for code in ["1", "2", "6"]:
                    upload_template_rows.append({
                        "AC_NUM": form_row["Account Number"],
                        "AC_Address_Type": code,
                        "invoice option": code,
                        "AC_Name": ups_row_for_template["AC_Name"],
                        "Address_Line1": remove_tones(new_addr1),
                        "Address_Line2": remove_tones(new_addr2),
                        "City": city,
                        "Postal_Code": ups_row_for_template["Postal_Code"],
                        "Country_Code": ups_row_for_template["Country_Code"],
                        "Attention_Name": contact,
                        "Address_Line22": remove_tones(new_addr3),
                        "Address_Country_Code": ups_row_for_template["Address_Country_Code"]
                    })
            else:
                unmatched_dict = form_row.to_dict()
                unmatched_dict['Unmatched Reason'] = "Billing address (type 01) not matched in UPS system"
                unmatched_rows.append(unmatched_dict)

        # === Case 2: Billing, Delivery, Pickup separately ===
        else:
            billing_addr1 = form_row.get("New Billing Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
            billing_addr2 = form_row.get("New Billing Address Line 2 (Street Name)-In English Only", "")
            billing_addr3 = form_row.get("New Billing Address Line 3 (Ward/Commune)-In English Only", "")
            billing_city = form_row.get("New Billing City / Province", "")

            delivery_addr1 = form_row.get("New Delivery Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
            delivery_addr2 = form_row.get("New Delivery Address Line 2 (Street Name)-In English Only", "")
            delivery_addr3 = form_row.get("New Delivery Address Line 3 (Ward/Commune)-In English Only", "")
            delivery_city = form_row.get("New Delivery City / Province", "")

            try:
                pickup_num = int(form_row.get("How Many Pick Up Address Do You Have?", 0))
            except:
                pickup_num = 0
            if pickup_num > 3:
                pickup_num = 3

            pickup_addrs = []
            for i in range(1, pickup_num + 1):
                prefix = ["First", "Second", "Third"][i-1] + " New Pick Up Address"
                pu_addr1 = form_row.get(f"{prefix} Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
                pu_addr2 = form_row.get(f"{prefix} Line 2 (Street Name)-In English Only", "")
                pu_addr3 = form_row.get(f"{prefix} Line 3 (Ward/Commune)-In English Only", "")
                pu_city = form_row.get(f"{prefix} City / Province", "")
                pickup_addrs.append((pu_addr1, pu_addr2, pu_addr3, pu_city))

            def check_address_in_ups(addr1, addr2, addr_type_code):
                for _, ups_row in ups_acc_df.iterrows():
                    if ups_row["Address Type"] == addr_type_code:
                        if is_address_match(addr1, addr2, ups_row["Address Line 1"], ups_row["Address Line 2"]):
                            return ups_row
                return None

            billing_match = check_address_in_ups(billing_addr1, billing_addr2, "03")
            delivery_match = check_address_in_ups(delivery_addr1, delivery_addr2, "13")

            if len(pickup_addrs) != ups_pickup_count:
                unmatched_dict = form_row.to_dict()
                unmatched_dict['Unmatched Reason'] = f"Pickup address count mismatch: Forms={len(pickup_addrs)}, UPS={ups_pickup_count}"
                unmatched_rows.append(unmatched_dict)
                continue
            else:
                pickup_matches = []
                for pu_addr in pickup_addrs:
                    match = check_address_in_ups(pu_addr[0], pu_addr[1], "02")
                    if match is None:
                        unmatched_dict = form_row.to_dict()
                        unmatched_dict['Unmatched Reason'] = f"Pickup address not matched: {pu_addr[0]}, {pu_addr[1]}"
                        unmatched_rows.append(unmatched_dict)
                        break
                    else:
                        pickup_matches.append(match)
                else:
                    # All pickup addresses matched
                    processed_form_indices.add(idx)
                    matched_dict = form_row.to_dict()
                    # Add tone-free address columns for matched file output
                    matched_dict["New Billing Address Line 1 (Tone-free)"] = remove_tones(billing_addr1)
                    matched_dict["New Billing Address Line 2 (Tone-free)"] = remove_tones(billing_addr2)
                    matched_dict["New Billing Address Line 3 (Tone-free)"] = remove_tones(billing_addr3)
                    matched_dict["New Delivery Address Line 1 (Tone-free)"] = remove_tones(delivery_addr1)
                    matched_dict["New Delivery Address Line 2 (Tone-free)"] = remove_tones(delivery_addr2)
                    matched_dict["New Delivery Address Line 3 (Tone-free)"] = remove_tones(delivery_addr3)
                    for i, pu_addr in enumerate(pickup_addrs, 1):
                        matched_dict[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 1 (Tone-free)"] = remove_tones(pu_addr[0])
                        matched_dict[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 2 (Tone-free)"] = remove_tones(pu_addr[1])
                        matched_dict[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 3 (Tone-free)"] = remove_tones(pu_addr[2])

                    matched_rows.append(matched_dict)

                    # Add pickup addresses to upload template, each one a row
                    for pu_addr in pickup_addrs:
                        upload_template_rows.append({
                            "AC_NUM": form_row["Account Number"],
                            "AC_Address_Type": "02",
                            "invoice option": "",  # pickup address - empty invoice option
                            "AC_Name": ups_ac_info.loc[acc_norm, "AC_Name"],
                            "Address_Line1": remove_tones(pu_addr[0]),
                            "Address_Line2": remove_tones(pu_addr[1]),
                            "City": pu_addr[3],
                            "Postal_Code": ups_ac_info.loc[acc_norm, "Postal_Code"],
                            "Country_Code": ups_ac_info.loc[acc_norm, "Country_Code"],
                            "Attention_Name": form_row.get("Full Name of Contact-In English Only", ""),
                            "Address_Line22": remove_tones(pu_addr[2]),
                            "Address_Country_Code": ups_ac_info.loc[acc_norm, "Address_Country_Code"]
                        })

                    # Add billing addresses - 3 rows with codes 1, 2, 6
                    for code in ["1", "2", "6"]:
                        upload_template_rows.append({
                            "AC_NUM": form_row["Account Number"],
                            "AC_Address_Type": code,
                            "invoice option": code,
                            "AC_Name": ups_ac_info.loc[acc_norm, "AC_Name"],
                            "Address_Line1": remove_tones(billing_addr1),
                            "Address_Line2": remove_tones(billing_addr2),
                            "City": billing_city,
                            "Postal_Code": ups_ac_info.loc[acc_norm, "Postal_Code"],
                            "Country_Code": ups_ac_info.loc[acc_norm, "Country_Code"],
                            "Attention_Name": form_row.get("Full Name of Contact-In English Only", ""),
                            "Address_Line22": remove_tones(billing_addr3),
                            "Address_Country_Code": ups_ac_info.loc[acc_norm, "Address_Country_Code"]
                        })

                    # Add delivery address - one row
                    upload_template_rows.append({
                        "AC_NUM": form_row["Account Number"],
                        "AC_Address_Type": "13",
                        "invoice option": "",
                        "AC_Name": ups_ac_info.loc[acc_norm, "AC_Name"],
                        "Address_Line1": remove_tones(delivery_addr1),
                        "Address_Line2": remove_tones(delivery_addr2),
                        "City": delivery_city,
                        "Postal_Code": ups_ac_info.loc[acc_norm, "Postal_Code"],
                        "Country_Code": ups_ac_info.loc[acc_norm, "Country_Code"],
                        "Attention_Name": form_row.get("Full Name of Contact-In English Only", ""),
                        "Address_Line22": remove_tones(delivery_addr3),
                        "Address_Country_Code": ups_ac_info.loc[acc_norm, "Address_Country_Code"]
                    })

    # Handle any Forms rows that were not processed (no match)
    unmatched_not_processed = forms_df.loc[~forms_df.index.isin(processed_form_indices)]
    for _, row in unmatched_not_processed.iterrows():
        unmatched_dict = row.to_dict()
        unmatched_dict['Unmatched Reason'] = "No matching address found or not processed"
        unmatched_rows.append(unmatched_dict)

    matched_df = pd.DataFrame(matched_rows)
    unmatched_df = pd.DataFrame(unmatched_rows)
    upload_template_df = pd.DataFrame(upload_template_rows)

    return matched_df, unmatched_df, upload_template_df


# --- Streamlit UI ---
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

            # Check required columns (basic check)
            required_forms_cols = [
                "Account Number",
                "Is Your New Billing Address the Same as Your Pickup and Delivery Address?",
                "New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only",
                "New Address Line 2 (Street Name)-In English Only",
                "New Address Line 3 (Ward/Commune)-In English Only",
                "City / Province",
                "New Billing Address Line 1 (Address No., Industrial Park Name, etc)-In English Only",
                "New Billing Address Line 2 (Street Name)-In English Only",
                "New Billing Address Line 3 (Ward/Commune)-In English Only",
                "New Billing City / Province",
                "New Delivery Address Line 1 (Address No., Industrial Park Name, etc)-In English Only",
                "New Delivery Address Line 2 (Street Name)-In English Only",
                "New Delivery Address Line 3 (Ward/Commune)-In English Only",
                "New Delivery City / Province",
                "How Many Pick Up Address Do You Have?",
                "First New Pick Up Address Line 1 (Address No., Industrial Park Name, etc)-In English Only",
                "First New Pick Up Address Line 2 (Street Name)-In English Only",
                "First New Pick Up Address Line 3 (Ward/Commune)-In English Only",
                "First New Pick Up City / Province",
                "Second New Pick Up Address Line 1 (Address No., Industrial Park Name, etc)-In English Only",
                "Second New Pick Up Address Line 2 (Street Name)-In English Only",
                "Second New Pick Up Address Line 3 (Ward/Commune)-In English Only",
                "Second New Pick Up City / Province",
                "Third New Pick Up Address Line 1 (Address No., Industrial Park Name, etc)-In English Only",
                "Third New Pick Up Address Line 2 (Street Name)-In English Only",
                "Third New Pick Up Address Line 3 (Ward/Commune)-In English Only",
                "Third New Pick Up City / Province",
                "Full Name of Contact-In English Only"
            ]
            missing_cols = [c for c in required_forms_cols if c not in forms_df.columns]
            if missing_cols:
                st.error(f"Missing required columns in Forms file: {missing_cols}")
                return

            required_ups_cols = [
                "Account Number",
                "Address Type",
                "Address Line 1",
                "Address Line 2",
                "Address Line 3",
                "City",
                "Postal_Code",
                "Country_Code",
                "AC_Name",
                "Address_Country_Code"
            ]
            missing_ups_cols = [c for c in required_ups_cols if c not in ups_df.columns]
            if missing_ups_cols:
                st.error(f"Missing required columns in UPS file: {missing_ups_cols}")
                return

            matched_df, unmatched_df, upload_template_df = process_files(forms_df, ups_df)

            st.success(f"✅ Validation completed! Matched: {len(matched_df)}, Unmatched: {len(unmatched_df)}")

            def to_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False)
                return output.getvalue()

            if not matched_df.empty:
                st.download_button(
                    label="Download Matched Records",
                    data=to_excel(matched_df),
                    file_name="matched_records.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            if not unmatched_df.empty:
                st.download_button(
                    label="Download Unmatched Records",
                    data=to_excel(unmatched_df),
                    file_name="unmatched_records.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            if not upload_template_df.empty:
                st.download_button(
                    label="Download Upload Template",
                    data=to_excel(upload_template_df),
                    file_name="upload_template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()
