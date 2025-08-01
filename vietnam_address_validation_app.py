import streamlit as st
import pandas as pd
from io import BytesIO
import unicodedata

# Remove Vietnamese tones from a string
def remove_tones(text):
    if not isinstance(text, str):
        return text
    text = unicodedata.normalize('NFD', text)
    text = ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')
    return text

# Normalize string columns: lowercase, strip spaces, remove tones
def normalize_col(col):
    return col.astype(str).str.lower().str.strip().apply(remove_tones)

# Match address line 1 & 2 (number + street), ignoring case and spaces
def address_match(new_line1, new_line2, old_line1, old_line2):
    if not all(isinstance(x, str) for x in [new_line1, new_line2, old_line1, old_line2]):
        return False
    return (remove_tones(new_line1).strip().lower() == remove_tones(old_line1).strip().lower() and
            remove_tones(new_line2).strip().lower() == remove_tones(old_line2).strip().lower())

def process_files(forms_df, ups_df):
    matched_rows = []
    unmatched_rows = []
    upload_template_rows = []

    # Normalize Account Number for matching
    ups_df['Account Number_norm'] = normalize_col(ups_df['Account Number'])
    forms_df['Account Number_norm'] = normalize_col(forms_df['Account Number'])

    # Remove tones and strip spaces for relevant UPS address fields to simplify matching
    for col in ["AC_Name", "Attention_Name", "Address Line 1", "Address Line 2", "City", "Address Line 3", "Postal_Code", "Country_Code", "Address_Country_Code"]:
        if col in ups_df.columns:
            ups_df[col] = ups_df[col].astype(str).apply(remove_tones).str.strip()

    # Remove tones and strip spaces for relevant Forms address fields for output
    forms_addr_cols = [
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
        # Pickup addresses (up to 3)
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
        "Full Name of Contact-In English Only",
        "Please Provide Your Email Address-In English Only",
        "Contact Phone Number"
    ]
    for col in forms_addr_cols:
        if col in forms_df.columns:
            forms_df[col] = forms_df[col].astype(str).apply(remove_tones).str.strip()

    # Group UPS data by account for easy lookup
    ups_grouped = ups_df.groupby('Account Number_norm')

    processed_form_indices = set()

    for idx, form_row in forms_df.iterrows():
        acc_norm = form_row['Account Number_norm']
        is_same_billing = str(form_row.get("Is Your New Billing Address the Same as Your Pickup and Delivery Address?", "")).strip().lower()

        if acc_norm not in ups_grouped.groups:
            # No UPS data for this account, unmatched
            unmatched_rows.append(form_row.to_dict())
            continue

        ups_acc_df = ups_grouped.get_group(acc_norm)

        # Count pickup addresses in UPS system for this account
        ups_pickup_count = (ups_acc_df['Address Type'] == '02').sum()

        if is_same_billing == "yes":
            # Single "All" address type 01 from Forms
            new_addr1 = form_row["New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only"]
            new_addr2 = form_row["New Address Line 2 (Street Name)-In English Only"]
            new_addr3 = form_row["New Address Line 3 (Ward/Commune)-In English Only"]
            city = form_row["City / Province"]
            contact = form_row.get("Full Name of Contact-In English Only", "")
            email = form_row.get("Please Provide Your Email Address-In English Only", "")
            phone = form_row.get("Contact Phone Number", "")

            # Key validation rule 1: Address number + street must match at least one UPS record (any address type)
            key_match_found = False
            ups_row_for_template = None
            for _, ups_row in ups_acc_df.iterrows():
                if address_match(new_addr1, new_addr2, ups_row["Address Line 1"], ups_row["Address Line 2"]):
                    key_match_found = True
                    ups_row_for_template = ups_row
                    break

            if not key_match_found:
                # Unmatched if address number+street do not match any UPS address
                unmatched_rows.append(form_row.to_dict())
                continue

            # Key validation rule 2: number of pickup addresses must be same as UPS
            form_pickup_num = 0
            try:
                form_pickup_num = int(form_row.get("How Many Pick Up Address Do You Have?", 0))
            except:
                form_pickup_num = 0
            if form_pickup_num != ups_pickup_count:
                unmatched_rows.append(form_row.to_dict())
                continue

            # Passed validations -> matched
            matched_rows.append(form_row.to_dict())
            processed_form_indices.add(idx)

            # Upload template requires 3 rows with codes 01, 06, 03 and invoice options  (adjusted per sample)

            # 01 - "All" address type, invoice option blank
            upload_template_rows.append({
                "AC_NUM": form_row["Account Number"],
                "AC_Address_Type": "01",
                "invoice option": "",
                "AC_Name": ups_row_for_template["AC_Name"],
                "Address_Line1": new_addr1,
                "Address_Line2": "",  # blank as in sample
                "City": city,
                "Postal_Code": ups_row_for_template["Postal_Code"],
                "Country_Code": ups_row_for_template["Country_Code"],
                "Attention_Name": contact,
                "Address_Line22": new_addr3,
                "Address_Country_Code": ups_row_for_template["Address_Country_Code"]
            })
            # 06 - secondary code, invoice option blank
            upload_template_rows.append({
                "AC_NUM": form_row["Account Number"],
                "AC_Address_Type": "06",
                "invoice option": "",
                "AC_Name": ups_row_for_template["AC_Name"],
                "Address_Line1": new_addr1,
                "Address_Line2": new_addr2,
                "City": city,
                "Postal_Code": ups_row_for_template["Postal_Code"],
                "Country_Code": ups_row_for_template["Country_Code"],
                "Attention_Name": contact,
                "Address_Line22": new_addr3,
                "Address_Country_Code": ups_row_for_template["Address_Country_Code"]
            })
            # 03 - Billing expanded into invoice options 1,2,6
            for inv_opt in ["1", "2", "6"]:
                upload_template_rows.append({
                    "AC_NUM": form_row["Account Number"],
                    "AC_Address_Type": "03",
                    "invoice option": inv_opt,
                    "AC_Name": ups_row_for_template["AC_Name"],
                    "Address_Line1": new_addr1,
                    "Address_Line2": new_addr2,
                    "City": city,
                    "Postal_Code": ups_row_for_template["Postal_Code"],
                    "Country_Code": ups_row_for_template["Country_Code"],
                    "Attention_Name": contact,
                    "Address_Line22": new_addr3,
                    "Address_Country_Code": ups_row_for_template["Address_Country_Code"]
                })

        else:
            # When "No" for billing same as pickup/delivery,
            # Process billing, delivery, and up to 3 pickup addresses separately.

            # Billing address fields
            billing_addr1 = form_row.get("New Billing Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
            billing_addr2 = form_row.get("New Billing Address Line 2 (Street Name)-In English Only", "")
            billing_addr3 = form_row.get("New Billing Address Line 3 (Ward/Commune)-In English Only", "")
            billing_city = form_row.get("New Billing City / Province", "")

            # Delivery address fields
            delivery_addr1 = form_row.get("New Delivery Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
            delivery_addr2 = form_row.get("New Delivery Address Line 2 (Street Name)-In English Only", "")
            delivery_addr3 = form_row.get("New Delivery Address Line 3 (Ward/Commune)-In English Only", "")
            delivery_city = form_row.get("New Delivery City / Province", "")

            # Number of pickup addresses in form
            pickup_num = 0
            try:
                pickup_num = int(form_row.get("How Many Pick Up Address Do You Have?", 0))
            except:
                pickup_num = 0
            if pickup_num > 3:
                pickup_num = 3  # max 3 pickups

            # Check pickup addresses from the Form: First, Second, Third New Pick Up Address
            pickup_addrs = []
            for i in range(1, pickup_num + 1):
                prefix = ["First", "Second", "Third"][i-1] + " New Pick Up Address"
                pu_addr1 = form_row.get(f"{prefix} Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
                pu_addr2 = form_row.get(f"{prefix} Line 2 (Street Name)-In English Only", "")
                pu_addr3 = form_row.get(f"{prefix} Line 3 (Ward/Commune)-In English Only", "")
                pu_city = form_row.get(f"{prefix} City / Province", "")
                pickup_addrs.append((pu_addr1, pu_addr2, pu_addr3, pu_city))

            # Validate address number and street for each type: must exist in UPS system for the account
            def check_address_in_ups(addr1, addr2, addr_type_code):
                for _, ups_row in ups_acc_df.iterrows():
                    if ups_row["Address Type"] == addr_type_code:
                        if address_match(addr1, addr2, ups_row["Address Line 1"], ups_row["Address Line 2"]):
                            return ups_row
                return None

            # Key validation rule 2: Number of pickup addresses must match UPS count
            if pickup_num != ups_pickup_count:
                unmatched_rows.append(form_row.to_dict())
                continue

            # Flags to track matching for billing/delivery/pickups
            billing_match = check_address_in_ups(billing_addr1, billing_addr2, "03")
            delivery_match = check_address_in_ups(delivery_addr1, delivery_addr2, "13")

            # Check each pickup address matches UPS
            all_pickups_matched = True
            pickup_matches = []
            for pu_addr in pickup_addrs:
                match = check_address_in_ups(pu_addr[0], pu_addr[1], "02")
                if match is None:
                    all_pickups_matched = False
                    break
                pickup_matches.append(match)

            # Key validation rule 1: billing, delivery, pickup addresses number+street must match UPS
            if (billing_match is None) or (delivery_match is None) or not all_pickups_matched:
                unmatched_rows.append(form_row.to_dict())
                continue

            # Passed all validation
            matched_rows.append(form_row.to_dict())
            processed_form_indices.add(idx)

            # Add pickup addresses to upload template (code 02, invoice option blank)
            for pu_addr in pickup_addrs:
                upload_template_rows.append({
                    "AC_NUM": form_row["Account Number"],
                    "AC_Address_Type": "02",
                    "invoice option": "",
                    "AC_Name": ups_acc_df["AC_Name"].values[0],
                    "Address_Line1": pu_addr[0],
                    "Address_Line2": pu_addr[1],
                    "City": pu_addr[3],
                    "Postal_Code": ups_acc_df["Postal_Code"].values[0],
                    "Country_Code": ups_acc_df["Country_Code"].values[0],
                    "Attention_Name": form_row.get("Full Name of Contact-In English Only", ""),
                    "Address_Line22": pu_addr[2],
                    "Address_Country_Code": ups_acc_df["Address_Country_Code"].values[0]
                })

            # Add billing address (expand to 01, 06, 03 with invoice option 1, 2, 6 as in sample)
            upload_template_rows.append({
                "AC_NUM": form_row["Account Number"],
                "AC_Address_Type": "01",
                "invoice option": "",
                "AC_Name": ups_acc_df["AC_Name"].values[0],
                "Address_Line1": billing_addr1,
                "Address_Line2": "",  # blank as in sample
                "City": billing_city,
                "Postal_Code": ups_acc_df["Postal_Code"].values[0],
                "Country_Code": ups_acc_df["Country_Code"].values[0],
                "Attention_Name": form_row.get("Full Name of Contact-In English Only", ""),
                "Address_Line22": billing_addr3,
                "Address_Country_Code": ups_acc_df["Address_Country_Code"].values[0]
            })
            upload_template_rows.append({
                "AC_NUM": form_row["Account Number"],
                "AC_Address_Type": "06",
                "invoice option": "",
                "AC_Name": ups_acc_df["AC_Name"].values[0],
                "Address_Line1": billing_addr1,
                "Address_Line2": billing_addr2,
                "City": billing_city,
                "Postal_Code": ups_acc_df["Postal_Code"].values[0],
                "Country_Code": ups_acc_df["Country_Code"].values[0],
                "Attention_Name": form_row.get("Full Name of Contact-In English Only", ""),
                "Address_Line22": billing_addr3,
                "Address_Country_Code": ups_acc_df["Address_Country_Code"].values[0]
            })
            for inv_opt in ["1", "2", "6"]:
                upload_template_rows.append({
                    "AC_NUM": form_row["Account Number"],
                    "AC_Address_Type": "03",
                    "invoice option": inv_opt,
                    "AC_Name": ups_acc_df["AC_Name"].values[0],
                    "Address_Line1": billing_addr1,
                    "Address_Line2": billing_addr2,
                    "City": billing_city,
                    "Postal_Code": ups_acc_df["Postal_Code"].values[0],
                    "Country_Code": ups_acc_df["Country_Code"].values[0],
                    "Attention_Name": form_row.get("Full Name of Contact-In English Only", ""),
                    "Address_Line22": billing_addr3,
                    "Address_Country_Code": ups_acc_df["Address_Country_Code"].values[0]
                })

            # Add delivery address (AC_Address_Type 13, invoice option blank)
            upload_template_rows.append({
                "AC_NUM": form_row["Account Number"],
                "AC_Address_Type": "13",
                "invoice option": "",
                "AC_Name": ups_acc_df["AC_Name"].values[0],
                "Address_Line1": delivery_addr1,
                "Address_Line2": delivery_addr2,
                "City": delivery_city,
                "Postal_Code": ups_acc_df["Postal_Code"].values[0],
                "Country_Code": ups_acc_df["Country_Code"].values[0],
                "Attention_Name": form_row.get("Full Name of Contact-In English Only", ""),
                "Address_Line22": delivery_addr3,
                "Address_Country_Code": ups_acc_df["Address_Country_Code"].values[0]
            })

    # Add Forms rows never matched to unmatched
    unmatched_rows.extend(forms_df.loc[~forms_df.index.isin(processed_form_indices)].to_dict('records'))

    # Convert lists of dicts back to DataFrames
    matched_df = pd.DataFrame(matched_rows)
    unmatched_df = pd.DataFrame(unmatched_rows)
    upload_template_df = pd.DataFrame(upload_template_rows)

    # Fix column order of upload template exactly as per sample
    correct_cols = [
        "AC_NUM", "AC_Address_Type", "invoice option", "AC_Name", "Address_Line1",
        "Address_Line2", "City", "Postal_Code", "Country_Code", "Attention_Name",
        "Address_Line22", "Address_Country_Code"
    ]

    # Sometimes upload_template_df may miss columns if no rows present; fix by reindexing
    upload_template_df = upload_template_df.reindex(columns=correct_cols, fill_value="")

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

            matched_df, unmatched_df, upload_template_df = process_files(forms_df, ups_df)

            st.success(f"✅ Completed: {len(matched_df)} matched, {len(unmatched_df
