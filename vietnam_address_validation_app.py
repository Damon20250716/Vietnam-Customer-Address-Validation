import streamlit as st
import pandas as pd
import unicodedata
from io import BytesIO
import re

# --- Helper functions ---

def remove_tones(text):
    if not isinstance(text, str):
        return ""
    text = unicodedata.normalize('NFD', text)
    return ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')

def normalize_address(val):
    if not isinstance(val, str):
        return ""
    val = val.strip().lower()
    val = remove_tones(val)
    return val

def extract_number_and_street(addr_line1):
    addr_line1 = normalize_address(addr_line1)
    # Extract digits + street name before comma or ward keyword (very basic heuristic)
    match = re.match(r"(\d+\w*\s[\w\s]+?)(,| ward|$)", addr_line1)
    if match:
        return match.group(1).strip()
    # fallback: return normalized addr_line1 fully
    return addr_line1

def flexible_address_match(forms_line1, forms_line2, ups_line1, ups_line2):
    forms_ns = extract_number_and_street(forms_line1) + " " + normalize_address(forms_line2)
    ups_ns = extract_number_and_street(ups_line1) + " " + normalize_address(ups_line2)
    return forms_ns == ups_ns

# --- Main processing function ---

def process_files(forms_df, ups_df):
    matched_rows = []
    unmatched_rows = []
    upload_template_rows = []

    # Normalize account number for matching
    ups_df['Account Number_norm'] = ups_df['Account Number'].astype(str).str.lower().str.strip()
    forms_df['Account Number_norm'] = forms_df['Account Number'].astype(str).str.lower().str.strip()

    ups_grouped = ups_df.groupby('Account Number_norm')

    processed_form_indices = set()

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

        if is_same_billing == "yes":
            new_addr1 = form_row.get("New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
            new_addr2 = form_row.get("New Address Line 2 (Street Name)-In English Only", "")
            new_addr3 = form_row.get("New Address Line 3 (Ward/Commune)-In English Only", "")
            city = form_row.get("City / Province", "")
            contact = form_row.get("Full Name of Contact-In English Only", "")

            matched_in_ups = False
            ups_row_for_template = None

            for _, ups_row in ups_acc_df.iterrows():
                if ups_row['Address Type'] == '01':  # All address type
                    if flexible_address_match(new_addr1, new_addr2, ups_row["Address Line 1"], ups_row["Address Line 2"]):
                        matched_in_ups = True
                        ups_row_for_template = ups_row
                        break

            if matched_in_ups:
                matched_dict = form_row.to_dict()
                # Remove tones in matched file
                matched_dict["New Address Line 1 (Tone-free)"] = normalize_address(new_addr1)
                matched_dict["New Address Line 2 (Tone-free)"] = normalize_address(new_addr2)
                matched_dict["New Address Line 3 (Tone-free)"] = normalize_address(new_addr3)
                matched_rows.append(matched_dict)
                processed_form_indices.add(idx)

                # Upload template: 3 rows with codes 1, 2, 6
                for code in ["1", "2", "6"]:
                    upload_template_rows.append({
                        "AC_NUM": form_row["Account Number"],
                        "AC_Address_Type": code,
                        "invoice option": code,
                        "AC_Name": ups_row_for_template["AC_Name"],
                        "Address_Line1": normalize_address(new_addr1),
                        "Address_Line2": normalize_address(new_addr2),
                        "City": city,
                        "Postal_Code": ups_row_for_template["Postal_Code"],
                        "Country_Code": ups_row_for_template["Country_Code"],
                        "Attention_Name": contact,
                        "Address_Line22": normalize_address(new_addr3),
                        "Address_Country_Code": ups_row_for_template["Address_Country_Code"]
                    })
            else:
                unmatched_dict = form_row.to_dict()
                unmatched_dict['Unmatched Reason'] = "Billing address (type 01) not matched in UPS system"
                unmatched_rows.append(unmatched_dict)

        else:
            # Extract billing, delivery, pickup addresses
            billing_addr1 = form_row.get("New Billing Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
            billing_addr2 = form_row.get("New Billing Address Line 2 (Street Name)-In English Only", "")
            billing_addr3 = form_row.get("New Billing Address Line 3 (Ward/Commune)-In English Only", "")
            billing_city = form_row.get("New Billing City / Province", "")

            delivery_addr1 = form_row.get("New Delivery Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
            delivery_addr2 = form_row.get("New Delivery Address Line 2 (Street Name)-In English Only", "")
            delivery_addr3 = form_row.get("New Delivery Address Line 3 (Ward/Commune)-In English Only", "")
            delivery_city = form_row.get("New Delivery City / Province", "")

            pickup_num = 0
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
                        if flexible_address_match(addr1, addr2, ups_row["Address Line 1"], ups_row["Address Line 2"]):
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
                all_pickup_matched = True
                for pu_addr in pickup_addrs:
                    match = check_address_in_ups(pu_addr[0], pu_addr[1], "02")
                    if match is None:
                        unmatched_dict = form_row.to_dict()
                        unmatched_dict['Unmatched Reason'] = f"Pickup address not matched: {pu_addr[0]}, {pu_addr[1]}"
                        unmatched_rows.append(unmatched_dict)
                        all_pickup_matched = False
                        break
                    else:
                        pickup_matches.append(match)

                if all_pickup_matched and billing_match and delivery_match:
                    processed_form_indices.add(idx)
                    matched_dict = form_row.to_dict()

                    # Remove tones in matched
                    matched_dict["New Billing Address Line 1 (Tone-free)"] = normalize_address(billing_addr1)
                    matched_dict["New Billing Address Line 2 (Tone-free)"] = normalize_address(billing_addr2)
                    matched_dict["New Billing Address Line 3 (Tone-free)"] = normalize_address(billing_addr3)
                    matched_dict["New Delivery Address Line 1 (Tone-free)"] = normalize_address(delivery_addr1)
                    matched_dict["New Delivery Address Line 2 (Tone-free)"] = normalize_address(delivery_addr2)
                    matched_dict["New Delivery Address Line 3 (Tone-free)"] = normalize_address(delivery_addr3)
                    for i, pu_addr in enumerate(pickup_addrs, 1):
                        matched_dict[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 1 (Tone-free)"] = normalize_address(pu_addr[0])
                        matched_dict[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 2 (Tone-free)"] = normalize_address(pu_addr[1])
                        matched_dict[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 3 (Tone-free)"] = normalize_address(pu_addr[2])
                    matched_rows.append(matched_dict)

                    # Pickup addresses (type 02), each a separate row, invoice option empty
                    for pu_addr in pickup_addrs:
                        upload_template_rows.append({
                            "AC_NUM": form_row["Account Number"],
                            "AC_Address_Type": "02",
                            "invoice option": "",
                            "AC_Name": ups_acc_df["AC_Name"].values[0],
                            "Address_Line1": normalize_address(pu_addr[0]),
                            "Address_Line2": normalize_address(pu_addr[1]),
                            "City": pu_addr[3],
                            "Postal_Code": ups_acc_df["Postal_Code"].values[0],
                            "Country_Code": ups_acc_df["Country_Code"].values[0],
                            "Attention_Name": form_row.get("Full Name of Contact-In English Only", ""),
                            "Address_Line22": normalize_address(pu_addr[2]),
                            "Address_Country_Code": ups_acc_df["Address_Country_Code"].values[0]
                        })

                    # Billing address (type 03), 3 rows with invoice option = code (1, 2, 6)
                    for code in ["1", "2", "6"]:
                        upload_template_rows.append({
                            "AC_NUM": form_row["Account Number"],
                            "AC_Address_Type": code,
                            "invoice option": code,
                            "AC_Name": ups_acc_df["AC_Name"].values[0],
                            "Address_Line1": normalize_address(billing_addr1),
                            "Address_Line2": normalize_address(billing_addr2),
                            "City": billing_city,
                            "Postal_Code": ups_acc_df["Postal_Code"].values[0],
                            "Country_Code": ups_acc_df["Country_Code"].values[0],
                            "Attention_Name": form_row.get("Full Name of Contact-In English Only", ""),
                            "Address_Line22": normalize_address(billing_addr3),
                            "Address_Country_Code": ups_acc_df["Address_Country_Code"].values[0]
                        })

                    # Delivery address (type 13), single row with empty invoice option
                    upload_template_rows.append({
                        "AC_NUM": form_row["Account Number"],
                        "AC_Address_Type": "13",
                        "invoice option": "",
                        "AC_Name": ups_acc_df["AC_Name"].values[0],
                        "Address_Line1": normalize_address(delivery_addr1),
                        "Address_Line2": normalize_address(delivery_addr2),
                        "City": delivery_city,
                        "Postal_Code": ups_acc_df["Postal_Code"].values[0],
                        "Country_Code": ups_acc_df["Country_Code"].values[0],
                        "Attention_Name": form_row.get("Full Name of Contact-In English Only", ""),
                        "Address_Line22": normalize_address(delivery_addr3),
                        "Address_Country_Code": ups_acc_df["Address_Country_Code"].values[0]
                    })
                else:
                    # If any part unmatched
                    unmatched_dict = form_row.to_dict()
                    reasons = []
                    if not billing_match:
                        reasons.append("Billing address not matched")
                    if not delivery_match:
                        reasons.append("Delivery address not matched")
                    if not all_pickup_matched:
                        reasons.append("One or more pickup addresses not matched")
                    unmatched_dict['Unmatched Reason'] = "; ".join(reasons)
                    unmatched_rows.append(unmatched_dict)

    # Add unmatched forms rows not processed for any reason
    unmatched_not_processed = forms_df.loc[~forms_df.index.isin(processed_form_indices)]
    for _, row in unmatched_not_processed.iterrows():
        unmatched_dict = row.to_dict()
        if 'Unmatched Reason' not in unmatched_dict or unmatched_dict['Unmatched Reason'] == "":
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

            # Basic validation of required columns
            required_forms_cols = ["Account Number", "Is Your New Billing Address the Same as Your Pickup and Delivery Address?"]
            required_ups_cols = ["Account Number", "Address Type", "Address Line 1", "Address Line 2", "AC_Name", "Postal_Code", "Country_Code", "Address_Country_Code"]

            for col in required_forms_cols:
                if col not in forms_df.columns:
                    st.error(f"Missing column in Forms file: {col}")
                    return
            for col in required_ups_cols:
                if col not in ups_df.columns:
                    st.error(f"Missing column in UPS file: {col}")
                    return

            matched_df, unmatched_df, upload_template_df = process_files(forms_df, ups_df)

            st.success(f"Validation completed! Matched: {len(matched_df)}, Unmatched: {len(unmatched_df)}")

            def to_excel(df):
                output = BytesIO()
                writer = pd.ExcelWriter(output, engine='xlsxwriter')
                df.to_excel(writer, index=False, sheet_name='Sheet1')
                writer.save()
                processed_data = output.getvalue()
                return processed_data

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
                    label="Download Uploading Template",
                    data=to_excel(upload_template_df),
                    file_name="upload_template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()
